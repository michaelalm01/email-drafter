<#
.SYNOPSIS
Creates Outlook drafts from a CSV file and an Outlook template. Never sends email.
#>
[CmdletBinding()]
param(
    [string] $CsvPath,
    [string] $TemplatePath,
    [string] $AttachmentsDirectory,
    [string] $SettingsPath,
    [string] $LogPath,
    [switch] $DryRun,
    [ValidateRange(1, 100000)]
    [int] $LargeBatchThreshold = 25,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\EmailDraftSupport.ps1"
. "$PSScriptRoot\EmailDraftOutlook.ps1"

function Resolve-FullPath {
    param([string] $Path)
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path (Get-Location) $Path
    }
    return [System.IO.Path]::GetFullPath($Path)
}

function Get-RepositoryRoot {
    return (Split-Path -Parent $PSScriptRoot)
}

function Get-ResolvedRunPaths {
    $root = Get-RepositoryRoot
    $resolvedCsv = $CsvPath
    $resolvedTemplate = $TemplatePath
    $resolvedAttachments = $AttachmentsDirectory
    $resolvedSettings = $SettingsPath
    if ([string]::IsNullOrWhiteSpace($resolvedCsv)) { $resolvedCsv = Join-Path $root 'input\Recipients.csv' }
    if ([string]::IsNullOrWhiteSpace($resolvedTemplate)) { $resolvedTemplate = Join-Path $root 'templates\EmailTemplate.oft' }
    if ([string]::IsNullOrWhiteSpace($resolvedAttachments)) { $resolvedAttachments = Join-Path $root 'attachments' }
    if ([string]::IsNullOrWhiteSpace($resolvedSettings)) { $resolvedSettings = Join-Path $root 'input\settings.json' }
    return [pscustomobject]@{
        CsvPath = Resolve-FullPath -Path $resolvedCsv
        TemplatePath = Resolve-FullPath -Path $resolvedTemplate
        AttachmentsDirectory = Resolve-FullPath -Path $resolvedAttachments
        SettingsPath = Resolve-FullPath -Path $resolvedSettings
    }
}

function Resolve-RunLogPath {
    $resolved = $LogPath
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $resolved = Join-Path (Get-RepositoryRoot) 'logs\execution.log'
    }
    return (Resolve-FullPath -Path $resolved)
}

function Get-RelaunchArguments {
    $arguments = @('-LargeBatchThreshold', [string]$LargeBatchThreshold)
    if ($Force) { $arguments += '-Force' }
    if ($DryRun) { $arguments += '-DryRun' }
    foreach ($pair in @(
        @{ Name = 'CsvPath'; Value = $CsvPath },
        @{ Name = 'TemplatePath'; Value = $TemplatePath },
        @{ Name = 'AttachmentsDirectory'; Value = $AttachmentsDirectory },
        @{ Name = 'SettingsPath'; Value = $SettingsPath },
        @{ Name = 'LogPath'; Value = $LogPath }
    )) {
        if (-not [string]::IsNullOrWhiteSpace([string]$pair.Value)) {
            $arguments += "-$($pair.Name)"
            $arguments += [string]$pair.Value
        }
    }
    return $arguments
}

function Confirm-TemplateFile {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Outlook template was not found: $Path"
    }
}

function Confirm-LargeBatch {
    param([int] $Count)
    if ($Count -lt $LargeBatchThreshold -or $Force) { return }
    if (-not [Environment]::UserInteractive) {
        throw "Refusing to create $Count drafts without -Force."
    }
    $answer = Read-Host "Create $Count Outlook drafts? Type YES to continue"
    if ($answer -cne 'YES') { throw 'Draft creation was cancelled.' }
}

function Confirm-TemplateUnchanged {
    param([string] $TemplatePath, [string] $HashBefore, [string] $RunLogPath)
    $hashAfter = Get-FileContentHash -Path $TemplatePath
    if ($hashBefore -ne $hashAfter) {
        Write-DraftLog -LogPath $RunLogPath -Message 'The Outlook template file changed during processing.'
        throw 'The Outlook template file changed during processing.'
    }
    Write-DraftLog -LogPath $RunLogPath -Message 'Template file was not modified.'
}

function Write-ValidationLogs {
    param($Validated, [string] $RunLogPath)
    foreach ($warning in @($Validated.Warnings)) {
        Write-DraftLog -LogPath $RunLogPath -Message "Warning: $warning"
    }
    foreach ($invalid in @($Validated.Invalid)) {
        foreach ($errorMessage in @($invalid.Errors)) {
            Write-DraftLog -LogPath $RunLogPath -Message "Row $($invalid.RowNumber) skipped. $errorMessage"
        }
    }
}

function Save-OutlookDraftBatch {
    param($OutlookApplication, [string] $TemplatePath, $Validations, [string] $RunLogPath, $Settings)
    $created = 0
    $failed = 0
    foreach ($validation in @($Validations)) {
        try {
            Save-RecipientOutlookDraft -OutlookApplication $OutlookApplication -TemplatePath $TemplatePath -Validation $validation -Settings $Settings
            $created++
            $storeName = $validation.StoreName
            $rowNumber = $validation.RowNumber
            Write-DraftLog -LogPath $RunLogPath -Message "Draft saved for store $storeName on row $rowNumber."
        }
        catch {
            $failed++
            $storeName = $validation.StoreName
            $rowNumber = $validation.RowNumber
            $failure = $_.Exception.Message
            Write-DraftLog -LogPath $RunLogPath -Message "Draft failed for store $storeName on row $rowNumber. $failure"
        }
    }
    return [pscustomobject]@{ Created = $created; Failed = $failed }
}

function New-ValidatedOutlookDrafts {
    param([string] $TemplatePath, $Validations, [string] $RunLogPath, $Settings)
    $outlookWasRunning = $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    $outlook = New-OutlookApplication
    try {
        return Save-OutlookDraftBatch -OutlookApplication $outlook -TemplatePath $TemplatePath -Validations $Validations -RunLogPath $RunLogPath -Settings $Settings
    }
    finally {
        Close-OutlookApplication -OutlookApplication $outlook -Quit:(-not $outlookWasRunning)
    }
}

function Get-DraftCounts {
    param($Validated, [string] $TemplatePath, [string] $RunLogPath, $Settings)
    if ($DryRun) {
        Write-DraftLog -LogPath $RunLogPath -Message 'Dry run completed. No drafts were created.'
        return [pscustomobject]@{ Created = 0; Failed = 0 }
    }
    if (@($Validated.Valid).Count -eq 0) {
        return [pscustomobject]@{ Created = 0; Failed = 0 }
    }
    Confirm-LargeBatch -Count @($Validated.Valid).Count
    return New-ValidatedOutlookDrafts -TemplatePath $TemplatePath -Validations $Validated.Valid -RunLogPath $RunLogPath -Settings $Settings
}

function Get-ValidatedDraftSettings {
    param([string] $Path, [string] $RunLogPath)
    $settings = Import-DraftSettings -Path $Path
    foreach ($errorMessage in @($settings.Errors)) {
        Write-DraftLog -LogPath $RunLogPath -Message $errorMessage
    }
    if (@($settings.Errors).Count -gt 0) { throw 'Settings file contains invalid values.' }
    $count = @($settings.CcAddresses).Count
    Write-DraftLog -LogPath $RunLogPath -Message "Settings loaded. CC addresses: $count."
    return $settings
}

function New-DraftRunResult {
    param($Validated, $DraftCounts, [string] $RunLogPath)
    $skipped = @($Validated.Invalid).Count + $DraftCounts.Failed
    $exitCode = 0
    if ($skipped -gt 0) { $exitCode = 1 }
    return [pscustomobject]@{
        Processed = @($Validated.Rows).Count
        Created = $DraftCounts.Created
        Skipped = $skipped
        Warnings = @($Validated.Warnings).Count
        DryRun = [bool]$DryRun
        LogPath = $RunLogPath
        ExitCode = $exitCode
    }
}

function Invoke-DraftRun {
    param([string] $RunLogPath)
    $paths = Get-ResolvedRunPaths
    Confirm-TemplateFile -Path $paths.TemplatePath
    $hashBefore = Get-FileContentHash -Path $paths.TemplatePath
    $validated = Get-ValidatedRecipientRows -CsvPath $paths.CsvPath -AttachmentsDirectory $paths.AttachmentsDirectory
    $settings = Get-ValidatedDraftSettings -Path $paths.SettingsPath -RunLogPath $RunLogPath
    Write-ValidationLogs -Validated $validated -RunLogPath $RunLogPath
    $draftCounts = Get-DraftCounts -Validated $validated -TemplatePath $paths.TemplatePath -RunLogPath $RunLogPath -Settings $settings
    Confirm-TemplateUnchanged -TemplatePath $paths.TemplatePath -HashBefore $hashBefore -RunLogPath $RunLogPath
    return New-DraftRunResult -Validated $validated -DraftCounts $draftCounts -RunLogPath $RunLogPath
}

function Write-DraftRunSummary {
    param($Result)
    $summary = 'Rows processed: {0}. Drafts created: {1}. Rows skipped: {2}. Warnings: {3}.' -f $Result.Processed, $Result.Created, $Result.Skipped, $Result.Warnings
    Write-DraftLog -LogPath $Result.LogPath -Message $summary
    if ($Result.DryRun) { Write-Host 'Dry run. No drafts were created.' }
    else { Write-Host 'Drafts saved to Outlook.' }
    Write-Host ("Rows processed : {0}" -f $Result.Processed)
    Write-Host ("Drafts created : {0}" -f $Result.Created)
    Write-Host ("Rows skipped   : {0}" -f $Result.Skipped)
    Write-Host ("Warnings       : {0}" -f $Result.Warnings)
    Write-Host ("Log            : {0}" -f $Result.LogPath)
}

function Start-EmailDraftRun {
    if (-not $DryRun) {
        Start-StaRelaunchIfNeeded -ScriptPath $PSCommandPath -Arguments (Get-RelaunchArguments)
    }
    $runLogPath = Resolve-RunLogPath
    Write-DraftLog -LogPath $runLogPath -Message 'Execution started.'
    try {
        $result = Invoke-DraftRun -RunLogPath $runLogPath
        Write-DraftRunSummary -Result $result
        return $result.ExitCode
    }
    catch {
        Write-DraftLog -LogPath $runLogPath -Message "Execution failed. $($_.Exception.Message)"
        Write-Host $_.Exception.Message
        return 1
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    exit (Start-EmailDraftRun)
}

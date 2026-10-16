Set-StrictMode -Version Latest

function ConvertTo-CommandLineArgument {
    param([string] $Value)
    if ($Value -notmatch '[\s"]') { return $Value }
    return '"' + ($Value.Replace('"', '\"')) + '"'
}

function Start-StaRelaunchIfNeeded {
    param(
        [Parameter(Mandatory = $true)][string] $ScriptPath,
        [string[]] $Arguments = @()
    )
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    $isWindowsPowerShell = $PSVersionTable.PSEdition -eq 'Desktop'
    if ($isWindowsPowerShell -and $apartment -eq [System.Threading.ApartmentState]::STA) { return }
    $hostPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $hostPath)) {
        throw 'Outlook automation requires Windows PowerShell because Classic Outlook needs an STA host.'
    }
    $launch = @('-NoProfile', '-Sta', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath) + $Arguments
    $commandLine = ($launch | ForEach-Object { ConvertTo-CommandLineArgument -Value ([string]$_) }) -join ' '
    $process = Start-Process -FilePath $hostPath -ArgumentList $commandLine -Wait -PassThru -NoNewWindow
    exit $process.ExitCode
}

function New-OutlookApplication {
    return New-Object -ComObject Outlook.Application
}

function Close-OutlookApplication {
    param($OutlookApplication, [switch] $Quit)
    if ($Quit) {
        try { $OutlookApplication.Quit() } catch { }
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($OutlookApplication)
}

function New-OutlookTemplateDraft {
    param($OutlookApplication, [string] $TemplatePath)
    $fullPath = [System.IO.Path]::GetFullPath($TemplatePath)
    return $OutlookApplication.CreateItemFromTemplate($fullPath)
}

function Set-OutlookDraftContent {
    param($Draft, $Row, $Settings)
    $values = Get-PlaceholderValues -Row $Row
    $values['Store Name'] = ([string]$Row.StoreName).Trim()
    $values['Settlement Date'] = [string]$Settings.SettlementDate
    $values['Sender Names'] = [string]$Settings.SenderNames
    $Draft.Subject = Expand-DraftPlaceholders -TemplateText ([string]$Draft.Subject) -Values $values
    $html = [string]$Draft.HTMLBody
    if (-not [string]::IsNullOrWhiteSpace($html)) {
        $Draft.HTMLBody = Expand-DraftPlaceholders -TemplateText $html -Values $values -EncodeHtml
        return
    }
    $Draft.Body = Expand-DraftPlaceholders -TemplateText ([string]$Draft.Body) -Values $values
}

function Set-OutlookDraftRecipients {
    param($Draft, $Row, [string[]] $CcAddresses)
    $addresses = @(Get-DelimitedList -Value ([string]$Row.Recipients))
    $Draft.To = $addresses -join '; '
    $Draft.CC = @($CcAddresses) -join '; '
    $Draft.BCC = ''
}

function Add-OutlookDraftAttachments {
    param($Draft, [string[]] $AttachmentPaths)
    foreach ($path in @($AttachmentPaths)) {
        [void]$Draft.Attachments.Add($path)
    }
}

function Save-OutlookDraft {
    param($Draft)
    $Draft.Save()
}

function Close-OutlookDraft {
    param($Draft, [int] $SaveMode)
    try { $Draft.Close($SaveMode) } catch { }
}

function Save-RecipientOutlookDraft {
    param($OutlookApplication, [string] $TemplatePath, $Validation, $Settings)
    $draft = $null
    $saved = $false
    try {
        $draft = New-OutlookTemplateDraft -OutlookApplication $OutlookApplication -TemplatePath $TemplatePath
        Set-OutlookDraftContent -Draft $draft -Row $Validation.Row -Settings $Settings
        Set-OutlookDraftRecipients -Draft $draft -Row $Validation.Row -CcAddresses $Settings.CcAddresses
        Add-OutlookDraftAttachments -Draft $draft -AttachmentPaths $Validation.AttachmentPaths
        Save-OutlookDraft -Draft $draft
        $saved = $true
    }
    finally {
        if ($null -ne $draft) {
            $saveMode = 1
            if ($saved) { $saveMode = 0 }
            Close-OutlookDraft -Draft $draft -SaveMode $saveMode
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($draft)
        }
    }
}

function Export-OutlookEmailTemplate {
    param(
        [Parameter(Mandatory = $true)][string] $OutputPath,
        [Parameter(Mandatory = $true)][string] $HtmlBody
    )
    $directory = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    $outlookWasRunning = $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    $outlook = New-OutlookApplication
    $mail = $null
    try {
        $mail = $outlook.CreateItem(0)
        $mail.Subject = 'Collection Reminder - {Store Name}'
        $mail.BodyFormat = 2
        $mail.HTMLBody = $HtmlBody
        $mail.SaveAs([System.IO.Path]::GetFullPath($OutputPath), 2)
        $mail.Close(1)
    }
    finally {
        if ($null -ne $mail) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)
        }
        Close-OutlookApplication -OutlookApplication $outlook -Quit:(-not $outlookWasRunning)
    }
}

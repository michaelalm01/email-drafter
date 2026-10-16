[CmdletBinding()]
param(
    [string] $OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\EmailDraftSupport.ps1"
. "$PSScriptRoot\EmailDraftOutlook.ps1"

$script:EmailTemplateHtml = @'
<!DOCTYPE html>
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222222; line-height: 1.4;">
<p>Dear Sir/Madam,</p>
<p>Please see attached Collection Reminder regarding your outstanding account.</p>
<p>We request your immediate attention and settlement on or before {Settlement Date}.</p>
<p>Thank You.</p>
<p>Kind Regards,<br>{Sender Names}</p>
</body>
</html>
'@

function Get-EmailTemplateOutputPath {
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { return $OutputPath }
    $root = Split-Path -Parent $PSScriptRoot
    return (Join-Path $root 'templates\EmailTemplate.oft')
}

if ($MyInvocation.InvocationName -ne '.') {
    $relaunchArguments = @()
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $relaunchArguments = @('-OutputPath', $OutputPath)
    }
    Start-StaRelaunchIfNeeded -ScriptPath $PSCommandPath -Arguments $relaunchArguments
    $templatePath = Get-EmailTemplateOutputPath
    Export-OutlookEmailTemplate -OutputPath $templatePath -HtmlBody $script:EmailTemplateHtml
    Write-Host "Outlook template saved: $templatePath"
}

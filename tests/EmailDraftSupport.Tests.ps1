Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\..\scripts\EmailDraftSupport.ps1"
. "$PSScriptRoot\..\scripts\New-SamplePdf.ps1"

$script:Failures = New-Object System.Collections.Generic.List[string]

function Assert-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw $Message }
}

function Invoke-Scenario {
    param([string] $Name, [scriptblock] $Action)
    try {
        & $Action
        Write-Host "PASS  $Name"
    }
    catch {
        $script:Failures.Add("$Name :: $($_.Exception.Message)")
        Write-Host "FAIL  $Name :: $($_.Exception.Message)"
    }
}

function New-RecipientRow {
    param([hashtable] $Values, [int] $RowNumber = 2)
    $row = [pscustomobject]$Values
    Add-Member -InputObject $row -NotePropertyName 'RowNumber' -NotePropertyValue $RowNumber -Force
    return $row
}

Invoke-Scenario -Name 'Replaces supported placeholders' -Action {
    $values = @{ 'Store Name' = 'Northwind Traders, Inc.' }
    $text = 'Collection Reminder - {Store Name}'
    $expanded = Expand-DraftPlaceholders -TemplateText $text -Values $values
    Assert-Condition ($expanded -eq 'Collection Reminder - Northwind Traders, Inc.') 'Placeholder text was not replaced.'
}

Invoke-Scenario -Name 'Encodes HTML replacement values' -Action {
    $values = @{ 'Store Name' = 'Tom & Jerry' }
    $expanded = Expand-DraftPlaceholders -TemplateText 'Dear {Store Name}' -Values $values -EncodeHtml
    Assert-Condition ($expanded -eq 'Dear Tom &amp; Jerry') 'HTML encoding was not applied.'
}

Invoke-Scenario -Name 'Keeps unknown placeholders' -Action {
    $expanded = Expand-DraftPlaceholders -TemplateText '{Unknown}' -Values @{}
    Assert-Condition ($expanded -eq '{Unknown}') 'Unknown placeholder was removed.'
}

Invoke-Scenario -Name 'Rejects an invalid email address' -Action {
    $valid = Test-RecipientEmailAddress -EmailAddress 'alex.morgan@example.com'
    $invalid = Test-RecipientEmailAddress -EmailAddress 'not-an-email'
    Assert-Condition ($valid -and -not $invalid) 'Email validation returned an unexpected result.'
}

Invoke-Scenario -Name 'Rejects a path outside the attachments directory' -Action {
    $directory = Join-Path ([System.IO.Path]::GetTempPath()) ('email-drafter-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $directory | Out-Null
    try {
        $threw = $false
        try { Resolve-SafeAttachmentPath -AttachmentsDirectory $directory -FileName '..\secret.pdf' }
        catch { $threw = $true }
        Assert-Condition $threw 'Path traversal was accepted.'
    }
    finally {
        Remove-Item -LiteralPath $directory -Recurse -Force
    }
}

Invoke-Scenario -Name 'Reads draft settings and rejects an invalid CC address' -Action {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ('email-drafter-settings-' + [guid]::NewGuid().ToString('N') + '.json')
    try {
        @{
            cc = @('audit@example.com', 'not-an-email')
            senderNames = @('Jordan Lee', 'Sam Patel')
            settlementDate = '31 October 2026'
        } | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding utf8
        $loaded = Import-DraftSettings -Path $path
        Assert-Condition (@($loaded.CcAddresses).Count -eq 1) 'Expected one CC address.'
        Assert-Condition (@($loaded.Errors).Count -eq 1) 'Expected one invalid CC address.'
        Assert-Condition ($loaded.SettlementDate -eq '31 October 2026') 'Settlement date was not read.'
        Assert-Condition ($loaded.SenderNames -eq 'Jordan Lee / Sam Patel') 'Sender names were not read.'
    }
    finally {
        Remove-Item -LiteralPath $path -Force
    }
}

Invoke-Scenario -Name 'Requires both PDF attachments' -Action {
    $directory = Join-Path ([System.IO.Path]::GetTempPath()) ('email-drafter-' + [guid]::NewGuid().ToString('N'))
    New-SampleAttachmentFiles -Directory $directory
    try {
        $csvPath = Join-Path $directory 'recipients.csv'
        @(
            'StoreName,Recipients',
            '"Northwind Traders, Inc.",alex.morgan@example.com',
            'Missing Store,riley.chen@example.com',
            '"Northwind Traders, Inc.",billing@example.com'
        ) | Set-Content -LiteralPath $csvPath -Encoding utf8
        $validated = Get-ValidatedRecipientRows -CsvPath $csvPath -AttachmentsDirectory $directory
        Assert-Condition (@($validated.Valid).Count -eq 2) 'Expected two valid rows.'
        Assert-Condition (@($validated.Invalid).Count -eq 1) 'Expected one invalid row.'
        Assert-Condition (@($validated.Warnings).Count -eq 1) 'Expected one duplicate warning.'
        $invalidText = @($validated.Invalid)[0].Errors -join ' '
        Assert-Condition ($invalidText -match 'store name') 'Missing attachment was not reported.'
    }
    finally {
        Remove-Item -LiteralPath $directory -Recurse -Force
    }
}

Invoke-Scenario -Name 'Generated PDF has a valid xref offset' -Action {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ('email-drafter-' + [guid]::NewGuid().ToString('N') + '.pdf')
    try {
        New-SimplePdfDocument -Path $path -Lines @('DEMAND LETTER', 'Test data only')
        Assert-Condition (Test-PdfFileSignature -Path $path) 'Generated file is not a PDF.'
        $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($path))
        $marker = "startxref`n"
        $markerIndex = $text.LastIndexOf($marker)
        $offsetText = $text.Substring($markerIndex + $marker.Length).Split("`n")[0]
        $offset = [int]$offsetText
        Assert-Condition ($text.Substring($offset, 4) -eq 'xref') 'startxref does not point at xref.'
    }
    finally {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}

Invoke-Scenario -Name 'Script does not send email' -Action {
    $root = Split-Path -Parent $PSScriptRoot
    $scriptNames = @(
        'scripts\Create-EmailDrafts.ps1',
        'scripts\EmailDraftOutlook.ps1',
        'scripts\EmailDraftSupport.ps1',
        'scripts\New-EmailTemplate.ps1'
    )
    $pattern = '(?i)(\.Send\s*\(|Send-MailMessage|SubmitMessage)'
    foreach ($scriptName in $scriptNames) {
        $content = Get-Content -LiteralPath (Join-Path $root $scriptName) -Raw
        Assert-Condition ($content -notmatch $pattern) "$scriptName contains a send operation."
    }
}

if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) test(s) failed."
    exit 1
}

Write-Host 'All tests passed.'
exit 0

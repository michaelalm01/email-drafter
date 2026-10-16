Set-StrictMode -Version Latest

function Get-RequiredRecipientColumns {
    return @(
        'StoreName',
        'Recipients'
    )
}

function Test-RecipientEmailAddress {
    param([AllowEmptyString()][string] $EmailAddress)
    $candidate = ([string]$EmailAddress).Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) { return $false }
    return $candidate -match '^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$'
}

function Get-DelimitedList {
    param([AllowNull()][string] $Value, [string] $Separators = ';,')
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    $parts = $Value.Split($Separators.ToCharArray(), [System.StringSplitOptions]::RemoveEmptyEntries)
    return @($parts | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-CsvColumnNames {
    param([Parameter(Mandatory = $true)][string] $Path)
    $header = @(Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
    if ($header.Count -eq 0 -or [string]::IsNullOrWhiteSpace($header[0])) {
        throw 'CSV file is empty.'
    }
    $parsed = ConvertFrom-Csv -InputObject ($header[0] + [Environment]::NewLine + 'x')
    return @($parsed.PSObject.Properties.Name)
}

function Confirm-RequiredCsvColumns {
    param([Parameter(Mandatory = $true)][string[]] $ColumnNames)
    $missing = @(Get-RequiredRecipientColumns | Where-Object { $ColumnNames -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw "CSV is missing required columns: $($missing -join ', ')"
    }
}

function Import-RecipientCsv {
    param([Parameter(Mandatory = $true)][string] $Path)
    $rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8)
    $numberedRows = New-Object System.Collections.ArrayList
    $rowNumber = 2
    foreach ($row in $rows) {
        Add-Member -InputObject $row -NotePropertyName 'RowNumber' -NotePropertyValue $rowNumber -Force
        [void]$numberedRows.Add($row)
        $rowNumber++
    }
    return @($numberedRows)
}

function Get-AttachmentRootPath {
    param([Parameter(Mandatory = $true)][string] $AttachmentsDirectory)
    if (-not (Test-Path -LiteralPath $AttachmentsDirectory -PathType Container)) {
        throw "Attachments directory was not found: $AttachmentsDirectory"
    }
    $root = [System.IO.Path]::GetFullPath($AttachmentsDirectory)
    $separator = [System.IO.Path]::DirectorySeparatorChar
    if (-not $root.EndsWith($separator)) { $root += $separator }
    return $root
}

function Assert-AttachmentFileName {
    param([string] $FileName)
    if ([string]::IsNullOrWhiteSpace($FileName)) { throw 'Attachment file name is empty.' }
    $extension = [System.IO.Path]::GetExtension($FileName)
    $hasSeparator = $FileName.IndexOfAny([char[]]@('\', '/', ':')) -ge 0
    $extensionIsPdf = $extension.Equals('.pdf', [System.StringComparison]::OrdinalIgnoreCase)
    $isRelativeSegment = $FileName -eq '.' -or $FileName -eq '..'
    if ($hasSeparator -or $isRelativeSegment -or -not $extensionIsPdf) {
        throw "Attachment file name is not allowed: $FileName"
    }
}

function Test-PdfFileSignature {
    param([Parameter(Mandatory = $true)][string] $Path)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $buffer = New-Object byte[] 5
        $read = $stream.Read($buffer, 0, 5)
        if ($read -lt 5) { return $false }
        return [System.Text.Encoding]::ASCII.GetString($buffer).StartsWith('%PDF-')
    }
    finally {
        $stream.Dispose()
    }
}

function Resolve-SafeAttachmentPath {
    param(
        [Parameter(Mandatory = $true)][string] $AttachmentsDirectory,
        [AllowEmptyString()][string] $FileName
    )
    $trimmed = ([string]$FileName).Trim()
    Assert-AttachmentFileName -FileName $trimmed
    $root = Get-AttachmentRootPath -AttachmentsDirectory $AttachmentsDirectory
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $root $trimmed))
    if (-not $fullPath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Attachment path is outside the attachments directory: $trimmed"
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Attachment file was not found: $trimmed"
    }
    if (-not (Test-PdfFileSignature -Path $fullPath)) {
        throw "Attachment is not a valid PDF: $trimmed"
    }
    return $fullPath
}

function Add-RequiredTextError {
    param($Errors, $Row, [string] $ColumnName)
    if ([string]::IsNullOrWhiteSpace([string]$Row.$ColumnName)) {
        $Errors.Add("$ColumnName is required.")
    }
}

function Get-JsonPropertyValue {
    param($Settings, [string] $PropertyName)
    $property = $Settings.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-JsonStringList {
    param($Value)
    $items = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Value) { return @() }
    foreach ($item in @($Value)) {
        $text = ([string]$item).Trim()
        if (-not [string]::IsNullOrWhiteSpace($text)) { $items.Add($text) }
    }
    return @($items)
}

function Add-CcSettingErrors {
    param($Errors, [string[]] $Addresses)
    $valid = New-Object System.Collections.Generic.List[string]
    foreach ($address in @($Addresses)) {
        if (Test-RecipientEmailAddress -EmailAddress $address) { $valid.Add($address) }
        else { $Errors.Add('Settings file cc contains an email address that is not valid.') }
    }
    return @($valid)
}

function Import-DraftSettings {
    param([Parameter(Mandatory = $true)][string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Settings file was not found: $Path"
    }
    $settings = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $errors = New-Object System.Collections.Generic.List[string]
    $settlementDate = ([string](Get-JsonPropertyValue -Settings $settings -PropertyName 'settlementDate')).Trim()
    $senderNames = @(Get-JsonStringList -Value (Get-JsonPropertyValue -Settings $settings -PropertyName 'senderNames'))
    if ([string]::IsNullOrWhiteSpace($settlementDate)) { $errors.Add('Settings file settlementDate is required.') }
    if ($senderNames.Count -eq 0) { $errors.Add('Settings file senderNames is required.') }
    $ccAddresses = @(Get-JsonStringList -Value (Get-JsonPropertyValue -Settings $settings -PropertyName 'cc'))
    $validCc = @(Add-CcSettingErrors -Errors $errors -Addresses $ccAddresses)
    return [pscustomobject]@{
        CcAddresses = $validCc
        SenderNames = ($senderNames -join ' / ')
        SettlementDate = $settlementDate
        Errors = @($errors)
    }
}

function Add-RecipientErrors {
    param($Errors, $Row)
    $addresses = @(Get-DelimitedList -Value ([string]$Row.Recipients))
    if ($addresses.Count -eq 0) {
        $Errors.Add('Recipients must contain at least one email address.')
        return
    }
    foreach ($address in $addresses) {
        if (-not (Test-RecipientEmailAddress -EmailAddress $address)) {
            $Errors.Add('Recipients contains an email address that is not valid.')
        }
    }
}

function Find-StoreAttachmentPaths {
    param(
        [Parameter(Mandatory = $true)][string] $AttachmentsDirectory,
        [AllowEmptyString()][string] $StoreName
    )
    $trimmedStoreName = ([string]$StoreName).Trim()
    $paths = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($trimmedStoreName)) { return @() }
    $root = Get-AttachmentRootPath -AttachmentsDirectory $AttachmentsDirectory
    foreach ($file in @(Get-ChildItem -LiteralPath $root -File -Filter '*.pdf')) {
        $nameContainsStore = $file.Name.IndexOf($trimmedStoreName, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        if ($nameContainsStore -and (Test-PdfFileSignature -Path $file.FullName)) {
            [void]$paths.Add($file.FullName)
        }
    }
    return @($paths)
}

function Get-RecipientRowValidation {
    param($Row, [Parameter(Mandatory = $true)][string] $AttachmentsDirectory)
    $errors = New-Object System.Collections.Generic.List[string]
    Add-RequiredTextError -Errors $errors -Row $Row -ColumnName 'StoreName'
    Add-RecipientErrors -Errors $errors -Row $Row
    $paths = @(Find-StoreAttachmentPaths -AttachmentsDirectory $AttachmentsDirectory -StoreName ([string]$Row.StoreName))
    if ($paths.Count -eq 0) {
        $errors.Add('No PDF file name contains the store name.')
    }
    return [pscustomobject]@{
        IsValid = ($errors.Count -eq 0)
        RowNumber = $Row.RowNumber
        StoreName = ([string]$Row.StoreName).Trim()
        Errors = @($errors)
        AttachmentPaths = $paths
        Row = $Row
    }
}

function Get-ValidatedRecipientRows {
    param(
        [Parameter(Mandatory = $true)][string] $CsvPath,
        [Parameter(Mandatory = $true)][string] $AttachmentsDirectory
    )
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "CSV file was not found: $CsvPath"
    }
    Confirm-RequiredCsvColumns -ColumnNames (Get-CsvColumnNames -Path $CsvPath)
    $rows = @(Import-RecipientCsv -Path $CsvPath)
    $valid = New-Object System.Collections.ArrayList
    $invalid = New-Object System.Collections.ArrayList
    foreach ($row in $rows) {
        $validation = Get-RecipientRowValidation -Row $row -AttachmentsDirectory $AttachmentsDirectory
        if ($validation.IsValid) { [void]$valid.Add($validation) } else { [void]$invalid.Add($validation) }
    }
    return [pscustomobject]@{
        Rows = $rows
        Valid = $valid
        Invalid = $invalid
        Warnings = @(Get-DuplicateRecipientWarnings -Rows $rows)
    }
}

function Get-StoreRowKey {
    param($Row)
    return ([string]$Row.StoreName).Trim().ToLowerInvariant()
}

function Get-DuplicateRecipientWarnings {
    param($Rows)
    $warnings = New-Object System.Collections.Generic.List[string]
    $groups = @($Rows | Group-Object -Property { Get-StoreRowKey $_ })
    foreach ($group in $groups) {
        if ($group.Count -lt 2) { continue }
        $rowNumbers = @($group.Group | ForEach-Object { $_.RowNumber })
        $storeName = ([string]$group.Group[0].StoreName).Trim()
        if ([string]::IsNullOrWhiteSpace($storeName)) { $storeName = '(blank)' }
        $warnings.Add("Rows $($rowNumbers -join ', ') repeat the store $storeName.")
    }
    return @($warnings)
}

function Get-PlaceholderReplacement {
    param([string] $Key, [hashtable] $Values, [bool] $EncodeHtml)
    if ($Key -notmatch '^[A-Za-z0-9_ ]+$' -or -not $Values.ContainsKey($Key)) { return $null }
    $replacement = [string]$Values[$Key]
    if (-not $EncodeHtml) { return $replacement }
    $encoded = [System.Net.WebUtility]::HtmlEncode($replacement)
    return $encoded.Replace("`n", '<br>')
}

function Expand-DraftPlaceholders {
    param(
        [AllowNull()][string] $TemplateText,
        [hashtable] $Values,
        [switch] $EncodeHtml
    )
    if ([string]::IsNullOrEmpty($TemplateText)) { return '' }
    $result = New-Object System.Text.StringBuilder
    $index = 0
    while ($index -lt $TemplateText.Length) {
        $token = Read-NextPlaceholderToken -TemplateText $TemplateText -StartIndex $index
        [void]$result.Append($token.LiteralText)
        if ($token.IsEnd) { break }
        $replacement = Get-PlaceholderReplacement -Key $token.Key -Values $Values -EncodeHtml $EncodeHtml.IsPresent
        if ($null -eq $replacement) { [void]$result.Append($token.RawToken) }
        else { [void]$result.Append($replacement) }
        $index = $token.NextIndex
    }
    return $result.ToString()
}

function Read-NextPlaceholderToken {
    param([string] $TemplateText, [int] $StartIndex)
    $start = $TemplateText.IndexOf('{', $StartIndex)
    if ($start -lt 0 -or ($end = $TemplateText.IndexOf('}', $start + 1)) -lt 0) {
        return [pscustomobject]@{
            LiteralText = $TemplateText.Substring($StartIndex)
            RawToken = ''
            Key = ''
            NextIndex = $TemplateText.Length
            IsEnd = $true
        }
    }
    return [pscustomobject]@{
        LiteralText = $TemplateText.Substring($StartIndex, $start - $StartIndex)
        RawToken = $TemplateText.Substring($start, ($end + 1) - $start)
        Key = $TemplateText.Substring($start + 1, $end - ($start + 1))
        NextIndex = $end + 1
        IsEnd = $false
    }
}

function Get-PlaceholderValues {
    param($Row)
    $values = @{}
    foreach ($property in $Row.PSObject.Properties) {
        if ($property.Name -eq 'RowNumber') { continue }
        $values[$property.Name] = ([string]$property.Value).Trim()
    }
    return $values
}

function Write-DraftLog {
    param(
        [Parameter(Mandatory = $true)][string] $LogPath,
        [Parameter(Mandatory = $true)][string] $Message
    )
    $directory = Split-Path -Parent $LogPath
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }
    $entry = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Message
    Add-Content -LiteralPath $LogPath -Value $entry -Encoding utf8
}

function Get-FileContentHash {
    param([Parameter(Mandatory = $true)][string] $Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $hash = $sha.ComputeHash($stream)
        return ([System.BitConverter]::ToString($hash)).Replace('-', '')
    }
    finally {
        $stream.Dispose()
        $sha.Dispose()
    }
}

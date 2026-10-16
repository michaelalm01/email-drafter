Set-StrictMode -Version Latest

function ConvertTo-PdfTextLine {
    param([string] $Line)
    foreach ($character in $Line.ToCharArray()) {
        if ([int]$character -gt 127) {
            throw 'PDF sample text must use ASCII characters only.'
        }
    }
    return $Line.Replace('\', '\\').Replace('(', '\(').Replace(')', '\)')
}

function Get-PdfContentStream {
    param([string[]] $Lines)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append("BT`n/F1 16 Tf`n72 740 Td`n")
    $isFirstLine = $true
    foreach ($line in $Lines) {
        if (-not $isFirstLine) { [void]$builder.Append("0 -24 Td`n") }
        $escaped = ConvertTo-PdfTextLine -Line $line
        [void]$builder.Append("($escaped) Tj`n")
        $isFirstLine = $false
    }
    [void]$builder.Append("ET`n")
    return $builder.ToString()
}

function New-PdfObjectList {
    param([string] $ContentStream)
    $length = [System.Text.Encoding]::ASCII.GetByteCount($ContentStream)
    $streamObject = "4 0 obj`n<< /Length $length >>`nstream`n" + $ContentStream + "endstream`nendobj`n"
    return @(
        "1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj`n",
        "2 0 obj`n<< /Type /Pages /Kids [3 0 R] /Count 1 >>`nendobj`n",
        "3 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>`nendobj`n",
        $streamObject,
        "5 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>`nendobj`n"
    )
}

function Add-AsciiBytes {
    param($Buffer, [string] $Text)
    $Buffer.AddRange([System.Text.Encoding]::ASCII.GetBytes($Text))
}

function Get-PdfTrailer {
    param($Offsets, [int] $XrefOffset)
    $builder = New-Object System.Text.StringBuilder
    $size = $Offsets.Count + 1
    [void]$builder.Append("xref`n0 $size`n")
    [void]$builder.Append("0000000000 65535 f `n")
    foreach ($offset in $Offsets) {
        [void]$builder.Append(("{0:D10} 00000 n `n" -f $offset))
    }
    [void]$builder.Append("trailer`n<< /Size $size /Root 1 0 R >>`n")
    [void]$builder.Append("startxref`n$XrefOffset`n%%EOF`n")
    return $builder.ToString()
}

function New-SimplePdfDocument {
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][string[]] $Lines
    )
    $objects = New-PdfObjectList -ContentStream (Get-PdfContentStream -Lines $Lines)
    $buffer = New-Object System.Collections.Generic.List[byte]
    Add-AsciiBytes -Buffer $buffer -Text "%PDF-1.4`n"
    $offsets = New-Object System.Collections.Generic.List[int]
    foreach ($object in $objects) {
        $offsets.Add($buffer.Count)
        Add-AsciiBytes -Buffer $buffer -Text $object
    }
    $xrefOffset = $buffer.Count
    Add-AsciiBytes -Buffer $buffer -Text (Get-PdfTrailer -Offsets $offsets -XrefOffset $xrefOffset)
    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }
    [System.IO.File]::WriteAllBytes($Path, $buffer.ToArray())
}

function New-SampleAttachmentFiles {
    param([Parameter(Mandatory = $true)][string] $Directory)
    foreach ($storeName in @('Northwind Traders, Inc.', 'Contoso Supplies')) {
        $fileName = "Collection Reminder  - $storeName.pdf"
        New-SimplePdfDocument -Path (Join-Path $Directory $fileName) -Lines @(
            'COLLECTION REMINDER',
            'Test data only',
            "Store: $storeName",
            'Sample collection reminder for draft automation.',
            'This document is not addressed to a real customer.'
        )
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $repositoryRoot = Split-Path -Parent $PSScriptRoot
    New-SampleAttachmentFiles -Directory (Join-Path $repositoryRoot 'attachments')
}

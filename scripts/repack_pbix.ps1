# Repack a PBIR-format PBIX source tree into a .pbix that Power BI Desktop opens cleanly.
# Strips SecurityBindings + its Content_Types entry (PBI regenerates it on Save).
# Preserves DataModel compression. Sets ExternalAttributes=0 (the single most
# common cause of "corrupted" errors when repacking with Python zipfile).
#
# Usage:  .\repack_pbix.ps1 -Src .\pbix_src -Dst .\Report_patched.pbix

param(
    [Parameter(Mandatory=$true)][string]$Src,
    [Parameter(Mandatory=$true)][string]$Dst
)

Add-Type -AssemblyName System.IO.Compression

if (Test-Path $Dst) { Remove-Item $Dst -Force }

$srcFull = (Resolve-Path $Src).Path
$fs = [System.IO.File]::Open($Dst, 'Create')
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

try {
    Get-ChildItem $srcFull -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($srcFull.Length + 1).Replace('\','/')

        if ($rel -eq 'SecurityBindings') { return }

        if ($rel -eq '[Content_Types].xml') {
            $xml = [System.IO.File]::ReadAllText($_.FullName)
            $xml = $xml -replace '<Override[^/]*PartName="/SecurityBindings"[^/]*/>',''
            $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
            $entry.ExternalAttributes = 0
            $w = New-Object System.IO.StreamWriter($entry.Open())
            $w.Write($xml); $w.Close()
            return
        }

        $level = if ($rel -eq 'DataModel') {
            [System.IO.Compression.CompressionLevel]::NoCompression
        } else {
            [System.IO.Compression.CompressionLevel]::Optimal
        }

        $entry = $zip.CreateEntry($rel, $level)
        $entry.ExternalAttributes = 0
        $out = $entry.Open()
        $in  = [System.IO.File]::OpenRead($_.FullName)
        $in.CopyTo($out)
        $in.Close(); $out.Close()
    }
} finally {
    $zip.Dispose(); $fs.Close()
}

Write-Host "Wrote $Dst"

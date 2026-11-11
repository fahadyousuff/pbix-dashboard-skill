---
name: pbix-dashboard
description: Programmatically add or edit visuals, pages, and visual properties inside a Power BI Desktop .pbix file (PBIR Enhanced Report Format) without corrupting it. Use when the user asks to build, modify, generate, or automate a Power BI dashboard directly in a saved .pbix file — triggers include "edit pbix", "add visuals to Power BI", "build a Power BI dashboard programmatically", "modify Report/definition/pages", "patch a pbix", "generate Power BI report from a dataset", or anywhere a .pbix file is the input and output. Also use when the user complains that a hand-edited .pbix opens as "corrupted" (Mashup / SecurityBindings errors) — this skill explains the root cause and the fix.
---

# PBIX Dashboard Patcher

Edit a Power BI Desktop `.pbix` file in place: add pages, add visuals, change layout, rebind fields — then repack it so Power BI Desktop opens it without a corruption error.

## When to use
- User has a saved `.pbix` and wants visuals added / a dashboard generated.
- User wants to automate Power BI report edits (CI, bulk templating, data-driven layout).
- Earlier edit attempt produced `MashupValidationError`, `Invalid SecurityBindings`, or "corrupted file" on open.

## The core problem (read this first)

Modern `.pbix` files use the **PBIR Enhanced Report Format**: the report layout is stored as per-visual JSON files under `Report/definition/pages/<pageId>/visuals/<visualId>/visual.json`, not the old monolithic `Report/Layout`.

A PBIX is a ZIP containing (among others):
- `DataModel` — the Analysis Services model (must NOT be recompressed; store as-is).
- `Report/definition/...` — PBIR per-visual JSON tree.
- `SecurityBindings` — a DPAPI-encrypted manifest signing the file layout.
- `[Content_Types].xml` — lists every file in the ZIP.

**Three gotchas that corrupt the file:**

1. **Python `zipfile` sets external file attributes that differ from .NET's `ZipArchive`.** Power BI validates them. Use .NET `ZipArchive` via PowerShell, with `ExternalAttributes = 0` on every entry.
2. **Recompressing `DataModel` breaks it.** It must be stored with the **same compression method** as the source (usually `Deflate` but sometimes `Stored`); safest is `NoCompression` if you rewrote it, otherwise preserve the original bytes verbatim.
3. **Adding any file breaks `SecurityBindings`.** It's a signed manifest of what's in the ZIP. If you add `Report/definition/pages/<new>/visuals/.../visual.json`, the signature no longer matches and PBI refuses to open.

**The fix:** strip `SecurityBindings` from the ZIP **and** remove its entry from `[Content_Types].xml`. Power BI Desktop regenerates `SecurityBindings` on the next Save. The file opens cleanly.

## Procedure

### 1. Inspect the file
```powershell
# Extract to a working directory using .NET so attributes come back clean
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("Test.pbix", "pbix_src")
```

Check whether it's PBIR (new) or legacy:
- PBIR: has `Report\definition\pages\<pageId>\visuals\<visualId>\visual.json`.
- Legacy: has only `Report\Layout` (binary JSON string). This skill targets PBIR.

### 2. Understand the visual.json schema (version 2.12.0)

Minimal cardVisual (KPI tile):
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.12.0/schema.json",
  "name": "<20-hex-char id>",
  "position": {"x": 20, "y": 20, "z": 0, "height": 160, "width": 440, "tabOrder": 0},
  "visual": {
    "visualType": "cardVisual",
    "query": {
      "queryState": {
        "Data": {
          "projections": [{
            "field": {
              "Aggregation": {
                "Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "financials"}}, "Property": "Sales"}},
                "Function": 0
              }
            },
            "queryRef": "Sum(financials.Sales)",
            "nativeQueryRef": "Sum of Sales",
            "active": true
          }]
        }
      },
      "sortDefinition": {
        "sort": [{"field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "financials"}}, "Property": "Sales"}}, "Function": 0}}, "direction": "Descending"}],
        "isDefaultSort": true
      }
    },
    "drillFilterOtherVisuals": true
  }
}
```

**Key facts:**
- `name` is a 20-character hex id; must be unique within the page.
- `visualType` examples: `cardVisual`, `clusteredBarChart`, `clusteredColumnChart`, `lineChart`, `donutChart`, `pieChart`, `tableEx`, `pivotTable`.
- `cardVisual` uses role **`Data`**. Bar/column/line use `Category` + `Y` (+ optional `Series`). Donut/pie use `Category` + `Y`.
- `Aggregation.Function`: 0=Sum, 1=Avg, 2=Count, 3=Min, 4=Max, 5=CountNonNull, 6=Median, 7=StDev, 8=Variance.
- `drillFilterOtherVisuals: true` is required; omitting it corrupts the file.

### 3. Register the page

`Report/definition/pages/pages.json`:
```json
{"$schema":"...pagesMetadata/1.1.0/schema.json","pageOrder":["<pageId1>","<pageId2>"],"activePageName":"<pageIdX>"}
```

Each page folder has a `page.json`:
```json
{"$schema":"...page/2.1.0/schema.json","name":"<pageId>","displayName":"Dashboard","displayOption":"FitToPage","height":1080,"width":1920}
```

### 4. Repack — the ONLY reliable method

Do NOT use Python `zipfile`. Use PowerShell with .NET `ZipArchive`:

```powershell
Add-Type -AssemblyName System.IO.Compression

$src = "pbix_src"
$dst = "Test_patched.pbix"
if (Test-Path $dst) { Remove-Item $dst }

$fs = [System.IO.File]::Open($dst, 'Create')
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

Get-ChildItem $src -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring((Resolve-Path $src).Path.Length + 1).Replace('\','/')

    # STRIP SecurityBindings — PBI regenerates it on save
    if ($rel -eq 'SecurityBindings') { return }

    # Edit Content_Types.xml to drop the SecurityBindings reference
    if ($rel -eq '[Content_Types].xml') {
        $xml = [System.IO.File]::ReadAllText($_.FullName)
        $xml = $xml -replace '<Override[^/]*PartName="/SecurityBindings"[^/]*/>',''
        $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
        $entry.ExternalAttributes = 0
        $sw = New-Object System.IO.StreamWriter($entry.Open())
        $sw.Write($xml); $sw.Close()
        return
    }

    # DataModel must NOT be recompressed
    $level = if ($rel -eq 'DataModel') {
        [System.IO.Compression.CompressionLevel]::NoCompression
    } else {
        [System.IO.Compression.CompressionLevel]::Optimal
    }

    $entry = $zip.CreateEntry($rel, $level)
    $entry.ExternalAttributes = 0   # critical — Python sets this to a non-zero value PBI rejects
    $es = $entry.Open()
    $src_fs = [System.IO.File]::OpenRead($_.FullName)
    $src_fs.CopyTo($es)
    $src_fs.Close(); $es.Close()
}
$zip.Dispose(); $fs.Close()
```

### 5. Validate before opening
```powershell
# Confirm SecurityBindings is gone, DataModel is intact
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead("Test_patched.pbix")
$z.Entries | Where-Object { $_.FullName -match 'Security|DataModel' } | Select FullName, Length
$z.Dispose()
```
Expect: `DataModel` present, `SecurityBindings` absent.

## Connecting to a live Power BI Desktop session (optional — for data discovery)

If PBI Desktop is already running, it exposes an Analysis Services endpoint on localhost:

```powershell
# Discover the port
Get-Process msmdsrv | ForEach-Object {
    $port = Get-NetTCPConnection -OwningProcess $_.Id -State Listen |
            Where-Object { $_.LocalAddress -eq '127.0.0.1' } |
            Select-Object -First 1 -ExpandProperty LocalPort
    Write-Host "Analysis Services port: $port"
}
```

Then connect via ADOMD (install Power BI Desktop; DLL ships with it):
```python
# pip install pyadomd pythonnet
import clr
clr.AddReference(r"C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.PowerBI.AdomdClient.dll")
from Microsoft.AnalysisServices.AdomdClient import AdomdConnection  # real namespace, not the file name

conn = AdomdConnection(f"Data Source=localhost:{port}")
conn.Open()
# Run DAX: EVALUATE SUMMARIZECOLUMNS(financials[Segment], "Sales", SUM(financials[Sales]))
```

Use this only to **read schema and preview numbers**; final visuals are written as JSON files and repackaged above.

## Anti-patterns (things that will corrupt the file)
- Python `zipfile.ZipFile(...)` repack → `ExternalAttributes` mismatch → "corrupted".
- Recompressing `DataModel` → `MashupValidationError`.
- Leaving `SecurityBindings` when you added any new file → "SecurityBindings invalid".
- Guessing the `visual.json` schema → silent corruption. Always build real visuals in PBI once, inspect their JSON, then clone the shape.
- `pbi-tools` v1.2.0 does NOT round-trip PBIR (it extracts Model/TMDL but skips `Report/definition/pages/*`). Don't use it for this.

## Quick sanity test
Make a copy of a PBIX, repack it with the script above with NO changes. It must open cleanly. If it doesn't, the repack itself is wrong — fix that before editing content.

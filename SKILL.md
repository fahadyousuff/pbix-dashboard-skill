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

**Gotchas that corrupt the file (learn from scars):**

1. **Python `zipfile` sets external file attributes that differ from .NET's `ZipArchive`.** Power BI validates them. Use .NET `ZipArchive` via PowerShell, with `ExternalAttributes = 0` on every entry.
2. **Recompressing `DataModel` breaks it.** It must be stored with the **same compression method** as the source (usually `Deflate` but sometimes `Stored`); safest is `NoCompression` if you rewrote it, otherwise preserve the original bytes verbatim.
3. **Adding any file breaks `SecurityBindings`.** It's a signed manifest of what's in the ZIP. If you add `Report/definition/pages/<new>/visuals/.../visual.json`, the signature no longer matches and PBI refuses to open.
4. **Entry names must NOT have a leading `/`.** `MissingBIPackagePart: /Version` on open means your `Version` entry was zipped as `/Version` (leading slash). PowerShell's naive `Substring($src.Length + 1)` leaves a leading backslash → becomes `/` in the zip → PBI rejects. Compute the prefix from the resolved long path and `TrimStart('/')` every entry name.
5. **UTF-8 BOM in `Report/definition/**/*.json` is fatal.** PBI error: `Only text with UTF8 encoding without BOM (byte order marks) is supported. Detected BOM: 'UTF-8'`. `Set-Content -Encoding UTF8` and `File.WriteAllText` without an explicit encoding BOTH write a BOM in Windows PowerShell 5.1 / .NET Framework. Always write JSON with `New-Object System.Text.UTF8Encoding($false)` passed to `WriteAllText` or `StreamWriter`.

**The fix for SecurityBindings:** strip it from the ZIP **and** remove its entry from `[Content_Types].xml`. Power BI Desktop regenerates `SecurityBindings` on the next Save. The file opens cleanly.

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

### 4. Write JSON files — UTF-8 WITHOUT BOM

Every time you edit `visual.json`, `page.json`, `report.json`, or any file under `Report/definition/`, write it with no BOM. Windows PowerShell 5.1 adds a BOM by default.

```powershell
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
```

To strip a BOM already written by a previous mistake:
```powershell
$b = [System.IO.File]::ReadAllBytes($path)
if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) {
    $text = [System.Text.Encoding]::UTF8.GetString($b, 3, $b.Length - 3)
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
}
```

### 5. Repack — the ONLY reliable method

Do NOT use Python `zipfile`. Use PowerShell with .NET `ZipArchive`. Key details:
- Entry names use forward slashes and have NO leading `/`.
- Compute the source prefix length from `(Get-Item $src).FullName.TrimEnd('\').Length + 1` (same casing/long-path form as `FullName` on Get-ChildItem items). `TrimStart('/')` every entry name as insurance.
- Every text entry (including `[Content_Types].xml`) is written with a UTF-8 no-BOM `StreamWriter`.

```powershell
Add-Type -AssemblyName System.IO.Compression

$src = "pbix_src"
$dst = "Test_patched.pbix"
if (Test-Path $dst) { Remove-Item $dst }

$srcFull = (Get-Item $src).FullName.TrimEnd('\')
$prefixLen = $srcFull.Length + 1
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$fs = [System.IO.File]::Open($dst, 'Create')
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

Get-ChildItem $src -Recurse -File | ForEach-Object {
    $full = $_.FullName
    $rel = $full.Substring($prefixLen).Replace('\','/').TrimStart('/')

    # STRIP SecurityBindings — PBI regenerates it on save
    if ($rel -eq 'SecurityBindings') { return }

    # Edit Content_Types.xml to drop the SecurityBindings reference
    if ($rel -eq '[Content_Types].xml') {
        $xml = [System.IO.File]::ReadAllText($full) -replace '<Override[^/]*PartName="/SecurityBindings"[^/]*/>',''
        $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
        $entry.ExternalAttributes = 0
        $sw = New-Object System.IO.StreamWriter($entry.Open(), $utf8NoBom)
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
    $sfs = [System.IO.File]::OpenRead($full)
    $sfs.CopyTo($es)
    $sfs.Close(); $es.Close()
}
$zip.Dispose(); $fs.Close()
```

### 6. Validate before opening

Four checks, in order:

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead("Test_patched.pbix")

# a) Required parts present, SecurityBindings absent
$z.Entries | Where-Object { $_.FullName -match '^(Version|DataModel|Metadata|Settings|SecurityBindings|\[Content_Types\]\.xml)$' } |
    Select-Object FullName, Length

# b) No entry has a leading '/'
$z.Entries | Where-Object { $_.FullName.StartsWith('/') } | ForEach-Object { "BAD: $($_.FullName)" }

# c) No BOM on any text metadata or Report/definition JSON
foreach ($e in $z.Entries) {
    if ($e.FullName -match '\.(json|xml)$|^Version$|^Metadata$|^Settings$') {
        $s = $e.Open(); $ms = New-Object System.IO.MemoryStream; $s.CopyTo($ms); $s.Close()
        $b = $ms.ToArray()
        if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) {
            Write-Host "BOM found: $($e.FullName)"
        }
    }
}

# d) Every Report/definition JSON parses
foreach ($e in $z.Entries) {
    if ($e.FullName -like 'Report/definition/*.json' -or $e.FullName -like 'Report/definition/**/*.json') {
        $s = $e.Open(); $sr = New-Object System.IO.StreamReader($s)
        $txt = $sr.ReadToEnd(); $sr.Close()
        try { $null = $txt | ConvertFrom-Json } catch { "INVALID JSON: $($e.FullName) — $($_.Exception.Message)" }
    }
}
$z.Dispose()
```

Expect: `DataModel`, `Version`, `Metadata`, `Settings`, `[Content_Types].xml` present; `SecurityBindings` absent; no leading slashes; no BOMs; all JSON parses.

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

## Styling visuals — the property names that actually work

The formatting-pane labels in Power BI Desktop do NOT match the JSON property names. If you guess, PBI silently strips the bogus properties on next save — the file opens fine, but your style vanishes. Always build real changes in PBI once, save, extract, and copy the property names from the saved JSON.

### Where styling lives in `visual.json`

- `visual.visualContainerObjects` — the **outer frame** of the visual tile (background, border, drop shadow, title bar). Applies to every visual type.
- `visual.objects` — the **inside** of the visual (plot area, axes, legend, data points, card value, card inner fill). Property names here are per-visual-type and must match the theme's `visualStyles[visualType]` schema.

### Container frame (works for every visual)

```json
"visualContainerObjects": {
  "background": [{"properties":{
    "color":{"solid":{"color":{"expr":{"Literal":{"Value":"'#EEF2FF'"}}}}},
    "transparency":{"expr":{"Literal":{"Value":"0D"}}}
  }}],
  "border": [{"properties":{
    "show":{"expr":{"Literal":{"Value":"true"}}},
    "color":{"solid":{"color":{"expr":{"Literal":{"Value":"'#E2E8F0'"}}}}},
    "radius":{"expr":{"Literal":{"Value":"12D"}}}
  }}],
  "dropShadow": [{"properties":{
    "show":{"expr":{"Literal":{"Value":"true"}}},
    "color":{"solid":{"color":{"expr":{"Literal":{"Value":"'#0F172A'"}}}}},
    "transparency":{"expr":{"Literal":{"Value":"85D"}}},
    "position":{"expr":{"Literal":{"Value":"'Outer'"}}},
    "preset":{"expr":{"Literal":{"Value":"'BottomRight'"}}}
  }}]
}
```

### cardVisual (KPI tile) — gotchas

The big number is NOT `callout.color` and NOT `value.color`. PBI silently drops both.

- Big number font color: `value.fontColor` (not `color`).
- Big number font / size / weight: `value.fontFamily`, `value.fontSize`, `value.bold`.
- Category label under the number: `label.fontColor`, `label.fontFamily`, `label.fontSize`.
- Tinted fill behind the number: `cardCalloutArea.backgroundFillColor`, with a `selector` that names the metadata of the measure projection — e.g. `"selector":{"metadata":"Sum(financials.Profit)"}`. Read that string from `visual.query.queryState.Data.projections[0].queryRef`.

```json
"objects": {
  "value": [{"properties":{
    "fontColor":{"solid":{"color":{"expr":{"Literal":{"Value":"'#4F46E5'"}}}}},
    "fontFamily":{"expr":{"Literal":{"Value":"'Segoe UI Semibold'"}}},
    "fontSize":{"expr":{"Literal":{"Value":"40D"}}},
    "bold":{"expr":{"Literal":{"Value":"true"}}}
  }}],
  "label": [{"properties":{
    "fontColor":{"solid":{"color":{"expr":{"Literal":{"Value":"'#64748B'"}}}}}
  }}],
  "cardCalloutArea": [{
    "properties": {
      "show":{"expr":{"Literal":{"Value":"true"}}},
      "backgroundFillColor":{"solid":{"color":{"expr":{"Literal":{"Value":"'#E0E7FF'"}}}}}
    },
    "selector":{"metadata":"Sum(financials.Sales)"}
  }]
}
```

### Bar / column / line / donut — data colors

- `dataPoint.defaultColor` — the fallback color for all data points in a single-measure chart.
- `dataPoint.showAllDataPoints: true` — enables per-datapoint overrides. On its own it does NOT auto-rotate the theme palette per category; PBI uses `dataColors[0]` for every bar unless you provide per-category selectors.
- Per-category "rainbow" bars require a `dataPoint` entry per category value with a `selector.data[0].scopeId` that references that specific value. Constructing `scopeId` by hand is fragile — build it in PBI once, save, then copy the exact shape.
- Multi-series charts (line with two years, donut with categories) DO rotate the theme palette automatically — so a custom theme palette gives them rainbow colors for free.

### Custom report theme (vibrant palette, report-wide)

Drop a theme file at `Report/StaticResources/RegisteredResources/<name>.json`:
```json
{
  "name":"ModernVibrant",
  "dataColors":["#6366F1","#EC4899","#10B981","#F59E0B","#06B6D4","#8B5CF6","#F43F5E","#84CC16","#0EA5E9","#D946EF"],
  "background":"#FFFFFF","foreground":"#0F172A","tableAccent":"#6366F1",
  "good":"#10B981","neutral":"#F59E0B","bad":"#F43F5E"
}
```

Register in `Report/definition/report.json` by (a) appending a `resourcePackages` entry and (b) adding `themeCollection.customTheme`:
```json
"themeCollection": {
  "baseTheme": { ...existing SharedResources Fluent2... },
  "customTheme": {
    "name":"ModernVibrant",
    "reportVersionAtImport":{"visual":"2.12.0","report":"3.4.0","page":"2.3.1"},
    "type":"RegisteredResources"
  }
},
"resourcePackages": [
  { ...existing SharedResources... },
  {"name":"RegisteredResources","type":"RegisteredResources",
   "items":[{"name":"ModernVibrant","path":"ModernVibrant.json","type":"CustomTheme"}]}
]
```
`[Content_Types].xml` already has a `Default Extension="json"` so no entry needs to be added.

### Page-level styling (`page.json`)

```json
"objects": {
  "background": [{"properties":{
    "color":{"solid":{"color":{"expr":{"Literal":{"Value":"'#F8FAFC'"}}}}},
    "transparency":{"expr":{"Literal":{"Value":"0D"}}}
  }}],
  "outspace": [{"properties":{
    "color":{"solid":{"color":{"expr":{"Literal":{"Value":"'#0F172A'"}}}}}
  }}]
}
```

### Discovering property names when PBI drops yours

If your style vanishes on save:
1. Open the patched pbix in PBI Desktop and set the property manually via the Format pane.
2. Save, then re-extract the pbix and diff the visual JSON — the saved file shows the exact property name PBI expects.
3. Clone that shape in your writer.

## Anti-patterns (things that will corrupt the file)
- Python `zipfile.ZipFile(...)` repack → `ExternalAttributes` mismatch → "corrupted".
- Recompressing `DataModel` → `MashupValidationError`.
- Leaving `SecurityBindings` when you added any new file → "SecurityBindings invalid".
- Zip entries with a leading `/` (e.g. `/Version`) → `MissingBIPackagePart: /Version`.
- UTF-8 BOM in any `Report/definition/**` JSON → `Only text with UTF8 encoding without BOM (byte order marks) is supported`.
- `Set-Content -Encoding UTF8` and `[System.IO.File]::WriteAllText($p, $s)` without an explicit encoding on .NET Framework / PowerShell 5.1 → both write a BOM. Always pass `New-Object System.Text.UTF8Encoding($false)`.
- Guessing the `visual.json` schema → silent corruption OR silently stripped properties. Always build real visuals in PBI once, inspect their JSON, then clone the shape.
- `pbi-tools` v1.2.0 does NOT round-trip PBIR (it extracts Model/TMDL but skips `Report/definition/pages/*`). Don't use it for this.

## Quick sanity test
Make a copy of a PBIX, repack it with the script above with NO changes. It must open cleanly. If it doesn't, the repack itself is wrong — fix that before editing content.

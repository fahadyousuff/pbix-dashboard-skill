# pbix-dashboard-skill

A **Claude Code skill** for programmatically editing Power BI Desktop `.pbix` files
(PBIR Enhanced Report Format) **without corrupting them**.

## What it solves

Editing a modern `.pbix` by hand or with Python's `zipfile` almost always produces
"corrupted file" / `MashupValidationError` / `Invalid SecurityBindings` errors on
open. Three traps:

1. Python `zipfile` writes non-zero `external_attr` that Power BI rejects.
2. Recompressing the embedded `DataModel` breaks it.
3. Adding any file invalidates the DPAPI-signed `SecurityBindings` manifest.

This skill captures the method that actually works — strip `SecurityBindings`,
preserve `DataModel` compression, repack with .NET `ZipArchive` and
`ExternalAttributes = 0`. Power BI Desktop regenerates `SecurityBindings` on the
next Save.

## Install as a Claude Code skill

```bash
# User-level (available in every session)
mkdir -p ~/.claude/skills/pbix-dashboard
cp -r SKILL.md scripts/ ~/.claude/skills/pbix-dashboard/

# or project-level
mkdir -p .claude/skills/pbix-dashboard
cp -r SKILL.md scripts/ .claude/skills/pbix-dashboard/
```

Claude Code auto-loads the skill when the user asks anything matching its
description (edit pbix, add visuals, Power BI dashboard, etc.).

## Repo contents

- `SKILL.md` — the skill itself (frontmatter + full method).
- `scripts/repack_pbix.ps1` — the PowerShell repacker (the critical piece).
- `scripts/build_visuals.py` — helpers to emit PBIR `visual.json` for common chart types.
- `scripts/connect_powerbi.py` — ADOMD connection to a live Power BI Desktop session.

## Quick start

```powershell
# 1. Extract the pbix
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("Report.pbix", "pbix_src")

# 2. Edit Report\definition\pages\... (add page, add visual.json, update pages.json)
#    Use build_visuals.py helpers or write JSON directly.

# 3. Repack
.\scripts\repack_pbix.ps1 -Src pbix_src -Dst Report_patched.pbix

# 4. Open in Power BI Desktop. It will regenerate SecurityBindings on Save.
```

## Credits

Method discovered by trial and error across many corruption variants.
Packaged by Claude for reuse.

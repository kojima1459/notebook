# build/

This directory contains the build pipeline for the two distribution xlsm files.

## Required artifacts (not in Git)

- `template_chatbot.xlsm` - empty macro-enabled workbook with a Sheet1.
  Created once manually:
  1. Open Excel, save a new blank workbook as `template_chatbot.xlsm` here.
  2. Do not add any modules. The build script imports them.
- `template_admin.xlsm` - same, for the admin variant.
- `vendor/JsonConverter.bas` - VBA-JSON v2.3.1 from
  <https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1>
  - File: `JsonConverter.bas`
  - SHA-256: pin in your own copy (the upstream release page lists it).

Track template xlsm files via Git LFS once they exist:
```
git lfs track "build/*.xlsm"
```

## Build

On a Windows machine with Excel installed:
```
cd build
powershell -ExecutionPolicy Bypass -File build.ps1
```

Output appears in `../dist/`.

## Trust Center prerequisite

`build.ps1` uses the `VBProject` object to inject modules. This is gated by
**File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model**. Enable it on the build
machine only (not user PCs).

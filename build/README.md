# build/

This directory contains the build pipeline for the two distribution xlsm files.

## Required artifacts

The chat UI is rendered onto a worksheet at runtime by `modChatUI`, so no
UserForm and no template workbooks are required. `build.ps1` creates
blank macro-enabled workbooks from scratch via Excel COM, then injects
the modules.

VBA-JSON v2.3.1 (`vendor/JsonConverter.bas`, MIT licensed) is bundled in
the repo, so no manual download is needed. Source:
<https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1>

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

# build/

This directory contains the build pipeline for the two distribution xlsm files.

## Required artifacts (not in Git)

The chat UI is now rendered onto a worksheet at runtime by `modChatUI`,
so no UserForm and no template workbooks are required. `build.ps1`
creates blank macro-enabled workbooks from scratch via Excel COM, then
injects the modules.

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

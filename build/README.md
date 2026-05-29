# build/

This directory contains the build pipeline for the two distribution xlsm files.

## Required artifacts (not in Git)

- `template_chatbot.xlsm` - macro-enabled workbook used as the build seed.
  Created **once** manually because Excel's `.frm` text export cannot
  reproduce control layouts (the binary `.frx` is required). Steps:
  1. Save a new blank workbook as `template_chatbot.xlsm` here.
  2. Open VBA editor, **Insert > UserForm**, rename to `frmChat`.
  3. Drop on the form (default names matter - the build relies on them):
     - `txtConversation`: TextBox, `MultiLine=True`, `Locked=True`, large
     - `txtQuestion`: TextBox, `MultiLine=True`, `EnterKeyBehavior=False`
     - `btnSend`: CommandButton, caption "送信"
     - `chkConsent`: CheckBox, caption "ナレッジ改善に質問本文を提供する"
     - `lblStatus`: Label
  4. Save and close. Do not add any modules - the build script imports them.
- `template_admin.xlsm` - blank macro-enabled workbook, no form needed.
  Just save a new blank `.xlsm` here.
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

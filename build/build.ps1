# ============================================================================
# build.ps1 - assemble Chatbot.xlsm and Admin_KnowledgeBuilder.xlsm
# ----------------------------------------------------------------------------
# Drives Excel via COM to create blank macro-enabled workbooks from scratch
# and inject every .bas / .cls module. No UserForms are used; the chat UI
# is rendered onto a worksheet at runtime by modChatUI, so this build
# requires zero manual template preparation.
#
# Prerequisites:
#   - Windows + Excel installed
#   - Trust Center: "Trust access to the VBA project object model" enabled
#     (File > Options > Trust Center > Trust Center Settings > Macro Settings)
#   - build/vendor/JsonConverter.bas present (download VBA-JSON v2.3.1)
# ============================================================================

[CmdletBinding()]
param(
    [string]$OutDir = "..\dist",
    [switch]$SkipAdmin,
    [switch]$SkipChatbot
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$vendorJson = Join-Path $root "vendor\JsonConverter.bas"
if (-not (Test-Path $vendorJson)) {
    Write-Warning "vendor\JsonConverter.bas missing. Falling back to the stub in src\shared."
    $vendorJson = (Resolve-Path (Join-Path $root "..\src\shared\JsonConverter.bas")).Path
} else {
    $vendorJson = (Resolve-Path $vendorJson).Path
}

# Order matters: shared first (modConfig/modPaths used everywhere),
# then role-specific. ThisWorkbook last because we replace its code module.
$sharedModules = @(
    "modTypes.bas",
    "modConfig.bas",
    "modPaths.bas",
    "modHttpClient.bas",
    "modKeyVault.bas",
    "modApiGateway.bas"
)

$chatbotModules = @(
    "modIndexReader.bas",
    "modSimilarity.bas",
    "modRagEngine.bas",
    "modUserProfile.bas",
    "modPiiGuard.bas",
    "modRateLimiter.bas",
    "modUsageLogger.bas",
    "modChatUI.bas",
    "modBoot.bas",
    "ThisWorkbook.cls"
)

$adminModules = @(
    "modChunker.bas",
    "modExtractor.bas",
    "modExtractorWord.bas",
    "modExtractorAcrobat.bas",
    "modExtractorExcel.bas",
    "modIndexWriter.bas",
    "modKnowledgeBuilder.bas",
    "modKeyEnroller.bas",
    "modUsageAggregator.bas"
)

# Excel file format constants
$xlOpenXMLWorkbookMacroEnabled = 52

function New-BlankXlsm {
    param([object]$Excel, [string]$Path)
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $wb = $Excel.Workbooks.Add()
    $wb.SaveAs($Path, $xlOpenXMLWorkbookMacroEnabled)
    $wb.Close($false)
}

function Build-Xlsm {
    param(
        [object]$Excel,
        [string]$Target,
        [string]$ExtraSrcDir,
        [string[]]$ExtraModules
    )

    New-BlankXlsm -Excel $Excel -Path $Target

    $wb = $Excel.Workbooks.Open($Target)
    try {
        $vbProj = $wb.VBProject

        # Shared modules
        foreach ($mod in $sharedModules) {
            $path = (Resolve-Path (Join-Path "..\src\shared" $mod)).Path
            $vbProj.VBComponents.Import($path) | Out-Null
        }
        # JsonConverter (vendor copy preferred; falls back to stub)
        $vbProj.VBComponents.Import($vendorJson) | Out-Null

        # Role-specific
        foreach ($mod in $ExtraModules) {
            $path = Join-Path $ExtraSrcDir $mod
            if (-not (Test-Path $path)) {
                Write-Warning "Skipping missing module: $mod"
                continue
            }
            $path = (Resolve-Path $path).Path
            if ($mod -eq "ThisWorkbook.cls") {
                # Workbook_Open lives in the document module; overwrite its code
                $code = Get-Content $path -Raw
                # Drop the .cls header so we feed only Option Explicit + Subs
                $split = $code -split "Attribute VB_Exposed = True", 2
                if ($split.Count -eq 2) { $code = $split[1].TrimStart("`r","`n") }
                $tw = $vbProj.VBComponents.Item("ThisWorkbook").CodeModule
                if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
                $tw.AddFromString($code)
            } else {
                $vbProj.VBComponents.Import($path) | Out-Null
            }
        }

        $wb.Save()
    } finally {
        $wb.Close($true)
    }
}

# Single Excel instance for both builds
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
try {
    if (-not $SkipChatbot) {
        $tgt = Join-Path $OutDir "Chatbot.xlsm"
        Build-Xlsm -Excel $excel `
                   -Target $tgt `
                   -ExtraSrcDir (Resolve-Path "..\src\chatbot").Path `
                   -ExtraModules $chatbotModules
        Write-Host "Built $tgt"
    }
    if (-not $SkipAdmin) {
        $tgt = Join-Path $OutDir "Admin_KnowledgeBuilder.xlsm"
        Build-Xlsm -Excel $excel `
                   -Target $tgt `
                   -ExtraSrcDir (Resolve-Path "..\src\admin").Path `
                   -ExtraModules $adminModules
        Write-Host "Built $tgt"
    }
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

Write-Host ""
Write-Host "Done. Output: $OutDir"

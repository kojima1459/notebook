# ============================================================================
# build.ps1 - assemble Chatbot.xlsm and Admin_KnowledgeBuilder.xlsm
# ----------------------------------------------------------------------------
# Strategy: Excel COM automation drives a blank .xlsm template, imports each
# .bas/.cls/.frm module, and saves. Runs on Windows with Excel installed.
#
# Prerequisites:
#   - Excel installed
#   - Trust Center: "Trust access to the VBA project object model" enabled
#   - build/vendor/JsonConverter.bas present (download VBA-JSON v2.3.1)
#   - build/template_chatbot.xlsm and build/template_admin.xlsm exist
#     (empty .xlsm with a Sheet1 and macro-enabled)
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

$vendorJson = Join-Path $root "vendor\JsonConverter.bas"
if (-not (Test-Path $vendorJson)) {
    Write-Warning "vendor\JsonConverter.bas missing. Falling back to the stub in src\shared."
    $vendorJson = Join-Path $root "..\src\shared\JsonConverter.bas"
}

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
    "modBoot.bas",
    "frmChat.frm",
    "ThisWorkbook.cls"
)

$adminModules = @(
    "modChunker.bas",
    "modExtractor.bas",
    "modIndexWriter.bas",
    "modKnowledgeBuilder.bas",
    "modKeyEnroller.bas"
)

function Build-Xlsm {
    param(
        [string]$Template,
        [string]$Target,
        [string[]]$ExtraSrcDirs,
        [string[]]$ExtraModules
    )

    Copy-Item $Template $Target -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open((Resolve-Path $Target).Path)
        $vbProj = $wb.VBProject

        # Drop any pre-existing modules of the same name (for re-build)
        $namesToDrop = @($sharedModules + $ExtraModules) | ForEach-Object {
            [System.IO.Path]::GetFileNameWithoutExtension($_)
        }
        foreach ($comp in @($vbProj.VBComponents)) {
            if ($namesToDrop -contains $comp.Name -and $comp.Name -ne "ThisWorkbook") {
                $vbProj.VBComponents.Remove($comp)
            }
        }

        # Import shared first
        foreach ($mod in $sharedModules) {
            $path = Join-Path "..\src\shared" $mod
            if ($mod -eq "JsonConverter.bas") { $path = $vendorJson }
            $vbProj.VBComponents.Import((Resolve-Path $path).Path) | Out-Null
        }
        # Always import the JsonConverter (may be vendor or stub)
        $vbProj.VBComponents.Import((Resolve-Path $vendorJson).Path) | Out-Null

        # Import role-specific
        foreach ($dir in $ExtraSrcDirs) {
            foreach ($mod in $ExtraModules) {
                $path = Join-Path $dir $mod
                if (Test-Path $path) {
                    if ($mod -eq "ThisWorkbook.cls") {
                        # Replace the existing ThisWorkbook code rather than import
                        $code = Get-Content $path -Raw
                        $vbProj.VBComponents("ThisWorkbook").CodeModule.AddFromString $code
                    } else {
                        $vbProj.VBComponents.Import((Resolve-Path $path).Path) | Out-Null
                    }
                }
            }
        }

        $wb.Save()
        $wb.Close($false)
    } finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

if (-not $SkipChatbot) {
    Build-Xlsm `
        -Template "template_chatbot.xlsm" `
        -Target  (Join-Path $OutDir "Chatbot.xlsm") `
        -ExtraSrcDirs @("..\src\chatbot") `
        -ExtraModules $chatbotModules
    Write-Host "Built $OutDir\Chatbot.xlsm"
}

if (-not $SkipAdmin) {
    Build-Xlsm `
        -Template "template_admin.xlsm" `
        -Target  (Join-Path $OutDir "Admin_KnowledgeBuilder.xlsm") `
        -ExtraSrcDirs @("..\src\admin") `
        -ExtraModules $adminModules
    Write-Host "Built $OutDir\Admin_KnowledgeBuilder.xlsm"
}

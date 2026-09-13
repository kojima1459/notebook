# ============================================================================
# run_excel_tests.ps1 - 「Nexus Agent」Windows実機 自動スモークテスト
# ----------------------------------------------------------------------------
# 前提: Windows + デスクトップ版Excel(VBA必須。Web版/Storeアプリ版は不可)
# 使い方(PowerShellをそのまま実行):
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 -Target prod
# 何をするか:
#   1. レジストリでVBA信頼設定(AccessVBOM/マクロ許可)を現ユーザーに設定
#   2. dist\MyBookshelf_dev.xlsm (または --prod版) をCOMで開く
#      → 初回起動の自己インストーラが走り、vba_srcから全モジュールが組み上がる
#   3. modTestRunner.RunAllPureTests を実行し PASS/FAIL を取得
#   4. 結果を wintest\result_*.log に保存。FAILありなら終了コード1
# ============================================================================
param(
    [ValidateSet("dev", "prod")] [string]$Target = "dev",
    [string]$MockAddinPath = ""   # 「リボンちゃん(検証用).xlam」のフルパス(任意)
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot   # mybookshelf/
$xlsm = Join-Path $root ("dist\" + $(if ($Target -eq "dev") { "MyBookshelf_dev.xlsm" } else { "MyBookshelf.xlsm" }))
$log  = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

function Log([string]$msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

if (-not (Test-Path $xlsm)) { throw "ビルド成果物が見つかりません: $xlsm (先に python build\build_mybookshelf.py --$Target)" }

# --- 1) VBA信頼設定(現ユーザーのみ。テスト終了後に元の値へ復元する) ---
Log "VBA信頼設定(HKCU)を確認・設定します(終了時に元の値へ復元します)"
$officeVersions = @("16.0")   # Office 2016以降/365は16.0
$originalSecurity = @{}       # $v => @{ AccessVBOM = <元の値 or $null>; VBAWarnings = <元の値 or $null> }
foreach ($v in $officeVersions) {
    $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    New-Item -Path $key -Force | Out-Null

    $existing = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    $originalSecurity[$v] = @{
        AccessVBOM  = if ($existing -and ($existing.PSObject.Properties.Name -contains "AccessVBOM"))  { $existing.AccessVBOM }  else { $null }
        VBAWarnings = if ($existing -and ($existing.PSObject.Properties.Name -contains "VBAWarnings")) { $existing.VBAWarnings } else { $null }
    }

    Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord    # VBAプロジェクトOMを信頼
    Set-ItemProperty -Path $key -Name VBAWarnings -Value 1 -Type DWord   # マクロを有効化
}

function Restore-VbaTrustSettings {
    foreach ($v in $officeVersions) {
        $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
        $orig = $originalSecurity[$v]
        if ($null -ne $orig.AccessVBOM) {
            Set-ItemProperty -Path $key -Name AccessVBOM -Value $orig.AccessVBOM -Type DWord
        } else {
            Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue
        }
        if ($null -ne $orig.VBAWarnings) {
            Set-ItemProperty -Path $key -Name VBAWarnings -Value $orig.VBAWarnings -Type DWord
        } else {
            Remove-ItemProperty -Path $key -Name VBAWarnings -ErrorAction SilentlyContinue
        }
    }
    Log "VBA信頼設定(HKCU)を元の値へ復元しました"
}

$excel = $null
$wb = $null
$tmpWb = $null
$addin = $null
$exitCode = 1
try {
    # --- 2) Excel起動+(任意)ニセリボンちゃんの有効化 --------------------------
    Log "Excelを起動します"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true          # 目視も可能に。完全無人化するなら $false
    $excel.DisplayAlerts = $false

    if ($MockAddinPath -ne "") {
        if (-not (Test-Path $MockAddinPath)) { throw "アドインが見つかりません: $MockAddinPath" }
        Log "ニセリボンちゃんを有効化: $MockAddinPath"
        $tmpWb = $excel.Workbooks.Add()   # AddIns.Addはブックが1つ開いている必要がある
        $addin = $excel.AddIns.Add($MockAddinPath, $false)
        $addin.Installed = $true
        $tmpWb.Close($false)
    }

    # --- 3) マイ本棚AIを開く(初回は自己インストーラが走る) --------------------
    Log "開きます: $xlsm"
    $wb = $excel.Workbooks.Open($xlsm)
    Start-Sleep -Seconds 20   # 自己インストール+Boot完了待ち(遅いPCは伸ばす)

    # --- 4) 純ロジックテストを実機VBAで実行 ------------------------------------
    Log "modTestRunner.RunAllPureTests を実行します"
    $excel.Run("'" + $wb.Name + "'!modTestRunner.RunAllPureTests") | Out-Null
    $fails  = $excel.Run("'" + $wb.Name + "'!modTestRunner.Failures")
    $report = $excel.Run("'" + $wb.Name + "'!modTestRunner.ReportText")

    Log "---- テスト結果 ----"
    $report -split "`n" | ForEach-Object { Log $_ }
    if ([int]$fails -eq 0) {
        Log "実機テスト: 全PASS"
        $exitCode = 0
    } else {
        Log "実機テスト: FAIL ${fails}件"
    }

    $wb.Close($false)
}
catch {
    Log ("エラー: " + $_.Exception.Message)
}
finally {
    # 子COMオブジェクトを先にClose/Release→最後にExcel本体をQuit/Release。
    # 順序を守らないとExcelプロセスが残留し、次回実行が不安定になる。
    if ($tmpWb) {
        try { $tmpWb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($tmpWb)
        $tmpWb = $null
    }
    if ($addin) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($addin)
        $addin = $null
    }
    if ($wb) {
        try { $wb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        $wb = $null
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Restore-VbaTrustSettings

    Log "ログ: $log"
}
exit $exitCode

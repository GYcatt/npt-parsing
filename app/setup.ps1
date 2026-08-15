# NPT Parsing 環境建置（由 setup.bat 呼叫）
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$PyVersion = "3.12.10"
$GetPipUrl = "https://bootstrap.pypa.io/get-pip.py"
$Requirements = Join-Path $Root "requirements.txt"
$VenvPython = Join-Path $Root "venv\Scripts\python.exe"
$RuntimeDir = Join-Path $Root "runtime"
$RuntimePython = Join-Path $RuntimeDir "python.exe"

function Write-Step($num, $total, $msg) {
    Write-Host ""
    Write-Host "[$num/$total] $msg"
}

function Write-Ok($msg) { Write-Host "OK  $msg" }
function Write-Fail($msg) { Write-Host "FAIL  $msg" }

function Get-SystemPython {
    $commands = @(
        { & py -3 -c "import sys; print(sys.executable)" 2>$null },
        { & python -c "import sys; print(sys.executable if sys.version_info[0] >= 3 else '')" 2>$null }
    )
    foreach ($run in $commands) {
        try {
            $out = & $run
            if ($LASTEXITCODE -ne 0 -or -not $out) { continue }
            $path = ($out | Select-Object -Last 1).ToString().Trim()
            if (-not $path) { continue }
            if ($path -match "WindowsApps") { continue }
            if (Test-Path $path) { return $path }
        } catch { }
    }
    return $null
}

function Install-Requirements($pythonExe) {
    if (-not (Test-Path $Requirements)) {
        throw "找不到 requirements.txt"
    }
    & $pythonExe -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "升級 pip 失敗" }
    & $pythonExe -m pip install -r $Requirements
    if ($LASTEXITCODE -ne 0) { throw "安裝依賴失敗，請檢查網路連線" }
    & $pythonExe -c "import pandas, openpyxl"
    if ($LASTEXITCODE -ne 0) { throw "套件驗證失敗（pandas / openpyxl）" }
}

function Install-PortablePython {
    if ($env:PROCESSOR_ARCHITECTURE -eq "x86") {
        throw "目前只支援 64 位元 Windows。"
    }
    $zipName = "python-$PyVersion-embed-amd64.zip"
    $url = "https://www.python.org/ftp/python/$PyVersion/$zipName"
    $zipPath = Join-Path $env:TEMP $zipName
    $pipPath = Join-Path $env:TEMP "get-pip.py"

    Write-Host "正在下載免安裝 Python $PyVersion ..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing

    if (Test-Path $RuntimeDir) {
        Remove-Item $RuntimeDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $RuntimeDir | Out-Null
    Expand-Archive -Path $zipPath -DestinationPath $RuntimeDir -Force

    New-Item -ItemType Directory -Path (Join-Path $RuntimeDir "Lib\site-packages") -Force | Out-Null

    $pth = Get-ChildItem -Path $RuntimeDir -Filter "python*._pth" | Select-Object -First 1
    if (-not $pth) { throw "免安裝 Python 缺少 ._pth，下載可能不完整" }
    $zipBase = [IO.Path]::GetFileNameWithoutExtension($pth.Name)
    @"
$zipBase.zip
.
Lib\site-packages
import site
"@ | Set-Content -Path $pth.FullName -Encoding ASCII

    Write-Host "正在安裝 pip ..."
    Invoke-WebRequest -Uri $GetPipUrl -OutFile $pipPath -UseBasicParsing
    & $RuntimePython $pipPath --no-warn-script-location
    if ($LASTEXITCODE -ne 0) { throw "安裝 pip 失敗" }
}

try {
    Write-Host "========================================"
    Write-Host "  NPT Parsing - 環境建置"
    Write-Host "========================================"

    Write-Step 1 4 "檢查既有執行環境"
    $pythonExe = $null
    if (Test-Path $VenvPython) {
        $pythonExe = $VenvPython
        Write-Ok "使用既有 venv"
    } elseif (Test-Path $RuntimePython) {
        $pythonExe = $RuntimePython
        Write-Ok "使用既有免安裝 Python（runtime）"
    } else {
        Write-Host "尚未建立執行環境"
    }

    if (-not $pythonExe) {
        Write-Step 2 4 "尋找系統 Python 或下載免安裝版"
        $sysPy = Get-SystemPython
        if ($sysPy) {
            Write-Ok "找到系統 Python：$sysPy"
            Write-Host "正在建立 venv..."
            & $sysPy -m venv (Join-Path $Root "venv")
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
                throw "建立 venv 失敗"
            }
            $pythonExe = $VenvPython
            Write-Ok "venv 建立成功"
        } else {
            Write-Host "未偵測到 Python，改為下載免安裝版（不需系統管理員）"
            Install-PortablePython
            $pythonExe = $RuntimePython
            Write-Ok "免安裝 Python 已就緒"
        }
    } else {
        Write-Step 2 4 "執行環境已存在，跳過下載"
        Write-Ok "略過"
    }

    Write-Step 3 4 "升級 pip 並安裝套件（pandas、openpyxl）"
    Install-Requirements $pythonExe
    Write-Ok "依賴安裝完成"

    Write-Step 4 4 "完成"
    Write-Host ""
    Write-Host "========================================"
    Write-Host "  環境建置完成"
    Write-Host "  接下來請關閉 Excel，再雙擊 NPT Parsing.bat"
    Write-Host "========================================"
    exit 0
} catch {
    Write-Host ""
    Write-Fail $_.Exception.Message
    Write-Host "Python 官方下載頁：https://www.python.org/downloads/"
    Write-Host "若自行安裝 Python，請勾選 Add python.exe to PATH，再重跑 setup.bat"
    exit 1
}

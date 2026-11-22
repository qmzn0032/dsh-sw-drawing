# requires -version 5.1
<#
    dsh-sw-drawing installer.

    What it does:
      1. checks the bundle files are all here
      2. finds a Python 3 (SW_BRIDGE_PYTHON -> hermes venv -> py -3 -> python)
      3. checks pywin32 is importable by that Python
      4. looks for SOLIDWORKS (warning only)
      5. writes packages\dsh-sw-drawing\cordis.patch.yml with this machine's real path
      6. prints the two clicks needed in DSH

    Usage:
      install.cmd            (double click)
      powershell -File install.ps1 -Check     (only report, write nothing)
#>
[CmdletBinding()]
param(
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pkg = Join-Path $here 'packages\dsh-sw-drawing'
$tpl = Join-Path $pkg 'cordis.patch.yml.template'
$out = Join-Path $pkg 'cordis.patch.yml'

function Step($n, $text) { Write-Host ""; Write-Host "[$n] $text" -ForegroundColor Cyan }
function Ok($text) { Write-Host "    OK   $text" -ForegroundColor Green }
function Warn($text) { Write-Host "    !!   $text" -ForegroundColor Yellow }
function Die($text) { Write-Host "    FAIL $text" -ForegroundColor Red; Write-Host ""; exit 1 }

Write-Host ""
Write-Host "=== dsh-sw-drawing installer ===" -ForegroundColor White
Write-Host "bundle: $pkg"
if ($Check) { Write-Host "(dry run: nothing will be written)" -ForegroundColor DarkGray }

# ---------------------------------------------------------------- 1 files
Step 1 '检查 bundle 文件'
$required = @(
    'package.json',
    'cordis.patch.yml.template',
    'py\mcp_sw_server.py',
    'py\launch_server.cmd',
    'py\sw_com.py',
    'py\sw_drawing.py',
    'py\sw_features.py',
    'py\gb_rules.json'
)
foreach ($rel in $required) {
    $p = Join-Path $pkg $rel
    if (-not (Test-Path $p)) { Die "缺文件：$rel" }
}
Ok "8 个必需文件都在"

# ---------------------------------------------------------------- 2 python
Step 2 '找 Python 3'
$tries = New-Object System.Collections.ArrayList
if ($env:SW_BRIDGE_PYTHON) { [void]$tries.Add(@($env:SW_BRIDGE_PYTHON)) }
$venvPy = Join-Path $env:LOCALAPPDATA 'hermes\hermes-agent\venv\Scripts\python.exe'
if (Test-Path $venvPy) { [void]$tries.Add(@($venvPy)) }
$pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
if ($pyLauncher) { [void]$tries.Add(@($pyLauncher.Source, '-3')) }
$pyExe = Get-Command python.exe -ErrorAction SilentlyContinue
if ($pyExe) { [void]$tries.Add(@($pyExe.Source)) }

$python = $null
foreach ($t in $tries) {
    $exe = $t[0]
    $pre = @($t | Select-Object -Skip 1)
    try {
        $ver = & $exe @pre -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $ver) {
            $python = [pscustomobject]@{ Exe = $exe; Pre = $pre; Version = ($ver | Select-Object -First 1) }
            break
        }
    } catch { }
}
if (-not $python) {
    Die "没找到 Python 3。装一个（https://www.python.org/downloads/），或用 SW_BRIDGE_PYTHON 指定 python.exe。"
}
$shown = (@($python.Exe) + @($python.Pre)) -join ' '
Ok "$shown  (Python $($python.Version))"

# ---------------------------------------------------------------- 3 pywin32
Step 3 '检查 pywin32'
$canImport = & $python.Exe @($python.Pre) -c "import win32com.client; print('ok')" 2>$null
if ($LASTEXITCODE -ne 0 -or "$canImport".Trim() -ne 'ok') {
    Warn "这个 Python 里没有 pywin32，MCP 服务起不来。先跑："
    Write-Host ('    "' + $python.Exe + '" ' + ($python.Pre -join ' ') + ' -m pip install pywin32') -ForegroundColor Yellow
    if ($Check) { } else { Write-Host ""; Write-Host "装完再跑一次 install.cmd。" -ForegroundColor Yellow; exit 2 }
} else {
    Ok 'pywin32 可用'
}

# ---------------------------------------------------------------- 4 solidworks
Step 4 '找 SOLIDWORKS'
$swCandidates = @()
if ($env:SW_EXE) { $swCandidates += $env:SW_EXE }
$swCandidates += 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe'
$swCandidates += 'D:\SW\SOLIDWORKS\sldworks.exe'
$swFound = $swCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if ($swFound) {
    Ok "找到 $swFound"
    if (-not $env:SW_EXE) {
        Write-Host "       (工具只连已经运行的 SW，这个路径只有你显式要求启动时才用；可用 SW_EXE 覆盖)" -ForegroundColor DarkGray
    }
} else {
    Warn '没找到 SOLIDWORKS —— 只要它装在别处就没事；工具本身只连已经运行的实例。'
}

# ---------------------------------------------------------------- 5 patch
Step 5 '生成本机 patch'
if (-not (Test-Path $tpl)) { Die "缺模板：$tpl" }
$yaml = Get-Content -LiteralPath $tpl -Raw -Encoding UTF8
$escaped = $pkg.Replace("'", "''")            # YAML 单引号串里，单引号要写两次
$yaml = $yaml.Replace('__PLUGIN_DIR__', $escaped)
if ($yaml -match '__PLUGIN_DIR__') { Die '模板里还有没替换掉的占位符' }

if ($Check) {
    Warn "dry run：略过写 $out"
} else {
    [System.IO.File]::WriteAllText($out, $yaml, (New-Object System.Text.UTF8Encoding($false)))
    if (-not (Test-Path $out)) { Die "写不出来：$out" }
    $back = Get-Content -LiteralPath $out -Raw -Encoding UTF8
    if ($back -notmatch [regex]::Escape($pkg)) { Die '写完回读里找不到插件路径' }
    Ok "$out"
    Write-Host "       launcher: $(Join-Path $pkg 'py\launch_server.cmd')" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- 6 next
Step 6 '接下来两步（在 DSH 里做）'
Write-Host ""
Write-Host "  1) 设置 → 插件 → 添加插件，选这个目录：" -ForegroundColor White
Write-Host "       $pkg" -ForegroundColor Yellow
Write-Host "     （如果没有这个按钮，把下面这句话原样发给 DSH：）" -ForegroundColor DarkGray
Write-Host "       安装这个本地插件目录：$pkg" -ForegroundColor Yellow
Write-Host ""
Write-Host "  2) 重启 DSH，打开 SOLIDWORKS 并随便开一个零件，然后问：" -ForegroundColor White
Write-Host "       用 sw_status 看看 SW 状态" -ForegroundColor Yellow
Write-Host ""
Write-Host "  看到的 running 是 true、document 是你的零件名，就算通了。" -ForegroundColor White
Write-Host ""

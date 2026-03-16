# ===========================================================
# WeightToExcel.ps1 by Shawn.lee@sgs.com
# Mettler Toledo Scale to Excel via COM API
# ===========================================================

# === 設定檔路徑（與腳本/exe 同目錄） ===
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([Environment]::GetCommandLineArgs()[0]) }
$configPath = Join-Path $scriptDir "config.json"

# === 預設值 ===
$defaults = @{
    comPort            = "COM1"
    baudRate           = 9600
    parity             = "None"
    dataBits           = 8

    stopBits           = "One"
    moveDirection      = "down"
    idleTimeoutMinutes = 30
}

# === 設定精靈 ===
function Show-SetupWizard {
    param($defaults)
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host "  First-time Setup / Reconfigure" -ForegroundColor Yellow
    Write-Host "  Press Enter to accept [default]" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host ""

    $ports = [System.IO.Ports.SerialPort]::GetPortNames()
    if ($ports.Count -gt 0) {
        Write-Host "  Available COM Ports: $($ports -join ', ')" -ForegroundColor Cyan
    }
    else {
        Write-Host "  No COM Ports detected." -ForegroundColor Red
    }
    Write-Host ""

    $cfg = @{}

    $v = Read-Host "  COM Port [$($defaults.comPort)]"
    $cfg.comPort = if ($v) { $v } else { $defaults.comPort }

    $v = Read-Host "  Baud Rate [$($defaults.baudRate)]"
    $cfg.baudRate = if ($v) { [int]$v } else { $defaults.baudRate }

    $v = Read-Host "  Parity (None/Odd/Even) [$($defaults.parity)]"
    $cfg.parity = if ($v) { $v } else { $defaults.parity }

    $v = Read-Host "  Data Bits (7/8) [$($defaults.dataBits)]"
    $cfg.dataBits = if ($v) { [int]$v } else { $defaults.dataBits }

    $v = Read-Host "  Stop Bits (One/Two) [$($defaults.stopBits)]"
    $cfg.stopBits = if ($v) { $v } else { $defaults.stopBits }

    $v = Read-Host "  下一格游標移動方向 (down/right) [$($defaults.moveDirection)]"
    $cfg.moveDirection = if ($v) { $v } else { $defaults.moveDirection }

    $v = Read-Host "  閒置斷線時間 (min) [$($defaults.idleTimeoutMinutes)]"
    $cfg.idleTimeoutMinutes = if ($v) { [int]$v } else { $defaults.idleTimeoutMinutes }

    $cfg | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
    Write-Host ""
    Write-Host "  Config saved to: $configPath" -ForegroundColor Green
    Write-Host ""
    return $cfg
}

# === 載入或建立設定 ===
$runSetup = $false

if (Test-Path $configPath) {
    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
    Write-Host ""
    Write-Host "  Config loaded: $($cfg.comPort), $($cfg.baudRate)bps" -ForegroundColor Gray
    Write-Host "  Developed by Shawn Lee <shawn.lee@sgs.com>" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  按 [r] 重新設定參數或等待 3 秒..." -ForegroundColor Gray

    $timeout = 3000
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $timeout) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'R') {
                $runSetup = $true
                break
            }
        }
        Start-Sleep -Milliseconds 100
    }
}
else {
    $runSetup = $true
}

if ($runSetup) {
    $currentDefaults = if ($cfg) {
        @{
            comPort            = $cfg.comPort
            baudRate           = $cfg.baudRate
            parity             = $cfg.parity
            dataBits           = $cfg.dataBits
            stopBits           = $cfg.stopBits
            moveDirection      = $cfg.moveDirection
            idleTimeoutMinutes = $cfg.idleTimeoutMinutes
        }
    }
    else { $defaults }
    $cfg = Show-SetupWizard $currentDefaults
}

# === Excel COM 連線函數 ===
function Get-ExcelInstance {
    try {
        $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        if ($null -ne $xl -and $xl.Visible) {
            return $xl
        }
        return $null
    }
    catch {
        return $null
    }
}

# === 初始化 ===
$host.UI.RawUI.WindowTitle = "Scale Monitor [$($cfg.comPort)]"
$counter = 0

$port = New-Object System.IO.Ports.SerialPort $cfg.comPort, $cfg.baudRate, $cfg.parity, $cfg.dataBits, $cfg.stopBits
$port.DtrEnable = $true
$port.RtsEnable = $true
$port.ReadTimeout = 1000

$lastActionTime = Get-Date

# === 主程式 ===
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Scale Monitor - $($cfg.comPort)" -ForegroundColor Cyan
Write-Host "  Author: Shawn Lee (shawn.lee@sgs.com)" -ForegroundColor DarkCyan
Write-Host "------------------------------------------" -ForegroundColor Cyan
Write-Host "  $($cfg.baudRate)bps / $($cfg.parity) / $($cfg.dataBits) / $($cfg.stopBits)" -ForegroundColor Cyan
Write-Host "  游標移動方向: $($cfg.moveDirection) | 逾時斷線: $($cfg.idleTimeoutMinutes) min" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [註] 關閉視窗即斷開連線，釋放 Excel 模式，方能輸入至LIMS重量系統" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Opening $($cfg.comPort) ..." -ForegroundColor Yellow
    $port.Open()
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Connected." -ForegroundColor Green

    # 嘗試連接 Excel (靜默模式)
    $excel = Get-ExcelInstance
    if ($null -ne $excel) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Excel linked." -ForegroundColor Green
    }
    Write-Host ""
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Waiting for weight data..." -ForegroundColor Gray
    Write-Host ""

    while ($true) {
        if ((Get-Date) - $lastActionTime -gt [TimeSpan]::FromMinutes($cfg.idleTimeoutMinutes)) {
            Write-Host ""
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Idle timeout ($($cfg.idleTimeoutMinutes) min). Auto-disconnecting..." -ForegroundColor Yellow
            break
        }

        try {
            $data = $port.ReadLine()
            if ($data -match '([-+]?\d+\.\d+)\s*(g|mg|kg|ct|oz|lb)\b') {
                $counter++
                $value = $matches[1]
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] #$counter  Raw: $($data.Trim())  ->  Sent: $value" -ForegroundColor White
                $lastActionTime = Get-Date

                if ($null -eq $excel) { $excel = Get-ExcelInstance }
                if ($null -ne $excel) {
                    try {
                        $activeCell = $excel.ActiveCell
                        # 檢查是否為有效的儲存格物件
                        if ($null -ne $activeCell) {
                            $activeCell.Value2 = [double]$value
                            # 根據設定的方向移動
                            if ($cfg.moveDirection -eq "right") {
                                $activeCell.Offset(0, 1).Select() | Out-Null
                            }
                            else {
                                $activeCell.Offset(1, 0).Select() | Out-Null
                            }
                        }
                        else {
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel busy: No active cell found (Is cell editing?)" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        # 捕捉常見的編輯鎖定錯誤 (HRESULT: 0x800AC472)
                        if ($_.Exception.Message -match "0x800AC472") {
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel is BUSY (editing cell). Please finish editing." -ForegroundColor Yellow
                        }
                        else {
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel write error: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
                else {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel not available, data skipped." -ForegroundColor Red
                }
            }
        }
        catch [System.TimeoutException] {
            continue
        }
    }
}
catch {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($null -ne $port -and $port.IsOpen) {
        $port.Close()
    }
    Write-Host ""
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Port closed. Total sent: $counter" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Press any key to close..." -ForegroundColor Gray
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
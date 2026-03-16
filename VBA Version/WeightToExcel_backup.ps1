# ===========================================================
# WeightToExcel.ps1
# 用途：透過 RS232 串列埠讀取 Mettler Toledo 天平數據，
#       並透過 Excel COM API 直接寫入儲存格。
# 呼叫方式：由 ScaleConnector.bas (VBA) 在背景啟動
# ===========================================================

# === 參數區：接收 VBA 傳入的參數 ===
param (
    [string]$comPort = "COM1"
)

# === Excel COM 連線函數 ===
function Get-ExcelInstance {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        return $null
    }
}

# === 設定視窗標題，方便辨識 ===
$host.UI.RawUI.WindowTitle = "Scale Monitor [$comPort]"
$counter = 0   # 接收計數器

# === 串列埠設定 ===
$portName = $comPort
$port = New-Object System.IO.Ports.SerialPort $portName, 9600, None, 8, "One"
$port.DtrEnable = $true      # 啟用 DTR（Data Terminal Ready）訊號
$port.RtsEnable = $true      # 啟用 RTS（Request To Send）訊號
$port.ReadTimeout = 1000     # 讀取逾時 1 秒（避免無限等待）

# === 訊號檔案路徑（與 VBA 端約定，置於 Windows Temp 資料夾） ===
$tempDir = $env:TEMP
$stopSignal = "$tempDir\StopScaleSignal.txt"      # VBA 發送停止指令用
$successSignal = "$tempDir\ScaleSuccess.txt"       # 回報連線成功
$failSignal = "$tempDir\ScaleFail.txt"             # 回報連線失敗

# 清理上一次殘留的訊號檔案，避免誤判
if (Test-Path $stopSignal) { Remove-Item $stopSignal -Force }
if (Test-Path $successSignal) { Remove-Item $successSignal -Force }
if (Test-Path $failSignal) { Remove-Item $failSignal -Force }

# === 閒置自動斷開設定 ===
$idleTimeoutMinutes = 30     # 超過此時間(分鐘)無數據輸入則自動結束，釋放 COM Port
$lastActionTime = Get-Date

# === 主程式 ===
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Scale Monitor - $comPort" -ForegroundColor Cyan
Write-Host "  Idle timeout: $idleTimeoutMinutes min" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

try {
    # 開啟串列埠並回報成功訊號給 VBA
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Opening $comPort ..." -ForegroundColor Yellow
    $port.Open()
    New-Item -Path $successSignal -ItemType File -Force | Out-Null
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Connected. Waiting for data..." -ForegroundColor Green

    # 取得 Excel 實例
    $excel = Get-ExcelInstance
    if ($null -eq $excel) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] WARNING: Excel not found. Will retry on first data." -ForegroundColor Yellow
    } else {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Excel instance connected." -ForegroundColor Green
    }
    Write-Host ""
    
    # 進入無窮迴圈，持續監聽天平數據
    while ($true) {
        # 退出條件 1：VBA 發送了停止訊號（使用者按下「斷開連線」）
        if (Test-Path $stopSignal) {
            Remove-Item $stopSignal -Force
            Write-Host ""
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Stop signal received. Disconnecting..." -ForegroundColor Yellow
            break
        }
        
        # 退出條件 2：閒置超時，自動釋放資源防止長期佔用
        if ((Get-Date) - $lastActionTime -gt [TimeSpan]::FromMinutes($idleTimeoutMinutes)) {
            Write-Host ""
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Idle timeout ($idleTimeoutMinutes min). Auto-disconnecting..." -ForegroundColor Yellow
            break 
        }
        
        # 嘗試讀取天平傳來的數據
        try {
            $data = $port.ReadLine()
            # 只匹配真正的重量數據：帶小數點的數值 + 重量單位（排除選單雜訊）
            if ($data -match '([-+]?\d+\.\d+)\s*(g|mg|kg|ct|oz|lb)\b') {
                $counter++
                $value = $matches[1]
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] #$counter  Raw: $($data.Trim())  ->  Sent: $value" -ForegroundColor White
                $lastActionTime = Get-Date   # 重置閒置計時器

                # 透過 COM API 直接寫入 Excel 儲存格（不受輸入法影響）
                if ($null -eq $excel) { $excel = Get-ExcelInstance }
                if ($null -ne $excel) {
                    try {
                        $excel.ActiveCell.Value2 = [double]$value
                        # 讀取 VBA 設定的方向，移動到下一格
                        # xlDown = -4121, xlToRight = -4161
                        if ($excel.MoveAfterReturnDirection -eq -4161) {
                            $excel.ActiveCell.Offset(0, 1).Select() | Out-Null
                        } else {
                            $excel.ActiveCell.Offset(1, 0).Select() | Out-Null
                        }
                    } catch {
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel write failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')]   Excel not available, data skipped." -ForegroundColor Red
                }
            }
        }
        catch [System.TimeoutException] { continue }  # 讀取逾時屬正常現象，繼續等待
    }
}
catch {
    # 連線失敗時，將錯誤訊息寫入失敗訊號檔案供 VBA 讀取
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Set-Content -Path $failSignal -Value $_.Exception.Message -Force
}
finally {
    # 無論如何都要確保串列埠被正確關閉，釋放硬體資源
    if ($null -ne $port -and $port.IsOpen) { 
        $port.Close()
    }
    Write-Host ""
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Port closed. Total sent: $counter" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
}

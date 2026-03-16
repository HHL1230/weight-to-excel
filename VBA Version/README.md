# 天平自動化輸入系統 — 架構與佈署指南

## 🎯 專案概述

* **系統目的**：將 Mettler Toledo 天平數據透過 RS232 無縫輸入 Excel 儲存格。
* **設計原則**：隨插即用、多檔同步防脫節、真實狀態查勤、閒置自動釋放防佔用。

---

## 📁 檔案結構

| 檔案 | 類型 | 說明 |
| --- | --- | --- |
| `WeightToExcel.ps1` | PowerShell | 底層硬體通訊腳本，讀取天平數據並透過 Excel COM API 直接寫入儲存格 |
| `ScaleConnector.bas` | VBA 模組 | 前端介面控制器，含環境設定區、連線管理、燈號顯示 |
| `WorkbookEvents.cls` | VBA 事件 | 活頁簿生命週期管理，處理開啟/切換/關閉事件 |

---

## ⚙️ 佈署設定（僅需修改此區塊）

所有環境參數集中於 `ScaleConnector.bas` 頂部，佈署至不同電腦或 Excel 表格時只需調整以下常數：

```vba
Private Const SCRIPT_PATH As String = "C:\...\WeightToExcel.ps1"  ' PowerShell 腳本路徑
Public Const STATUS_CELL As String = "M2"        ' 狀態燈號儲存格位置
Public Const COMPORT_CELL As String = "N3"       ' COM Port 下拉選單儲存格位置
Public Const MOVE_DIRECTION As String = "down"   ' 輸入後游標方向："down" = 往下，"right" = 往右
```

---

## 🛠️ VBE 安裝指引

按下 `Alt + F11` 開啟 Visual Basic Editor (VBE)，依照以下步驟將程式碼匯入目標活頁簿：

### 步驟一：匯入 `ScaleConnector.bas`（標準模組）

1. 在 VBE 左側的**專案總管**中，找到目標活頁簿的專案名稱。
2. 點選功能表 **檔案 → 匯入檔案**（或在專案名稱上按右鍵 → **匯入檔案**）。
3. 選取 `ScaleConnector.bas` 檔案，按下**開啟**。
4. 匯入後會出現在專案的「**模組**」資料夾中，模組名稱為 `ScaleConnector`。
5. 開啟模組，**修改頂部環境設定區**的常數以符合該電腦與表格的需求。

### 步驟二：設定 `WorkbookEvents.cls`（ThisWorkbook 事件）

> ⚠️ `.cls` 不能直接匯入覆蓋 ThisWorkbook，需手動貼上程式碼。

1. 在專案總管中，展開目標活頁簿的專案，雙擊 **ThisWorkbook**。
2. 開啟 `WorkbookEvents.cls` 檔案（用記事本或 VS Code），**複製全部程式碼**。
3. 將程式碼**貼入** ThisWorkbook 的程式碼視窗中。
4. 確認左上角的物件下拉選單顯示為 **Workbook**。

### 步驟三：儲存活頁簿

* 必須儲存為 **`.xlsm`（啟用巨集的活頁簿）** 格式，否則 VBA 程式碼會被移除。
* 儲存路徑：**檔案 → 另存新檔 → 選擇「Excel 啟用巨集的活頁簿 (*.xlsm)」**。

### 完成後的專案結構（VBE 中）

```text
VBAProject (你的活頁簿.xlsm)
├── Microsoft Excel 物件
│   ├── Sheet1 (工作表)
│   └── ThisWorkbook        ← WorkbookEvents.cls 的程式碼在這裡
└── 模組
    └── ScaleConnector       ← ScaleConnector.bas 匯入於此
```

---

## 一、【底層硬體通訊】 `WeightToExcel.ps1`

* **參數接收**：接收 VBA 傳入的 COM Port 參數，嘗試開啟獨佔連線。
* **狀態回報**：將連線成功/失敗結果寫入 Windows Temp 資料夾，供 VBA 判斷。
* **數據寫入**：透過 **Excel COM Object API** 直接寫入儲存格，不受輸入法影響，不需 Excel 視窗焦點。
* **數據過濾**：僅匹配帶小數點且後接重量單位（g/mg/kg 等）的數據，自動排除天平選單雜訊。
* **游標方向**：從 Excel 的 `MoveAfterReturnDirection` 動態讀取，用 `ActiveCell.Offset()` 移動游標。
* **監控視窗**：執行時顯示即時狀態（連線、接收計數、原始數據 vs 過濾後數值、退出原因）。
* **安全退出機制（釋放 COM Port）**：
  1. **主動接收**：偵測到 Excel 產生的 `StopScaleSignal.txt` 則立刻斷開。
  2. **被動防護**：閒置超過 30 分鐘無數據輸入，自動斷開以防資源長期佔用。

---

## 二、【前端介面與控制器】 `ScaleConnector.bas`

* **環境設定區**：所有佈署相關參數集中於檔案頂部，方便維護。
* **真實查勤 (WMI)**：透過 `Win32_Process` 查詢系統背景是否有 PowerShell 腳本存活，做為唯一的「真實狀態」來源，拒絕燈號假象。
* **啟動連線 (StartScaleConnection)**：讀取 Excel 儲存格的 COM Port 設定，若未連線則在背景呼叫 PS 腳本。成功後亮起「綠燈」。
* **斷開連線 (StopScaleConnection)**：產生 `StopScaleSignal.txt` 讓 PS 腳本安全結束，並將燈號切換為「灰燈」。

---

## 三、【活頁簿生命週期管理】 `WorkbookEvents.cls`

* **`Workbook_Open`（開啟檔案）**：
  * 套用該活頁簿的游標移動方向設定。
  * 向作業系統查勤並同步當前燈號，不干擾現有連線。
* **`Workbook_Activate`（視窗活化）**：
  * **關鍵防呆機制**！當使用者從其他視窗切換回此 Excel 視窗時：
    * 瞬間切換游標方向至該活頁簿的設定（解決多檔不同方向需求）。
    * 向系統查勤並校正燈號（解決多檔案並存時的狀態脫節問題）。
* **`Workbook_BeforeClose`（關閉檔案）**：
  * **智慧判斷**：計算當前開啟的 Excel 數量。
    * 若為「最後一個檔案」：自動發送停止訊號釋放 COM Port。
    * 若「還有其他檔案開啟」：默默關閉，不干擾其他仍在連線的活頁簿。

Attribute VB_Name = "ScaleConnector"
' =============================================
' === 環境設定區（佈署時僅需修改此區塊） ===
' =============================================
Private Const SCRIPT_PATH As String = "C:\Users\shawn_lee\Desktop\WeightToExcel.ps1"  ' PowerShell 腳本路徑
Public Const STATUS_CELL As String = "M2"           ' 狀態燈號儲存格位置
Public Const COMPORT_CELL As String = "N3"          ' COM Port 下拉選單儲存格位置
Public Const MOVE_DIRECTION As String = "down"        ' 輸入後游標方向："down" = 往下，"right" = 往右

' =============================================
' === 以下為功能程式碼，一般不需修改 ===
' =============================================
Private Const STOP_SIGNAL_FILE As String = "StopScaleSignal.txt"  ' 內部通訊用，不需修改

' === 查勤函數：向 Windows 系統查詢腳本是否正在跑(天平連線中) ===
Function IsScaleScriptRunning() As Boolean
    Dim objWMIService As Object, colProcesses As Object
    Dim query As String

    On Error GoTo ErrorHandler
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    ' 查詢所有包含 "WeightToExcel.ps1" 參數的 PowerShell 執行序
    query = "Select * from Win32_Process Where Name = 'powershell.exe' And CommandLine Like '%WeightToExcel.ps1%'"
    Set colProcesses = objWMIService.ExecQuery(query)

    If colProcesses.Count > 0 Then
        IsScaleScriptRunning = True  ' 抓到了，正在背景執行 (連線中)
    Else
        IsScaleScriptRunning = False ' 沒抓到，完全沒在跑 (已斷開)
    End If
    Exit Function

ErrorHandler:
    IsScaleScriptRunning = False ' 若查詢失敗，預設為未執行 (已斷開)
End Function

' === 連線狀態燈號切換 ===
Sub UpdateStatusUI(statusCell As Range, isRunning As Boolean)
    With statusCell
        If isRunning Then
            .Value = "天平連線中"
            .Interior.Color = RGB(144, 238, 144)
            .Font.Size = 14
            .Font.Color = RGB(0, 100, 0)
        Else
            .Value = "天平已斷線"
            .Interior.Color = RGB(255, 192, 203)
            .Font.Size = 14
            .Font.Color = RGB(220, 20, 60)
        End If
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' === 啟動連線按鈕 ===
Sub StartScaleConnection()
    Dim psCommand As String
    Dim statusCell As Range, comPortCell As Range
    Dim userComPort As String

    ' === 1. 設定介面儲存格（來自環境設定區） ===
    Set statusCell = ActiveSheet.Range(STATUS_CELL)
    Set comPortCell = ActiveSheet.Range(COMPORT_CELL)

    ' 讀取使用者選擇的 COM Port 並去除多餘空白
    userComPort = Trim(comPortCell.Value)

    ' 防呆：如果同事忘記選 COM Port
    If userComPort = "" Then
        MsgBox "啟動失敗！請先在 " & COMPORT_CELL & " 儲存格選擇天平的 COM Port。", vbExclamation, "缺少設定"
        Exit Sub
    End If

    ' 2. 查勤：如果已經在跑了，就只更新燈號
    If IsScaleScriptRunning() Then
        UpdateStatusUI statusCell, True
        Exit Sub
    End If

    ' 清理舊訊號
    Dim stopSignal As String
    stopSignal = Environ("TEMP") & "\" & STOP_SIGNAL_FILE
    If Dir(stopSignal) <> "" Then Kill stopSignal

    ' === 3. 防呆：檢查腳本檔案是否存在於設定路徑 ===
    If Dir(SCRIPT_PATH) = "" Then
        MsgBox "找不到連線腳本！" & vbCrLf & "請確認系統環境已設定完成，且腳本放置於：" & vbCrLf & SCRIPT_PATH, vbCritical, "環境設定錯誤"
        Exit Sub
    End If

    ' === 4. 組合包含外部參數的 PowerShell 指令 ===
    psCommand = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Minimized -File """ & SCRIPT_PATH & """ -comPort """ & userComPort & """"

    ' 啟動背景腳本
    Shell psCommand, vbMinimizedNoFocus

    ' 5. 等待 2 秒讓 PowerShell 啟動，然後再次查勤確認真實狀態
    Application.Wait (Now + TimeValue("0:00:02"))
    UpdateStatusUI statusCell, IsScaleScriptRunning()
End Sub

' === 斷開連線按鈕 ===
Sub StopScaleConnection()
    Dim stopSignal As String
    Dim fileNum As Integer
    Dim statusCell As Range
    Set statusCell = ActiveSheet.Range(STATUS_CELL)

    ' 1. 只有在真的有跑的時候，才需要發送停止訊號
    If IsScaleScriptRunning() Then
        stopSignal = Environ("TEMP") & "\" & STOP_SIGNAL_FILE
        fileNum = FreeFile
        Open stopSignal For Output As #fileNum
        Close #fileNum

        ' 等待 PowerShell 收到訊號並結束
        Application.Wait (Now + TimeValue("0:00:02"))
    End If

    ' 2. 更新為真實狀態
    UpdateStatusUI statusCell, IsScaleScriptRunning()
End Sub

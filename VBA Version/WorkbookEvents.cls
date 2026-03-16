' === 活頁簿開啟事件：同步燈號 + 設定游標方向 ===
Private Sub Workbook_Open()
    ApplyMoveDirection
    UpdateStatusUI ActiveSheet.Range(STATUS_CELL), IsScaleScriptRunning()
End Sub

' === 視窗活化事件：切換回此檔案時，校正燈號 + 切換游標方向 ===
Private Sub Workbook_Activate()
    ApplyMoveDirection
    UpdateStatusUI ActiveSheet.Range(STATUS_CELL), IsScaleScriptRunning()
End Sub

' === 關閉檔案事件：智慧判斷是否需要釋放 COM Port ===
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim visibleWbCount As Integer
    Dim wb As Workbook
    Dim stopSignal As String
    Dim fileNum As Integer

    ' 如果天平根本沒在連線，就不用做任何事
    If Not IsScaleScriptRunning() Then Exit Sub

    ' 計算目前還有多少個 Excel 活頁簿開著
    visibleWbCount = 0
    For Each wb In Application.Workbooks
        ' 只計算可見的活頁簿（排除增益集等隱藏的活頁簿）
        If wb.IsAddin = False And wb.Windows(1).Visible = True Then
            visibleWbCount = visibleWbCount + 1
        End If
    Next wb

    ' 如果是最後一個檔案，自動發送停止訊號釋放 COM Port
    If visibleWbCount <= 1 Then
        stopSignal = Environ("TEMP") & "\StopScaleSignal.txt"
        fileNum = FreeFile
        Open stopSignal For Output As #fileNum
        Close #fileNum
    End If
End Sub

' === 根據 MOVE_DIRECTION 常數設定 Excel 的游標移動方向 ===
Private Sub ApplyMoveDirection()
    Application.MoveAfterReturn = True
    If MOVE_DIRECTION = "right" Then
        Application.MoveAfterReturnDirection = xlToRight
    Else
        Application.MoveAfterReturnDirection = xlDown
    End If
End Sub
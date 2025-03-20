在 Excel VBA 中，同時操作 3 個以上的工作簿（Workbook）時，需要注意以下幾點：

確保所有工作簿都已打開：可以用 Workbooks.Open 方法打開文件，或者用 Workbooks("文件名.xlsx") 確保它們已經打開。
使用變數存儲對應的 Workbook 對象：用變數來存儲 Workbook 物件，避免頻繁用 Workbooks("文件名.xlsx") 調用，提高效率。
切換工作簿時注意激活 (Activate) 或直接操作對象 (Avoid Select/Activate)：應直接使用 wb.Sheets("Sheet1").Cells(1,1).Value = "xxx" 而非 Activate + Select。
錯誤處理：用 On Error Resume Next 或 On Error GoTo 來防止錯誤中斷程序。
文件保存與關閉：結束時記得 wb.Save 或 wb.Close False（不保存直接關閉），避免數據丟失。
範例：同時操作三個 Excel 文件
假設有三個 Excel 文件：
Data1.xlsx
Data2.xlsx
Report.xlsx（這是我們的主報告文件）
需求：
從 Data1.xlsx 和 Data2.xlsx 讀取數據，將數據寫入 Report.xlsx。
Data1.xlsx 的 Sheet1 A1 資料 -> Report.xlsx 的 Sheet1 B2。
Data2.xlsx 的 Sheet1 A1 資料 -> Report.xlsx 的 Sheet1 C2。
VBA 程式碼

Sub MergeDataFromMultipleWorkbooks()
    Dim wbData1 As Workbook
    Dim wbData2 As Workbook
    Dim wbReport As Workbook
    
    Dim wsData1 As Worksheet
    Dim wsData2 As Worksheet
    Dim wsReport As Worksheet
    
    Dim data1Path As String
    Dim data2Path As String
    Dim reportPath As String
    
    ' 設定文件路徑（請修改為實際路徑）
    data1Path = "C:\Users\YourName\Desktop\Data1.xlsx"
    data2Path = "C:\Users\YourName\Desktop\Data2.xlsx"
    reportPath = "C:\Users\YourName\Desktop\Report.xlsx"
    
    ' 打開三個工作簿（如果已打開，則直接引用）
    On Error Resume Next
    Set wbData1 = Workbooks.Open(data1Path)
    Set wbData2 = Workbooks.Open(data2Path)
    Set wbReport = Workbooks.Open(reportPath)
    On Error GoTo 0
    
    ' 確保所有工作簿都打開
    If wbData1 Is Nothing Or wbData2 Is Nothing Or wbReport Is Nothing Then
        MsgBox "有文件無法打開，請檢查路徑！", vbCritical
        Exit Sub
    End If
    
    ' 設定對應的工作表
    Set wsData1 = wbData1.Sheets("Sheet1")
    Set wsData2 = wbData2.Sheets("Sheet1")
    Set wsReport = wbReport.Sheets("Sheet1")
    
    ' 讀取數據並寫入報告工作簿
    wsReport.Cells(2, 2).Value = wsData1.Cells(1, 1).Value ' Data1 A1 -> Report B2
    wsReport.Cells(2, 3).Value = wsData2.Cells(1, 1).Value ' Data2 A1 -> Report C2
    
    ' 保存並關閉工作簿
    wbReport.Save
    wbData1.Close False ' 不保存直接關閉
    wbData2.Close False
    wbReport.Close True ' 保存並關閉
    
    ' 釋放變數
    Set wbData1 = Nothing
    Set wbData2 = Nothing
    Set wbReport = Nothing
    Set wsData1 = Nothing
    Set wsData2 = Nothing
    Set wsReport = Nothing
    
    MsgBox "數據合併完成！", vbInformation
End Sub


解釋程式碼
設定文件路徑：
你需要修改 data1Path、data2Path 和 reportPath 為你的實際 Excel 文件路徑。
打開工作簿：
Workbooks.Open(data1Path) 來打開文件，若已打開則直接使用。
On Error Resume Next 防止文件不存在導致錯誤中斷。
設置變數來存儲對應的工作表：
wsData1 = wbData1.Sheets("Sheet1") 代表 Data1.xlsx 的 Sheet1。
數據傳輸：
直接用 .Cells(row, col).Value 來傳輸數據，避免 Select 或 Activate。
保存和關閉文件：
wbReport.Save 儲存報告文件。
wbData1.Close False 表示不儲存變更並關閉。
wbReport.Close True 儲存並關閉報告文件。
擴展應用
如果你需要處理更多文件，可以：

把 wbData1、wbData2 變數改為數組 wbData(i)，用 For 迴圈批量讀取數據。
使用 Dictionary 來存儲數據並在 Report.xlsx 內進行匹配處理。
這樣你可以更靈活地管理多個 Excel 檔案，而不會發生錯誤。

Attribute VB_Name = "Module1"
Dim FileName As String

Sub シート名取得()

Dim Wb As Workbook
Dim Sh
Dim ShName As String

Dim i As Integer

Application.ScreenUpdating = False

Call ファイル名取得

Call ジョブ名取得

Set Wb = Workbooks.Open(FileName)

i = 1

ThisWorkbook.Worksheets(1).Range("B:B").Clear
ThisWorkbook.Worksheets(1).Range("A:A").Interior.Color = RGB(255, 255, 255)
ThisWorkbook.Worksheets(1).Cells(1, 2).Value = "シート名"

For Each Sh In Wb.Worksheets

    ShName = Sh.Name
    ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value = ShName
    ThisWorkbook.Worksheets(1).Cells(i + 1, 1).Interior.Color = RGB(255, 0, 102)
    i = i + 1

Next Sh

Wb.Close

Application.ScreenUpdating = True

End Sub
Sub シート名設定()

Dim Wb As Workbook
Dim Sh
Dim ShName As String

Dim SetFileName As String

Dim i As Integer

Application.ScreenUpdating = False

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

i = 1

For Each Sh In Wb.Worksheets

    Sh.Name = CStr(i) + "小五郎"
    i = i + 1

Next Sh

i = 1

For Each Sh In Wb.Worksheets

    ShName = ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value
    Sh.Name = ShName
    Sh.Cells(1, 1).Value = "■" + CStr(ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value)
    Sh.Select
    Sh.Cells(1, 1).Select
    i = i + 1

Next Sh

Wb.Worksheets(1).Select

Application.DisplayAlerts = False

Wb.Save

Wb.Close

Application.DisplayAlerts = True

Application.ScreenUpdating = True

End Sub
Sub ジョブ名取得()

Dim Wb As Workbook
Dim Sh
Dim JobName As String

Dim SetFileName As String

Dim i As Integer

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

i = 1

ThisWorkbook.Worksheets(1).Range("C:C").Clear
ThisWorkbook.Worksheets(1).Cells(1, 3).Value = "ジョブ名"

For Each Sh In Wb.Worksheets

    JobName = Sh.Cells(3, 4).Value
    ThisWorkbook.Worksheets(1).Cells(i + 1, 3).Value = JobName
    i = i + 1

Next Sh

Wb.Close

End Sub
Sub ファイル名取得()

    FileName = Application.GetOpenFilename
    ThisWorkbook.Worksheets(1).Cells(1, 4).Value = FileName

End Sub
Sub シート名初期化()

Dim Wb As Workbook
Dim Sh
Dim ShName As String

Dim SetFileName As String

Dim i As Integer

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

i = 1

For Each Sh In Wb.Worksheets

    Sh.Name = CStr(i) + "小五郎"
    i = i + 1

Next Sh

Wb.Close

End Sub
Sub 重複チェック()

    Dim Str As String
    Dim MaxRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim ErrStr() As String
    Dim ErrFlag As Integer
    Dim Sh As Worksheet
    Dim Juhuku As String
    
    ErrCount = 0
    
    MaxRow = Range("B1").End(xlDown).Row
    
    Set Sh = ThisWorkbook.Worksheets(1)
    
    For i = 1 To MaxRow
    
        k = 0
    
        For j = 1 To MaxRow
        
            If Sh.Cells(j, 2).Value = Sh.Cells(i, 2).Value Then
            
                k = k + 1
            
            End If
            
        Next j
        
        If k >= 2 Then
        
            Juhuku = Juhuku & Sh.Cells(i, 2).Value & vbCrLf
            ErrFlag = 1
        
        End If
    
    Next i
    
    If ErrFlag = 1 Then
        
        MsgBox ("重複値があります" & vbCrLf & Juhuku)
        
    
    End If

End Sub
Sub シート追加()

Dim Wb As Workbook
Dim Sh
Dim ShName As String
Dim SheetNum As Integer

Dim SetFileName As String

Dim i As Integer

Application.ScreenUpdating = False

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

SheetNum = ThisWorkbook.Worksheets(1).Cells(2, 4).Value

For i = 1 To SheetNum

    Wb.Worksheets.Add

Next i

Application.DisplayAlerts = False

Wb.Save

Wb.Close

Application.DisplayAlerts = True

Application.ScreenUpdating = True


End Sub

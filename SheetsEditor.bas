Attribute VB_Name = "Module1"
Dim FileName As String

Sub �V�[�g���擾()

Dim Wb As Workbook
Dim Sh
Dim ShName As String

Dim i As Integer

Application.ScreenUpdating = False

Call �t�@�C�����擾

Call �W���u���擾

Set Wb = Workbooks.Open(FileName)

i = 1

ThisWorkbook.Worksheets(1).Range("B:B").Clear
ThisWorkbook.Worksheets(1).Range("A:A").Interior.Color = RGB(255, 255, 255)
ThisWorkbook.Worksheets(1).Cells(1, 2).Value = "�V�[�g��"

For Each Sh In Wb.Worksheets

    ShName = Sh.Name
    ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value = ShName
    ThisWorkbook.Worksheets(1).Cells(i + 1, 1).Interior.Color = RGB(255, 0, 102)
    i = i + 1

Next Sh

Wb.Close

Application.ScreenUpdating = True

End Sub
Sub �V�[�g���ݒ�()

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

    Sh.Name = CStr(i) + "���ܘY"
    i = i + 1

Next Sh

i = 1

For Each Sh In Wb.Worksheets

    ShName = ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value
    Sh.Name = ShName
    Sh.Cells(1, 1).Value = "��" + CStr(ThisWorkbook.Worksheets(1).Cells(i + 1, 2).Value)
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
Sub �W���u���擾()

Dim Wb As Workbook
Dim Sh
Dim JobName As String

Dim SetFileName As String

Dim i As Integer

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

i = 1

ThisWorkbook.Worksheets(1).Range("C:C").Clear
ThisWorkbook.Worksheets(1).Cells(1, 3).Value = "�W���u��"

For Each Sh In Wb.Worksheets

    JobName = Sh.Cells(3, 4).Value
    ThisWorkbook.Worksheets(1).Cells(i + 1, 3).Value = JobName
    i = i + 1

Next Sh

Wb.Close

End Sub
Sub �t�@�C�����擾()

    FileName = Application.GetOpenFilename
    ThisWorkbook.Worksheets(1).Cells(1, 4).Value = FileName

End Sub
Sub �V�[�g��������()

Dim Wb As Workbook
Dim Sh
Dim ShName As String

Dim SetFileName As String

Dim i As Integer

SetFileName = ThisWorkbook.Worksheets(1).Cells(1, 4).Value

Set Wb = Workbooks.Open(SetFileName)

i = 1

For Each Sh In Wb.Worksheets

    Sh.Name = CStr(i) + "���ܘY"
    i = i + 1

Next Sh

Wb.Close

End Sub
Sub �d���`�F�b�N()

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
        
        MsgBox ("�d���l������܂�" & vbCrLf & Juhuku)
        
    
    End If

End Sub
Sub �V�[�g�ǉ�()

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

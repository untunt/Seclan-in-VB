Attribute VB_Name = "Song"
Public Const DirPlayer = "C:\Program Files (x86)\foobar2000\foobar2000.exe"
Public Const DirLib = "E:\虚拟存储区\Symphony\"
Public Const SheetName = "曲目"
Public Const SheetDict = "字典"
Public Const RowStart = 3
Public Const ColumnStart = 1
Public Const ColumnLib = ColumnStart
Public Const ColumnNumber = ColumnLib + 1
Public Const ColumnDate = ColumnNumber + 1
Public Const ColumnDur = ColumnDate + 1
Public Const ColumnForm = ColumnDur + 4
Public Const ColumnName = ColumnForm + 1
Public Const ColumnComposer = ColumnName + 1
Dim RowLast As Long
Dim ColumnLast As Long
Dim Org As Boolean
Dim OrgRowName As String
Dim ColumnOrg As Long
Public Function Name2File(SongName As String, Optional WithMP3 As Boolean = True) As String
    Name2File = Replace(SongName, ":", "-")
    Name2File = Replace(Name2File, "/", "_")
    Name2File = Replace(Replace(Name2File, "? ", "？"), "?", "？")
    If WithMP3 Then Name2File = Name2File & ".mp3"
End Function ' E9.19
Public Function Name2Song(filename As String) As String
    Name2Song = Replace(filename, ".mp3", "")
    Name2Song = Replace(Name2Song, "- ", ": ")
    Name2Song = Replace(Name2Song, "_", "/")
    Name2Song = RTrim(Replace(Name2Song, "？", "? "))
End Function ' E9.19
Private Sub Set_Last()
    RowLast = ActiveSheet.UsedRange.Rows.Count
    ColumnLast = ActiveSheet.UsedRange.Columns.Count
End Sub ' E9.16
Private Sub Set_Org()
    Application.ScreenUpdating = False
    If ActiveCell.Row < RowStart Then
        Org = False
        Exit Sub
    End If
    Org = True
    OrgRowName = Cells(ActiveCell.Row, ColumnName)
    ColumnOrg = ActiveCell.Column
End Sub ' E9.20凌晨，E10.17更新
Private Sub Back_Org()
    Dim c As Long
    If Org Then
        c = Range(Cells(1, ColumnName), Cells(RowLast, ColumnName)).Find(OrgRowName, , , xlWhole).Row
        Range(Cells(1, 1), Cells(1, 1)).Find "", , , xlPart ' F4.1新增
        Range(Cells(c, ColumnOrg), Cells(c, ColumnOrg)).Select
        ActiveWindow.ScrollRow = c
    End If
    Application.ScreenUpdating = True
End Sub ' E9.20，E10.17更新
Sub Sort_Number()
Attribute Sort_Number.VB_ProcData.VB_Invoke_Func = " \n14"
    Set_Last
    Set_Org
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnLib), Cells(RowLast, ColumnLib)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnNumber), Cells(RowLast, ColumnNumber)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange Range(Cells(RowStart, ColumnStart), Cells(RowLast, ColumnLast))
        .Header = xlFalse
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Back_Org
End Sub ' E9.16
Sub Sort_NameIn()
    Set_Last
    Set_Org
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnLib), Cells(RowLast, ColumnLib)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnName), Cells(RowLast, ColumnName)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange Range(Cells(RowStart, ColumnStart), Cells(RowLast, ColumnLast))
        .Header = xlFalse
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Back_Org
End Sub ' E9.16
Sub Sort_Name()
    Set_Last
    Set_Org
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnName), Cells(RowLast, ColumnName)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange Range(Cells(RowStart, ColumnStart), Cells(RowLast, ColumnLast))
        .Header = xlFalse
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Back_Org
End Sub ' E9.16
Sub Sort_Composer()
    Set_Last
    Set_Org
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnComposer), Cells(RowLast, ColumnComposer)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(Cells(RowStart, ColumnName), Cells(RowLast, ColumnName)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange Range(Cells(RowStart, ColumnStart), Cells(RowLast, ColumnLast))
        .Header = xlFalse
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Back_Org
End Sub ' E9.19
Public Function Song_Name(Row As Long, Optional Name1 As Boolean = True) As String
    Song_Name = Cells(Row, ColumnName)
    If InStr(Song_Name, Chr(10)) Then
        If Name1 Then
            Song_Name = Left(Song_Name, InStr(Song_Name, Chr(10)) - 1)
        Else
            Song_Name = Mid(Song_Name, InStr(Song_Name, Chr(10)) + 1)
        End If
    End If
End Function ' E9.19
Public Function Song_Dir(Row As Long) As String
    Song_Dir = DirLib & Cells(Row, ColumnLib) & "\" & Name2File(Song_Name(Row))
End Function ' E9.19
Sub Song_Open()
    FileD = Song_Dir(ActiveCell.Row)
    If Len(Dir(FileD)) = 0 Then
        MsgBox ("文件不存在")
    Else
        shell ("""" & DirPlayer & """ """ & FileD & """")
    End If
End Sub ' E9.16
Sub Song_Locate()
    FileD = Song_Dir(ActiveCell.Row)
    If Len(Dir(FileD)) = 0 Then
        MsgBox ("文件不存在")
    Else
        shell ("explorer.exe /select, " & """" & FileD & """")
    End If
End Sub ' E9.16
Sub Song_New()
    Dim CurLib As Long
    Dim CurNumber As Long
    Set_Last
    CurLib = WorksheetFunction.Max(Range(Cells(RowStart, ColumnLib), Cells(RowLast, ColumnLib)))
    CurNumber = 1
    For i = RowStart To RowLast
        If Cells(i, ColumnLib) = CurLib And Cells(i, ColumnNumber) > CurNumber Then CurNumber = Cells(i, ColumnNumber)
    Next i
    RowLast = RowLast + 1
    Range(Cells(RowLast, ColumnLib), Cells(RowLast, ColumnLib)) = CurLib
    Range(Cells(RowLast, ColumnNumber), Cells(RowLast, ColumnNumber)) = CurNumber + 1
    Range(Cells(RowLast, ColumnDate), Cells(RowLast, ColumnDate)) = Chr(Year(Date) - 2011 + Asc("A")) & Format(Date, "m.d")
    Range(Cells(RowLast, ColumnForm), Cells(RowLast, ColumnForm)).Select
End Sub ' E9.19
Sub Song_Check()
    Dim SongN As String
    Dim FileN As String
    Dim i As Long
    SongN = Name2File(Song_Name(ActiveCell.Row))
    FileN = Dir(Song_Dir(ActiveCell.Row))
    If Len(FileN) = 0 Then
        MsgBox ("文件不存在")
        Exit Sub
    End If
    If StrComp(FileN, SongN) Then
        For i = 1 To Len(FileN)
            'MsgBox ("表格：" & Mid(SongN, i, 1) & Asc(Mid(SongN, i, 1)) & vbCrLf & "文件：" & Mid(FileN, i, 1) & Asc(Mid(FileN, i, 1)))
            If Asc(Mid(SongN, i, 1)) <> Asc(Mid(FileN, i, 1)) Then
                If MsgBox("表格：" & SongN & vbCrLf & "文件：" & FileN & vbCrLf & "是否打开文件夹？", vbYesNo, "大小写不匹配") = vbYes Then
                    shell ("explorer.exe " & """" & DirLib & Cells(ActiveCell.Row, ColumnLib) & "\""")
                End If
                Exit Sub
            End If
        Next i
    End If
    MsgBox ("正确匹配")
End Sub ' E9.19
Sub Song_Dur()
    Dim shell
    'With CreateObject("scripting.filesystemobject")
    Set shell = CreateObject("Shell.Application").Namespace(DirLib & Cells(ActiveCell.Row, ColumnLib) & "\")
    Range(Cells(ActiveCell.Row, ColumnDur), Cells(ActiveCell.Row, ColumnDur)) = _
        shell.GetDetailsOf(shell.Items.Item(Name2File(Song_Name(ActiveCell.Row))), 27)
    Set shell = Nothing
    'End With
End Sub ' E9.19
Sub Song_Rename()
    FileD = Song_Dir(ActiveCell.Row)
    If Len(Dir(FileD)) = 0 Then
        MsgBox ("文件不存在")
    Else
        RenameBox.Show (vbModal)
    End If
End Sub ' E9.20
Sub CheckAll()
    Dim SongN As String
    Dim FileN As String
    Dim i As Long
    Dim j As Long
    Set_Last
    For i = ActiveCell.Row To RowLast
        SongN = Name2File(Song_Name(i))
        FileN = Dir(Song_Dir(i))
        If Len(FileN) = 0 Then
            Range(Cells(i, ColumnName), Cells(i, ColumnName)).Select
            MsgBox (SongN & " 不存在")
            Exit Sub
        End If
        If StrComp(FileN, SongN) Then
            For j = 1 To Len(FileN)
                If Asc(Mid(SongN, j, 1)) <> Asc(Mid(FileN, j, 1)) Then
                    Range(Cells(i, ColumnName), Cells(i, ColumnName)).Select
                    If MsgBox("表格：" & SongN & vbCrLf & "文件：" & FileN & vbCrLf & "是否打开文件夹？", vbYesNo, "大小写不匹配") = vbYes Then
                        shell ("explorer.exe " & """" & DirLib & Cells(ActiveCell.Row, ColumnLib) & "\""")
                    End If
                Exit Sub
                End If
            Next j
        End If
    Next i
    MsgBox ("全部正确匹配")
End Sub ' E9.19
Sub ResetFormat()
    Dim Backup As Range
    Set Backup = Selection
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
    End With
    Backup.Select
End Sub ' E9.19
Sub AllDur()
    Dim shell
    Dim i As Long
    If MsgBox("确认执行？", vbOKCancel Or vbDefaultButton2) <> vbOK Then Exit Sub
    Set_Last
    For i = ActiveCell.Row To RowLast
        Set shell = CreateObject("Shell.Application").Namespace(DirLib & Cells(i, ColumnLib) & "\")
        Range(Cells(i, ColumnDur), Cells(i, ColumnDur)) = _
            shell.GetDetailsOf(shell.Items.Item(Name2File(Song_Name(i))), 27)
        Set shell = Nothing
    Next i
End Sub ' E9.19

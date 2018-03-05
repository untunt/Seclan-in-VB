VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEditor 
   Caption         =   "表达式编辑器"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12150
   OleObjectBlob   =   "FormEditor.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FormEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet
Dim Dict As Worksheet
Dim DictRows As Long
Dim Row As Long
Dim OrgName As String
Private Function hToF(Str As String) As String
    hToF = Replace(Replace(Replace(Replace(Str, "|", "｜"), "/", "／"), ")", "）"), "(", "（")
End Function
Private Function fToH(Str As String) As String
    fToH = Replace(Replace(Replace(Replace(Str, "／", "/"), "｜", "|"), "）", ")"), "（", "(")
End Function
Private Sub ClickOK()
    CommandButton4_Click
    ws.Cells(Row, ColumnForm) = fToH(TextBox1)
    If Mono = "~" Then
        ws.Cells(Row, ColumnName) = TextBox2
    ElseIf Mono = "`" Then
        ws.Cells(Row, ColumnName) = TextBox3
    Else
        ws.Cells(Row, ColumnName) = TextBox2 & Chr(10) & TextBox3
    End If
End Sub
Sub SetBoxVis()
    If Mono = "~" Then
        Label2.Visible = True
        TextBox2.Visible = True
        Label3.Visible = False
        TextBox3.Visible = False
    ElseIf Mono = "`" Then
        Label2.Visible = False
        TextBox2.Visible = False
        Label3.Visible = True
        TextBox3.Visible = True
    Else
        Label2.Visible = True
        TextBox2.Visible = True
        Label3.Visible = True
        TextBox3.Visible = True
    End If
End Sub
Private Sub CommandButton1_Click()
    ClickOK
    End
End Sub
Private Sub CommandButton2_Click()
    ClickOK
    'MsgBox (OrgName & vbCrLf & TextBox2 & vbCrLf & TextBox3)
    'MsgBox (DirLib & Cells(Row, ColumnLib) & "\" & Name2File(OrgName))
    If OrgName = "" Then End
    If Mono = "`" And TextBox3 <> OrgName Then
        Name DirLib & Cells(Row, ColumnLib) & "\" & Name2File(OrgName) As DirLib & Cells(Row, ColumnLib) & "\" & Name2File(TextBox3)
    ElseIf TextBox2 <> OrgName Then
        Name DirLib & Cells(Row, ColumnLib) & "\" & Name2File(OrgName) As DirLib & Cells(Row, ColumnLib) & "\" & Name2File(TextBox2)
    End If
    End
End Sub
Private Sub CommandButton3_Click()
    End
End Sub
Private Sub CommandButton4_Click()
    Form_ReadAll (fToH(TextBox1))
    TextBox2 = Str1
    TextBox3 = Str2
    SetBoxVis
End Sub
Private Sub CommandButton5_Click()
    TextBox1.SelText = hToF(ListBox.List(ListBox.ListIndex))
    TextBox1.SetFocus
End Sub
Private Sub CommandButton6_Click()
    Dim i As Long
    For i = 0 To ListBox.ListCount - 1
        If Asc(Left(ListBox.List(i), 1)) < 128 And Asc(Left(ListBox.List(i), 1)) > 32 And _
            UCase(TextBox1.SelText) < UCase(ListBox.List(i)) Then Exit For
    Next i
    ListBox.Selected(i) = True
End Sub
Private Sub CommandButton7_Click()
    DictInsert.Show (vbModel)
End Sub
Private Sub CommandButton8_Click()
    TextBox1.SelText = hToF(LCase(ListBox.List(ListBox.ListIndex)))
End Sub
Private Sub CommandButton9_Click()
    Form_ReadAll (fToH(TextBox1))
    If Mono = "~" Then
        If Str1 = TextBox2 Then
            MsgBox ("匹配")
        Else
            MsgBox ("不匹配")
        End If
    ElseIf Mono = "`" Then
        If Str2 = TextBox3 Then
            MsgBox ("匹配")
        Else
            MsgBox ("不匹配")
        End If
    Else
        If Str1 = TextBox2 And Str2 = TextBox3 Then
            CommandButton1_Click
        Else
            MsgBox ("不匹配")
        End If
    End If
End Sub
Private Sub ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton5_Click
End Sub
Private Sub TabStrip_Change()
    LoadList
End Sub
Sub LoadList()
    DictRows = Dict.UsedRange.Rows.Count
    ListBox.Clear
    For i = 2 To DictRows
        If Dict.Cells(i, 1) = TabStrip.SelectedItem.Name Then
            For j = 4 To 7
                If Dict.Cells(i, j) <> "" Then ListBox.AddItem (Dict.Cells(i, 2) & "/" & Dict.Cells(i, j))
            Next j
        End If
    Next i
End Sub
Sub Dict_Sort()
    DictRows = Dict.UsedRange.Rows.Count
    Dict.Sort.SortFields.Clear
    Dict.Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(DictRows, 1)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Dict.Sort.SortFields.Add Key:=Range(Cells(2, 2), Cells(DictRows, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Dict.Sort
        .SetRange Range(Cells(2, 1), Cells(DictRows, 8))
        .Header = xlFalse
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Private Sub TextBox1_Change()
    Dim a As Long
    a = TextBox1.SelStart
    TextBox1 = hToF(TextBox1)
    TextBox1.SelStart = a
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then
        If TextBox1.SelText = "" Then
            CommandButton9_Click
            'TextBox1.SetFocus
        Else
            CommandButton6_Click
        End If
    End If
End Sub
Private Sub UpdateLabel()
    Dim a As Long
    Dim n As Long
    Dim i As Long
    Dim p(3) As Long
    Dim ps(3) As Long
    Dim s(3) As String
    a = TextBox1.SelStart
    n = 0
    For i = 3 To a
        If Mid(TextBox1, i, 1) = "（" Then
            If n >= 0 And n <= 3 Then
                p(n) = i - 2
                ps(n) = 0
            End If
            n = n + 1
        ElseIf Mid(TextBox1, i, 1) = "｜" Then
            If n > 0 And n <= 4 Then ps(n - 1) = ps(n - 1) + 1
        ElseIf Mid(TextBox1, i, 1) = "）" Then
            n = n - 1
        End If
    Next i
    If n > 0 Then
        Label1 = hToF(Form_Conv(Mid(TextBox1, p(n - 1), 2), 10 + ps(n - 1), s))
        If Label1 = "／" Then Label1 = ""
    ElseIf n = 0 And a > 1 Then
        Label1 = hToF(Form_Conv(Mid(TextBox1, a - 1, 2), 9, s))
        If Label1 = "／" Then Label1 = ""
    Else
        Label1 = ""
    End If
    If Left(Label1, 2) = "tp" And n > 0 Then
        TabStrip.Value = 3
    End If
    'If n >= 0 Then Label1 = n & "|0:" & p(0) & "-" & ps(0) & "|1:" & p(1) & "-" & ps(1) & "|2:" & p(2) & "-" & ps(2) & "|3:" & p(3) & "-" & ps(3)
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UpdateLabel
End Sub
Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    UpdateLabel
End Sub
Private Sub UserForm_Initialize()
    Dim i As Long
    Dim j As Long
    Set Dict = Worksheets(SheetDict)
    Set ws = Worksheets(SheetName)
    Row = ActiveCell.Row
    If Cells(Row, ColumnName) <> "" Then
        If InStr(Cells(Row, ColumnName), Chr(10)) Then
            TextBox2 = Song_Name(Row, True)
            TextBox3 = Song_Name(Row, False)
            Mono = ""
        Else
            TextBox2 = Song_Name(Row)
            TextBox3 = Song_Name(Row)
            Mono = "~"
            Label3.Visible = False
            TextBox3.Visible = False
        End If
        If Cells(Row, ColumnForm) = "" Then
            TextBox1 = hToF(Str_Comb(TextBox2, TextBox3))
            OrgName = TextBox2
        Else
            TextBox1 = hToF(Cells(Row, ColumnForm))
            If Left(TextBox1, 1) = "~" Then
                Mono = "~"
                OrgName = TextBox2
            ElseIf Left(TextBox1, 1) = "`" Then
                Mono = "`"
                OrgName = TextBox3
            Else
                Mono = ""
                OrgName = TextBox2
            End If
            SetBoxVis
        End If
    End If
    Dict_Sort
    TabStrip.Tabs.Clear
    TabStrip.Tabs.Add "set", "曲集"
    TabStrip.Tabs.Add "form", "曲式"
    TabStrip.Tabs.Add "dance", "舞曲"
    TabStrip.Tabs.Add "tempo", "速度"
    TabStrip.Tabs.Add "inst", "乐器"
    LoadList
End Sub

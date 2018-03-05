VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DictInsert 
   Caption         =   "插入名词"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "DictInsert.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DictInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Boolean
Private Sub CommandButton1_Click()
    Dim Dict As Worksheet
    Dim DictRows As Long
    Set Dict = Worksheets(SheetDict)
    DictRows = Dict.UsedRange.Rows.Count
    
    For i = 2 To DictRows
        If Dict.Range(Dict.Cells(i, 2), Dict.Cells(DictRows, 2)).Find(TextBox2, , , xlWhole) Is Nothing Then
            DictRows = DictRows + 1
            Dict.Cells(DictRows, 1) = TextBox1
            Dict.Cells(DictRows, 2) = TextBox2
            Dict.Cells(DictRows, 4) = TextBox3
            FormEditor.Dict_Sort
            Exit For
        End If
        i = Dict.Range(Dict.Cells(i, 2), Dict.Cells(DictRows, 2)).Find(TextBox2, , , xlWhole).Row
        If Dict.Cells(i, 1) = TextBox1 Then
            For j = 4 To 7
                If Dict.Cells(i, j) = TextBox3 Then
                    MsgBox ("已存在")
                    Exit For
                End If
                If Dict.Cells(i, j) = "" Then
                    Dict.Cells(i, j) = TextBox3
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
    FormEditor.LoadList
    FormEditor.TextBox1.SetFocus
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then CommandButton1_Click
End Sub
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then CommandButton1_Click
End Sub
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then CommandButton1_Click
End Sub
Private Sub UserForm_Activate()
    If f Then Exit Sub
    If TextBox2 <> "" And TextBox3 = "" Then
        TextBox3.SetFocus
    Else
        TextBox2.SetFocus
    End If
    f = True
End Sub
Private Sub UserForm_Initialize()
    TextBox1 = FormEditor.TabStrip.SelectedItem.Name
    If InStr(FormEditor.TextBox1.SelText, "／") Then
        TextBox2 = Left(FormEditor.TextBox1.SelText, InStr(FormEditor.TextBox1.SelText, "／") - 1)
        TextBox3 = Mid(FormEditor.TextBox1.SelText, InStr(FormEditor.TextBox1.SelText, "／") + 1)
    Else
        TextBox2 = FormEditor.TextBox1.SelText
    End If
    f = False
End Sub

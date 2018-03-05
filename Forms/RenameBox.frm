VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameBox 
   Caption         =   "重命名"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   OleObjectBlob   =   "RenameBox.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "RenameBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' E9.20~21
Dim Row As Long
Dim OrgName As String
Private Sub CheckBox_Click()
    Label2.Visible = Not CheckBox
    TextBox2.Visible = Not CheckBox
End Sub
Private Sub CommandButton1_Click()
    If TextBox1 <> OrgName Then Name DirLib & Cells(Row, ColumnLib) & "\" & Name2File(Song_Name(Row)) As DirLib & Cells(Row, ColumnLib) & "\" & Name2File(TextBox1)
    If CheckBox Then
        Range(Cells(Row, ColumnName), Cells(Row, ColumnName)) = TextBox1
    Else
        Range(Cells(Row, ColumnName), Cells(Row, ColumnName)) = TextBox1 & Chr(10) & TextBox2
    End If
    End
End Sub
Private Sub CommandButton2_Click()
    End
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then CommandButton1_Click
End Sub
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = 13 Then CommandButton1_Click
End Sub
Private Sub UserForm_Initialize()
    Row = ActiveCell.Row
    If InStr(Cells(Row, ColumnName), Chr(10)) Then
        TextBox1 = Song_Name(Row, True)
        TextBox2 = Song_Name(Row, False)
    Else
        TextBox1 = Song_Name(Row)
        CheckBox = True
        Label2.Visible = False
        TextBox2.Visible = False
    End If
    OrgName = TextBox1
End Sub

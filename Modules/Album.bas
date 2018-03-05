Attribute VB_Name = "Album"
Option Explicit
Public Const St = 3

Sub Alb_Split() ' E12.31
    Dim RowC As Long
    Dim Pos As Long
    Dim Pos2 As Long
    Dim NameS As String
    Dim PerfS As String
    RowC = ActiveCell.Row
    NameS = Replace(Cells(RowC, St + 4), ": ", "- ")
    
    ' 拆分作曲家
    Pos = InStr(NameS, "- ")
    Pos2 = InStr(NameS, "：")
    If Pos = 0 And Pos2 = 0 Then End
    If Pos = 1 Then
        Cells(RowC, St + 3) = ":"
        NameS = Mid(NameS, 3)
    Else
        If Pos2 = 0 Then
            If Mid(NameS, Pos - 1, 1) = " " Then End
            Cells(RowC, St + 3) = Left(NameS, Pos - 1) & ":"
            NameS = Mid(NameS, Pos + 2)
        Else
            Cells(RowC, St + 3) = Left(NameS, Pos2 - 1) & "："
            NameS = Mid(NameS, Pos2 + 1)
        End If
    End If
    
    ' 拆分标签
    Pos = InStrRev(NameS, " [")
    If Pos > 0 Then
        Cells(RowC, St + 7) = Mid(NameS, Pos + 1)
        NameS = Left(NameS, Pos - 1)
    End If
    
    ' 拆分日期地点
    For Pos = Len(NameS) - 3 To 1 Step -1
        If IsNumeric(Mid(NameS, Pos, 1)) And IsNumeric(Mid(NameS, Pos + 1, 1)) And IsNumeric(Mid(NameS, Pos + 2, 1)) And IsNumeric(Mid(NameS, Pos + 3, 1)) Then
            Cells(RowC, St + 6) = Mid(NameS, Pos)
            NameS = Left(NameS, Pos - 2)
            Exit For
        End If
    Next Pos
    
    ' 拆分演奏者
    Pos = InStrRev(NameS, " - ")
    If Pos > 0 Then
        PerfS = Mid(NameS, Pos + 3)
        NameS = Left(NameS, Pos - 1)
    End If
    
    Cells(RowC, St + 4) = Replace(Replace(Replace(Replace(NameS, " - ", "*-*"), "- ", ": "), "*-*", " - "), "_", "/")
    Cells(RowC, St + 5) = PerfS
    If Cells(RowC, St + 1) = "" Then
        If Cells(RowC, 1) = "" Then
            Cells(RowC, St + 1) = Chr(Year(Date) - 2011 + Asc("A")) & Format(Date, "m.d")
        Else
            Cells(RowC, St + 1) = DateConv(RowC)
        End If
    End If
    Range(Cells(RowC + 1, ActiveCell.Column), Cells(RowC + 1, ActiveCell.Column)).Select
End Sub

Function DateConv(RowC As Long) As String  '转换时间予下载日期 F1.29
    Dim dt As String
    Dim m As Integer
    Dim d As Integer
    dt = Left(Cells(RowC, 1), 5)
    m = Mid(dt, 2, 2)
    d = Right(dt, 2)
    DateConv = Chr(Left(dt, 1) - 1 + Asc("A")) & m & "." & d
End Function

Sub TotalTime() ' F4.3～4.4°
    Dim i As Long
    Dim last As Long
    Dim tt As Double
    Dim tp As Double
    last = ActiveSheet.UsedRange.Rows.Count
    tt = 0
    For i = 2 To last
        tt = tt + Cells(i, 12)
        If Cells(i, 5) <> "" And Right(Cells(i, 5), 1) <> "~" Then tp = tp + Cells(i, 12)
    Next i
    MsgBox "专辑总时长：" & Int(tt) & " 天 " & Format(tt, "h:nn:ss") & vbCrLf & _
           "已听完时长：" & Int(tp) & " 天 " & Format(tp, "h:nn:ss")
End Sub

Attribute VB_Name = "Formula"
' E10.4开始
Public Mono As String
Public Str0 As String
Public Str1 As String
Public Str2 As String
Dim L As Long
Public FuncLast As String
Function NumR(n As Integer) As String
    NumR = Application.WorksheetFunction.Roman(n)
End Function
Function NumC(n As Integer) As String
    Dim Str As Variant
    Str = Array("一", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    If n < 11 Then
        NumC = Str(n)
    ElseIf n < 20 Then
        NumC = Str(10) & Str(n - 10)
    Else
        NumC = Str(n / 10) & Str(10) & Str(n Mod 10)
    End If
End Function
Sub Str_Div(Str As String, Div1 As String, Div2 As String)
    Str = Replace(Str, "//", "\")
    If Mono = "~" Then
        Div1 = Str
        Div2 = ""
    ElseIf Mono = "`" Then
        Div1 = ""
        Div2 = Str
    ElseIf InStr(Str, "/") > 0 Then
        Div1 = Left(Str, InStr(Str, "/") - 1)
        Div2 = Mid(Str, InStr(Str, "/") + 1)
    Else
        MsgBox ("缺少“/”")
    End If
    Str = Replace(Str, "\", "//")
    Div1 = Replace(Div1, "\", "//")
    Div2 = Replace(Div2, "\", "//")
End Sub
Function Str_Comb(Div1 As String, Div2 As String) As String
    If Mono = "~" Then
        Str_Comb = Div1
    ElseIf Mono = "`" Then
        Str_Comb = Div2
    Else
        Str_Comb = Div1 & "/" & Div2
    End If
End Function
Function Form_Conv(FuncName As String, ArgN As Long, Arg() As String) As String
    Dim Conv1 As String
    Dim Conv2 As String
    Dim Tmp1 As String
    Dim Tmp2 As String
    Dim i As Long
    Select Case FuncName
    Case "st"
        If ArgN > 8 Then
            Form_Conv = "st(数字|名称)" & vbCrLf & "曲集/set"
            Exit Function
        End If
        Str_Div Arg(1), Conv1, Conv2
        Conv1 = Arg(0) & " " & Conv1
        Conv2 = Arg(0) & "首" & Conv2
    Case "ke"
        If ArgN > 8 Then
            Select Case ArgN
            Case 9
                Form_Conv = "ke(调名|[调式]|[名称])" & vbCrLf & "调性/key"
            Case 10
                Form_Conv = "ke＞调名|[调式]|[名称])" & vbCrLf & "调性/key"
            Case 11
                Form_Conv = "ke(调名＞[调式]|[名称])" & vbCrLf & "1=大调，0=小调，空=调"
            Case 12
                Form_Conv = "ke(调名|[调式]＞[名称])" & vbCrLf & "调性/key"
            End Select
            Exit Function
        End If
        If Len(Arg(0)) = 2 Then
            If Right(Arg(0), 1) = "#" Then
                Conv1 = "-sharp"
                Conv2 = "升"
            Else
                Conv1 = "-flat"
                Conv2 = "降"
            End If
        End If
        If Arg(1) = "1" Then
            Conv1 = UCase(Left(Arg(0), 1)) & Conv1 & " major"
            Conv2 = Conv2 & UCase(Left(Arg(0), 1)) & "大调"
        ElseIf Arg(1) = "0" Then
            Conv1 = UCase(Left(Arg(0), 1)) & Conv1 & " minor"
            Conv2 = Conv2 & LCase(Left(Arg(0), 1)) & "小调"
        Else
            Conv1 = UCase(Left(Arg(0), 1)) & Conv1
            Conv2 = Conv2 & UCase(Left(Arg(0), 1)) & "调"
        End If
        Str_Div Arg(2), Tmp1, Tmp2
        Conv1 = Tmp1 & " in " & Conv1
        Conv2 = Conv2 & Tmp2
    Case "no"
        If ArgN > 8 Then
            Form_Conv = "no(数字|[体裁]|[中文用汉字=0])" & vbCrLf & "编号/number"
            Exit Function
        End If
        If ArgN = 0 Then
            Conv1 = "No. " & Arg(0)
            Conv2 = "第" & Arg(0) & "号"
        Else
            Str_Div Arg(1), Conv1, Conv2
            Conv1 = Conv1 & " No. " & Arg(0)
            If ArgN = 1 Then
                Conv2 = "第" & Arg(0) & "号" & Conv2
            Else
                Conv2 = "第" & NumC(Val(Arg(0))) & Conv2
            End If
        End If
    Case "na"
        If ArgN > 8 Then
            Form_Conv = "na(名称)" & vbCrLf & "名称/name"
            Exit Function
        End If
        Str_Div Arg(0), Conv1, Conv2
    Case "op"
        If ArgN > 8 Then
            Form_Conv = "op(数字|[作品号名]|[子作品号])" & vbCrLf & "作品号/opus"
            Exit Function
        End If
        If ArgN = 0 Or Arg(1) = "" Then Arg(1) = "Op."
        Conv1 = ", " & Arg(1) & " " & Arg(0)
        Conv2 = Arg(1) & " " & Arg(0)
        If FuncLast = "op" Then Conv2 = "，" & Conv2 ' G10.21
        If ArgN = 2 Then
            If Arg(1) = "Op." Then
                Conv1 = Conv1 & " No. " & Arg(2)
                Conv2 = Conv2 & "第" & Arg(2) & "号"
            Else
                Conv1 = Conv1 & "//" & Arg(2)
                Conv2 = Conv2 & "//" & Arg(2)
            End If
        End If
    Case "ti"
        If ArgN > 8 Then
            Select Case ArgN
            Case 9
                Form_Conv = "ti(名称|[类型])" & vbCrLf & "标题/title"
            Case 10
                Form_Conv = "ti＞名称|[类型])" & vbCrLf & "标题/title"
            Case 11
                Form_Conv = "ti(名称＞[类型])" & vbCrLf & "ti=标题 al=别名 ly=歌词 ex=解释 tp=速度"
                Select Case Mid(FormEditor.TextBox1, FormEditor.TextBox1.SelStart - 1, 2)
                Case "ti"
                    Form_Conv = "ti(名称＞[类型])" & vbCrLf & "标题/title：引号 用于古典作品"
                Case "al"
                    Form_Conv = "ti(名称＞[类型])" & vbCrLf & "别名/alternative：括号 用于歌剧片段"
                Case "ly"
                    Form_Conv = "ti(名称＞[类型])" & vbCrLf & "歌词/lyrics：冒号 用于歌剧唱段"
                Case "ex"
                    Form_Conv = "ti(名称＞[类型])" & vbCrLf & "解释/explain：冒号 用于歌剧场景"
                Case "tp"
                    Form_Conv = "ti(名称＞[类型])" & vbCrLf & "速度表情/tempo：点 用于乐章分曲"
                End Select
            End Select
            Exit Function
        End If
        If ArgN = 0 Then Arg(1) = "ti"
        Str_Div Arg(0), Conv1, Conv2
        Select Case Arg(1)
        Case "ti"
            Conv1 = " '" & Conv1 & "'"
            Conv2 = "“" & Conv2 & "”"
            If FuncLast <> "op" Then Conv1 = "," & Conv1
        Case "al"
            Conv1 = " (" & Conv1 & ")"
            Conv2 = "（" & Conv2 & "）"
        Case "ly"
            Conv1 = ": " & Conv1
            Conv2 = "：" & Conv2
        Case "ex"
            Conv1 = ": " & Conv1
            Conv2 = "：" & Conv2
        Case "tp"
            Conv1 = ". " & Conv1
            Conv2 = "，" & Conv2
        End Select
    Case "sp"
        If ArgN > 8 Then
            Select Case ArgN
            Case 9
                Form_Conv = "sp([类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "分曲/separate"
            Case 10
                Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "ac/pt=幕 bk=集 sc=景 mv=乐章 su=组曲 st=曲集 op=剧 符号"
                Select Case Mid(FormEditor.TextBox1, FormEditor.TextBox1.SelStart - 1, 2)
                Case "ac"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "幕/Act：逗号＋罗马"
                Case "pt"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "部分/Part：逗号＋罗马"
                Case "bk"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "集/Book：逗号＋阿拉伯"
                Case "sc"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "景/Scene：冒号＋阿拉伯，中文用汉字（另：sca、sct、scr）"
                Case "ca"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "景/Scene：冒号＋阿拉伯，中文用数字"
                Case "ct"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "场/Scene：冒号＋阿拉伯"
                Case "cr" ' E12.31
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "景/Scene：冒号＋罗马"
                Case "mv"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "乐章/Movement：罗马 用于交响曲、协奏曲、奏鸣曲等"
                Case "su"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "组曲/Suite：罗马 用于有主题、曲式关联的组曲（另：suc、suu）"
                Case "uc"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "组曲/Suite：罗马 用于有结构、构思关联的组曲（行星组曲）"
                Case "uu"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "组曲/Suite：阿拉伯 用于无关联的组曲（图画展览会）"
                Case "st"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "曲集/Set：阿拉伯（另：stn）"
                Case "tn"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "曲集/Set：No.＋阿拉伯（×首××）"
                Case "op"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "数码剧/Opera：No.＋阿拉伯（另：stn）"
                Case "pa"
                    Form_Conv = "sp＞[类型]|[数字]|[名词]|[使用连接符=1])" & vbCrLf & "数码剧/Opera：阿拉伯（Cantata等）"
                End Select
            Case 11
                Form_Conv = "sp([类型]＞[数字]|[名词]|[使用连接符=1])" & vbCrLf & "分曲/separate"
            Case 12
                Form_Conv = "sp([类型]|[数字]＞[名词]|[使用连接符=1])" & vbCrLf & "分曲/separate"
            Case 13
                Form_Conv = "sp([类型]|[数字]|[名词]＞[使用连接符=1])" & vbCrLf & "分曲/separate"
            End Select
            Exit Function
        End If
        Select Case Arg(0)
        Case "ac"
            If ArgN = 1 Then Arg(2) = "Act/幕"
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ", " & Conv1 & " " & NumR(Val(Arg(1)))
            Conv2 = "，第" & NumC(Val(Arg(1))) & Conv2
        Case "pt"
            If ArgN = 1 Then Arg(2) = "Part/部分"
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ", " & Conv1 & " " & NumR(Val(Arg(1)))
            Conv2 = "，第" & NumC(Val(Arg(1))) & Conv2
        Case "bk"
            If ArgN = 1 Then Arg(2) = "Book/集"
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ", " & Conv1 & " " & Arg(1)
            Conv2 = "，第" & Arg(1) & Conv2
        Case "sc"
            If ArgN = 1 Then Arg(2) = "Scene/景" ' E12.31
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ": " & Conv1 & " " & Arg(1)
            Conv2 = "：第" & NumC(Val(Arg(1))) & Conv2
        Case "sca"
            If ArgN = 1 Then Arg(2) = "Scene/景" ' F4.1
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ": " & Conv1 & " " & Arg(1)
            Conv2 = "：第" & Arg(1) & Conv2
        Case "sct"
            Conv1 = ": Scene " & Arg(1)
            Conv2 = "：第" & NumC(Val(Arg(1))) & "场"
        Case "scr" ' E12.31
            Str_Div Arg(2), Conv1, Conv2
            Conv1 = ": " & Conv1 & " " & NumR(Val(Arg(1)))
            Conv2 = "：第" & NumC(Val(Arg(1))) & Conv2
        Case "mv"
            Conv1 = ": " & NumR(Val(Arg(1))) & ". "
            Conv2 = "：" & NumC(Val(Arg(1))) & "、"
        Case "su"
            If Arg(3) = "0" Then
                Conv1 = NumR(Val(Arg(1))) & ". "
                Conv2 = NumC(Val(Arg(1))) & "、"
            Else
                Conv1 = ": " & NumR(Val(Arg(1))) & ". "
                Conv2 = "：" & NumC(Val(Arg(1))) & "、"
            End If
        Case "suc"
            Conv1 = ": " & NumR(Val(Arg(1))) & ". "
            Conv2 = "：" & NumC(Val(Arg(1))) & "、"
        Case "suu"
            If ArgN = 1 Then
                Conv1 = ": " & Arg(1) & ". "
                Conv2 = "：" & Arg(1) & "、"
            Else
                Str_Div Arg(2), Conv1, Conv2
                Conv1 = ": " & Conv1 & " " & Arg(1) & ". "
                Conv2 = "：" & Conv2 & Arg(1) & "、"
            End If
        Case "st"
            Conv1 = ": " & Arg(1) & ". "
            Conv2 = "：" & Arg(1) & "、"
        Case "stn"
            Conv1 = ": No. " & Arg(1) & " "
            Conv2 = "：" & Arg(1) & "、"
        Case "op"
            Conv1 = ": No. " & Arg(1) & " "
            Conv2 = "：" & Arg(1) & "、"
        Case "opa"
            Conv1 = ": " & Arg(1) & ". "
            Conv2 = "：" & Arg(1) & "、"
        Case ""
            Conv1 = ": "
            Conv2 = "："
        Case ","
            Conv1 = ", "
            Conv2 = "，"
        Case " "
            Conv1 = " "
            Conv2 = ""
        End Select
        If Arg(3) = "0" And (Left(Conv1, 1) = ":" Or Left(Conv1, 1) = ":") Then ' revised from above at E11.7
            Conv1 = Mid(Conv1, 3)
            Conv2 = Mid(Conv2, 2)
        End If
    Case "jc"
        If ArgN > 8 Then
            Select Case ArgN
            Case 9
                Form_Conv = "jc([模式=0])" & vbCrLf & "连接/junction"
            Case 10
                Form_Conv = "jc＞[模式=0])" & vbCrLf & "0=片段，1=乐章，2=唱段"
            End Select
            Exit Function
        End If
        If Arg(0) = "" Or Arg(0) = "0" Then
            Conv1 = " – "
            Conv2 = "－"
        ElseIf Arg(0) = "1" Then
            Conv1 = " & "
            Conv2 = "＆"
        Else
            Conv1 = " ... "
            Conv2 = "……"
        End If
    Case "tp"
        If ArgN > 8 Then
            Form_Conv = "tp(速度|[前加点=0])" & vbCrLf & "速度/tempo"
            Exit Function
        End If
        Str_Div Arg(0), Conv1, Conv2
        If Arg(1) = "1" Then
            Conv1 = ". " & Conv1
            Conv2 = "，" & Conv2
        End If
    Case "ar"
        If ArgN > 8 Then
            Form_Conv = "ar([改编者])" & vbCrLf & "改编/arrange（另：or）"
            Exit Function
        End If
        If Arg(0) = "" Then
            Conv1 = " (arr.)"
            Conv2 = "（改编）"
        Else
            Str_Div Arg(0), Conv1, Conv2
            Conv1 = " (arr. " & Conv1 & ")"
            Conv2 = "（" & Conv2 & "改编）"
        End If
    Case "or"
        If ArgN > 8 Then
            Form_Conv = "or([改编者])" & vbCrLf & "管弦乐版/orchestrate（另：ar）"
            Exit Function
        End If
        If Arg(0) = "" Then
            Conv1 = " (orch.)"
            Conv2 = "（管弦乐版）"
        Else
            Str_Div Arg(0), Conv1, Conv2
            Conv1 = " (orch. " & Conv1 & ")"
            Conv2 = "（" & Conv2 & "管弦乐版）"
        End If
    Case "nt"
        If ArgN > 8 Then
            Form_Conv = "nt(内容|[前有空格=1])" & vbCrLf & "注释/note"
            Exit Function
        End If
        If ArgN = 0 Then Arg(1) = 1 ' F4.23 增加
        Str_Div Arg(0), Conv1, Conv2
        If Left(Conv1, 2) = ", " Then Conv1 = Mid(Conv1, 3) ' E12.15 (v2.2) 增加
        If Left(Conv2, 1) = "，" Then Conv2 = Mid(Conv2, 2) ' E12.15 (v2.2) 增加
        Conv1 = "(" & Conv1 & ")"
        Conv2 = "（" & Conv2 & "）"
        If Arg(1) = 1 Then Conv1 = " " & Conv1 ' F4.23 增加
    ' 以上为2.1(2.0)版已有的函数
    Case "cb"
        If ArgN > 8 Then
            Form_Conv = "cb(片段0|片段1|[片段2]|[片段3])" & vbCrLf & "字符串拼接/combine"
            Exit Function
        End If
        Str_Div Arg(0), Conv1, Conv2
        For i = 1 To ArgN
            Str_Div Arg(i), Tmp1, Tmp2
            Conv1 = Conv1 & Tmp1
            Conv2 = Conv2 & Tmp2
        Next i
    Case "rv" ' E12.15 (v2.2)
        If ArgN > 8 Then
            Form_Conv = "rv(片段0|片段1)" & vbCrLf & "字符串颠倒/reverse"
            Exit Function
        End If
        Str_Div Arg(0), Conv1, Conv2
        Str_Div Arg(1), Tmp1, Tmp2
        Conv1 = Tmp1 & Conv1
        Conv2 = Tmp2 & Conv2
    Case "fo" ' F3.24 (v2.3)，F8.30增加Arg(2)=2
        If ArgN > 8 Then
            If ArgN = 12 Then
                Form_Conv = "fo(曲式|乐器＞[现代式=0])" & vbCrLf & "0=无中文翻译，1=为……而作的，2=无for"
            Else
                Form_Conv = "fo(曲式|乐器|[现代式=0])" & vbCrLf & "为……而作的/for"
            End If
            Exit Function
        End If
        Str_Div Arg(0), Conv1, Conv2
        Str_Div Arg(1), Tmp1, Tmp2
        If ArgN = 1 Then Arg(2) = 0
        If Arg(2) = 0 Then
            Conv1 = Conv1 & " for " & Tmp1
            Conv2 = Tmp2 & Conv2
        ElseIf Arg(2) = 1 Then
            Conv1 = Conv1 & " for " & Tmp1
            Conv2 = "为" & Tmp2 & "而作的" & Conv2
        Else
            If Left(Tmp1, 1) = " " Then
                Conv1 = Conv1 & Tmp1
            Else
                Conv1 = Conv1 & " " & Tmp1
            End If
            Conv2 = Tmp2 & Conv2
        End If
    End Select
    Form_Conv = Str_Comb(Conv1, Conv2)
    FuncLast = FuncName
End Function
Function Form_Read(Start As Long, Optional SubFunc As Boolean = False, Optional Conv As String) As Long
    Dim FuncName As String
    Dim StrTmp As String
    Dim StrTmp1 As String
    Dim StrTmp2 As String
    Dim Arg(3) As String
    Dim ArgN As Long
    Dim i As Long
    
    FuncName = Mid(Str0, Start, 2)
    Start = Start + 3
    ArgN = 0
    For i = Start To L
        StrTmp = Mid(Str0, i, 1)
        If StrTmp = ")" Then
            Arg(ArgN) = Mid(Str0, Start, i - Start)
            Exit For
        ElseIf StrTmp = "|" Then
            Arg(ArgN) = Mid(Str0, Start, i - Start)
            ArgN = ArgN + 1
            Start = i + 1
        ElseIf StrTmp = "(" Then
            i = Form_Read(i - 2, True, Arg(ArgN))
            If Mid(Str0, i, 1) = ")" Then
                Exit For
            ElseIf Mid(Str0, i, 1) = "|" Then
                ArgN = ArgN + 1
                Start = i + 1 ' E12.15 (v2.2) 增加此句。否则在遇到前一个参数是函数，后一个参数是字符的时候start判断就出错
            End If
        End If
    Next i
    
    If False Then ' 函数监视器 E12.15 添加
        Dim TempStr1 As String
        Dim j As Long
        TempStr1 = FuncName & "(" & ArgN & ")"
        For j = 0 To 3
            TempStr1 = TempStr1 & vbCrLf & "[" & j & "] " & Arg(j)
        Next j
        MsgBox TempStr1
    End If
    
    If SubFunc Then
        Conv = Form_Conv(FuncName, ArgN, Arg)
    Else
        Str_Div Form_Conv(FuncName, ArgN, Arg), StrTmp1, StrTmp2
        Str1 = Str1 & Replace(StrTmp1, "//", "/")
        Str2 = Str2 & Replace(StrTmp2, "//", "/")
    End If
    Form_Read = i + 1
End Function
Public Sub Form_ReadAll(Optional FormStr)
    Dim Start As Long
    If IsMissing(FormStr) Then
        Str0 = Cells(ActiveCell.Row, ColumnForm)
    Else
        Str0 = FormStr
    End If
    Str1 = ""
    Str2 = ""
    L = Len(Str0)
    FuncLast = ""
    
    If Left(Str0, 1) = "~" Then
        Mono = "~"
        Start = 2
    ElseIf Left(Str0, 1) = "`" Then
        Mono = "`"
        Start = 2
    Else
        Mono = ""
        Start = 1
    End If
    Do While Start < L
        Start = Form_Read(Start)
    Loop
End Sub
Sub Form_Test()
    Form_ReadAll
    If Mono = "~" Then
        If (Str1 = Cells(ActiveCell.Row, ColumnName)) Then
            MsgBox (Str1 & vbCrLf & "匹配")
        Else
            MsgBox (Str1 & vbCrLf & "不匹配")
        End If
    ElseIf Mono = "`" Then
        If (Str2 = Cells(ActiveCell.Row, ColumnName)) Then
            MsgBox (Str1 & vbCrLf & "匹配")
        Else
            MsgBox (Str1 & vbCrLf & "不匹配")
        End If
    Else
        If (Str1 & Chr(10) & Str2 = Cells(ActiveCell.Row, ColumnName)) Then
            MsgBox (Str1 & vbCrLf & Str2 & vbCrLf & "匹配")
        Else
            MsgBox (Str1 & vbCrLf & Str2 & vbCrLf & "不匹配")
        End If
    End If
End Sub
Sub Form_Produce()
    Form_ReadAll
    If Mono = "~" Then
        Cells(ActiveCell.Row, ColumnName) = Str1
    ElseIf Mono = "~" Then
        Cells(ActiveCell.Row, ColumnName) = Str2
    Else
        Cells(ActiveCell.Row, ColumnName) = Str1 & Chr(10) & Str2
    End If
End Sub
Sub Form_Box()
    FormEditor.Show (vbModeless)
End Sub ' E9.20

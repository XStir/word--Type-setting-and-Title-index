'====================================删掉选择区域每个段落前面的指定字符,并输入自定义序号====================================
Sub DelSelctedPre()
Dim SelStart, Selend, ParaStart, ParaEnd As Long
Dim m As Integer

SelStart = Selection.Start '获取当前选择区域的起点
Selend = Selection.End '获取当前选择区域的终点

ParaStart = Range(0, SelStart).Paragraphs.Count
ParaEnd = Range(0, Selend).Paragraphs.Count

m = 1
For n = ParaStart To ParaEnd
    Range(Paragraphs(n).Range.Start, Paragraphs(n).Range.Start).Select
    Do While (48 <= Asc(Selection) And Asc(Selection) <= 57) Or Selection = "." Or Selection = " " Or Selection = "?" Or Selection = "(" Or Selection = ")" Or Selection = "、" Or Selection = "[" Or Selection = "]"
        Selection.Delete
    Loop
    Call WriteInPre(m) '不需要写入字符就删掉这行
    m = m + 1
Next n
End Sub

Sub WriteInPre(m) '在每段前写入
Dim kuohao As String
kuohao = "（）"
Selection.TypeText kuohao
Selection.MoveLeft unit:=wdCharacter, Count:=1
Selection.TypeText m
End Sub

'====================================删除选中范围内的空段落====================================
Sub DelBlankPara()
SelStart = Selection.Start '获取当前选择区域的起点
Selend = Selection.End '获取当前选择区域的终点
ParaStart = Range(0, SelStart).Paragraphs.Count
ParaEnd = Range(0, Selend).Paragraphs.Count
For n = ParaStart To ParaEnd
    ActiveDocument.Range(ActiveDocument.Paragraphs(n).Range.Start, ActiveDocument.Paragraphs(n).Range.End).Select
    If Len(Selection) = 1 Then
        Selection.Delete
    End If
Next
End Sub



'====================================检查涂黑选中部分的大纲级别是否存在跨级错误====================================
'如果没有选择任何字符串或者选择部分小于1个段落，自动变为全文检查
Sub check()
Dim now, last As Integer
last = 10
SelStartPara = ActiveDocument.Range(0, Selection.Start).Paragraphs.Count    '起始段落
SelEndPara = ActiveDocument.Range(0, Selection.End).Paragraphs.Count        '终点段落

Dim StartPara, p, j As Long
If SelEndPara - SelStartPara > 1 Then '如果选中部分大于1个段落
    StartPara = SelStartPara
    p = SelEndPara
Else
    StartPara = 1
    p = ActiveDocument.Paragraphs.Count
End If

Do While j <= p
ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.Start).Select
now = Selection.ParagraphFormat.OutlineLevel
If now = 10 Then GoTo nx
If now - last > 1 Then
MsgBox "mistake discovered"
Exit Sub
End If
last = now
nx: j = j + 1
DoEvents
Loop

MsgBox "OK"
End Sub

'====================================整体升or降【选中范围涉及到的段落】的大纲级别====================================
Sub change()
Dim n As Integer
'升降次数，例如n=-2，所有段落大纲级别降低2级
n = 1

Dim SelStartPara, SelEndPara, j As Long
SelStartPara = ActiveDocument.Range(0, Selection.Start).Paragraphs.Count    '起始段落
SelEndPara = ActiveDocument.Range(0, Selection.End).Paragraphs.Count        '终点段落

Dim nowLv, newLv As Integer
For j = SelStartPara To SelEndPara
    ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.Start).Select
    nowLv = Selection.ParagraphFormat.OutlineLevel
    If nowLv < 10 Then
        newLv = nowLv - n
        If newLv >= 10 Then
            MsgBox "降级后的理论级别为" & newLv & "级，实际最小级别为9级，不能满足要求，程序终止。"
            Exit Sub
        End If
        If newLv < 1 Then
            MsgBox "升级后的理论级别为" & newLv & "级，实际最大级别为1级，不能满足要求，程序终止。"
            Exit Sub
        End If
        Selection.ClearFormatting
        Selection.ParagraphFormat.OutlineLevel = newLv
    End If
    DoEvents
Next j
MsgBox "Done"
End Sub
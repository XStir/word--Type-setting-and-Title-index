Dim ApplyStyle_SelAllPra, Title_Len,TypeIndex, Lv1Special As Boolean
Dim Title_Len_Min, Title_Len_Max As Integer
Dim lv1, lv2, lv3, lv4, lv5, lv6to9, Mid_SetTo, ChaAfterIndex, ChaAfterLv1, Main_body_mid, Main_body, Title_Of_Pic_or_Table, Picture, TextInTable, PicTitle, TableTitle, Lv1Index, ParaPreDel_body, ParaPreDel_title, ExsistTypeName As String

Sub ArgSetting()
'========参数-用户设定========
TypeIndex = true
'是否输入标题序号

ChaAfterIndex = " "
'标题序号之后的符号，例如【 】=【1.2.1 标题名字】，【、】=【1.2.1、标题名字】【.】=【1.2.1.标题名字】

ApplyStyle_SelAllPra = False
'套用样式的方案。对段落中部分内容含有多种文字格式的情况有效，比如部分文字加粗、倾斜，字号、颜色不同等。
'false=光标位于段落中，但是没有选中任何文字（len(selection)=0）的情况下套用样式。这样可以保留段落中的部分文字的格式。
'true=选中整个段落再点样式，会清除部分文字的原有格式，将整个段落内文字格式设置成段落样式。
'例子：一个段落，其中有一句话被加粗，另一句话字体字号不同。true将清这些效果，使整个段落的文字都被指定成样式格式；false将保留这些效果，只更改其余部分。

Title_Len = True
Title_Len_Min = 2
Title_Len_Max = 30
'是否启用标题长度限制；上下限的值。
'若启用，段落长度符合 【Title_Len_Min < 段落长度 < Title_Len_Max】才会被认为是段落，否则视为正文。可以防止空段落、分页符、较长的正文被错误地分配了大纲级别，导致后续标题序号错误的情况。

Mid_SetTo = Title_Of_Pic_or_Table
'对无大纲级别、且设置了居中的段落，将分配给哪个样式（填样式的变量名字）。

Lv1Special = True
'一级标题序号、标题序号之后的符号是否采用自定义格式。

Lv1Index = "一,二,三,四,五,六,七,八,九,十,十一,十二,十三,十四,十五,十六,十七,十八,十九,二十,二十一,二十二,二十三,二十四,二十五,二十六,二十七,二十八,二十九,三十,三十一,三十二,三十三,三十四,三十五,三十六,三十七,三十八,三十九,四十"
'自定义一级标题，用英文逗号 , 分隔

ChaAfterLv1 = "、"
'自定义一级标题序号后的符号

ParaPreDel_body = " 　"
'删除正文段落（无大纲级别的段落）前的指定字符，忽略表格中的文字。

ParaPreDel_title = " 　.、0123456789一二三四五六七八九十"
'删除标题段落（有大纲级别的段落）前的指定字符。

'1~9级大纲：
lv1 = "一级"
lv2 = "二级"
lv3 = "三级"
lv4 = "四级"
lv5 = "五级"
lv6to9 = "六级及以下"
'各样式的名字：
'注意：样式类型统一采用【链接段落和字符样式】，否则将影响ApplyStyle_SelAllPra的部分功能。

'图表共用标题
Title_Of_Pic_or_Table = "图表标题"

'正文
Main_body = "正文"

'正文居中
Main_body_mid = "居中正文"

'嵌入型图片(InlineShapes)，且本段只有图片，没有其他文字
Picture = "单张图片"

'表头
TableHead = "表头"

'表内文字
TextInTable = "表内文字"

'表题
TableTitle = "表题"

'图题
PicTitle = "图题"

'其他样式。将不希望被自动分配段落的样式名字放入其中，如封面、目录、摘要、参考文献等。程序将跳过具有这些样式的段落。
ExsistTypeName = "其它 封面文字 目录 摘要 参考文献"

End Sub
'========以下为程序内容========
Sub Main()
    Call ArgSetting
    Dim aaa() As Integer
    ReDim aaa(1)
    aaa(1) = 0
    
    Dim StartPara, p As Long
    p = ActiveDocument.Paragraphs.Count
    SelStart = Selection.Start '获取当前选择区域的起点
    Selend = Selection.End '获取当前选择区域的终点
    If Selend - SelStart > 1 Then '如果选中了一段话，就从头开始；如果没有，则从光标处开始
        StartPara = 1
    Else
        StartPara = Range(0, SelStart).Paragraphs.Count
    End If
    
    '主循环
    Dim n, m As Integer
    Dim j As Long
    For j = StartPara To p
        ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.Start).Select
        n = Selection.Paragraphs.OutlineLevel
        Dim CurrStyle As String
        CurrStyle = Selection.Style
        
        If n = 10 Then
n_eq_10:
            If InStr(ExsistTypeName, CurrStyle) = 0 Then
                If (Selection.Information(wdWithInTable) = False) Then  '不在表格中
                    If Len(ParaPreDel_body) > 0 Then '删除段落前的指定字符
                        Do While InStr(ParaPreDel_body, Selection) > 0
                            Selection.Delete
                        Loop
                    End If
                    Dim PhotoCount As Integer
                    PhotoCount = ActiveDocument.Paragraphs(j).Range.InlineShapes.Count
                    If PhotoCount > 0 Then
                        If (ActiveDocument.Paragraphs(j).Range.End - ActiveDocument.Paragraphs(j).Range.Start > PhotoCount + 2) Then
                            GoTo ZhengWen
                        Else
                            Call GiveStyle(j, Picture)
                        End If
                    Else
ZhengWen:
                        If (Selection.ParagraphFormat.Alignment = 1) Then
                            Call GiveStyle(j, Mid_SetTo)
                        Else
                            Call GiveStyle(j, Main_body)
                        End If
                    End If
                End If
            End If
        Else
            If Title_Len Then
                Dim TLen As Integer
                TLen = ActiveDocument.Paragraphs(j).Range.End - ActiveDocument.Paragraphs(j).Range.Start
                If ((Title_Len_Min > TLen) Or (TLen > Title_Len_Max)) Then
                    Selection.ParagraphFormat.OutlineLevel = 10
                    GoTo n_eq_10
                End If
            End If
            m = UBound(aaa)
            If n - m > 1 Then
                MsgBox "大纲级别设置错误，出现跳级"
				exit sub
            Else
                ReDim Preserve aaa(n)
                If n - m = 1 Then
                    aaa(n) = 1
                Else
                    aaa(n) = aaa(n) + 1
                End If
                Call ttfix(aaa, n, j)
            End If
        End If
        DoEvents
    Next j
    Call tablefix
    MsgBox "完成"
End Sub

'将段落j分配给样式StyleName
Sub GiveStyle(j, StyleName)
    If (ApplyStyle_SelAllPra) Then
        ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.End).Select
    Else
        ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.Start).Select
    End If
    Selection.Style = ActiveDocument.Styles(StyleName)
    ActiveDocument.Range(ActiveDocument.Paragraphs(j).Range.Start, ActiveDocument.Paragraphs(j).Range.Start).Select
End Sub

'TitleFix，修改标题
Sub ttfix(aaa, n, j)
    If Len(ParaPreDel_title) > 0 Then '删除段落前的指定字符
        Do While InStr(ParaPreDel_title, Selection) > 0
            Selection.Delete
        Loop
    End If

    Select Case n
    Case 1
    Selection.Style = ActiveDocument.Styles(lv1)
    Case 2
    Selection.Style = ActiveDocument.Styles(lv2)
    Case 3
    Selection.Style = ActiveDocument.Styles(lv3)
    Case 4
    Selection.Style = ActiveDocument.Styles(lv4)
    Case 5
    Selection.Style = ActiveDocument.Styles(lv5)
    Case Else
    Selection.Style = ActiveDocument.Styles(lv6to9)
    Selection.ParagraphFormat.OutlineLevel = n
    End Select

	if TypeIndex then
		If Lv1Special And n = 1 Then
			Dim xuhao() As String
			xuhao() = Split(Lv1Index, ",")
			Dim xh As String
			xh = xuhao(aaa(1) - 1)
			Selection.TypeText xh
			Selection.TypeText ChaAfterLv1
		Else
			Dim i As Integer
			i = 1
			Do While i < n
				Selection.TypeText aaa(i)
				Selection.TypeText "."
				i = i + 1
			Loop
			Selection.TypeText aaa(n)
			Selection.TypeText ChaAfterIndex
		End If
	end if
End Sub

'TableFix，修改表格内容，目前只写了分配样式，待完善。
Sub tablefix()
Dim t As Long
t = 1
Do While t <= ActiveDocument.Tables.Count
ActiveDocument.Tables(t).Select
Selection.Style = ActiveDocument.Styles(TextInTable)
t = t + 1
DoEvents
Loop
End Sub
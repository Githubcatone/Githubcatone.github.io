---
title: OfficeVBA
date: 2023-04-20
---

### 一键将word文档中的公式字体改为Times New Roman且非斜体

```VB
Sub ChangeEquationFormat()
    Dim eq As OMath '声明公式对象
    For Each eq In ActiveDocument.OMaths '遍历文档中所有公式
        eq.Range.OMaths(1).ConvertToNormalText '将公式输入格式改为Text
        eq.Range.Font.Name = "Times New Roman" '将字体更改为Times New Roman
        eq.Range.Font.Italic = False '不倾斜
    Next eq
End Sub
```

### 一键将文档中所有出现的“XXX”内容进行格式修改

```VB
Sub AddUnderlineAndBold()
    Dim findRange As Range
    Dim searchTerm As String
    
    searchTerm = "XXX" 'XXX为查找的内容
    
    Set findRange = ActiveDocument.Content
    
    With findRange.Find
        .Text = searchTerm
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            findRange.Font.Underline = True '加粗
            findRange.Font.Bold = True '下划线
            findRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub
```

### 一键将word文档中的15磅斜体文字改为添加下划线并加粗

与“选择格式相似的文本”功能相同，但是MS Office for Mac 并没有此功能。

```VB
Sub AddUnderlineAndBoldTo15ptItalicText()
    Dim range As Range
    
    Set range = ActiveDocument.Content
    
    With range.Find
        .ClearFormatting
        .Font.Size = 15
        .Font.Italic = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            range.Font.Underline = True
            range.Font.Bold = True
            range.Collapse wdCollapseEnd
        Loop
    End With
End Sub

```

全部参数：

| Font对象的属性              | 作用     |
| --------------------------- | -------- |
| Font.Name = "字体名称"      | 字体格式 |
| Font.Size = 12              | 字体大小 |
| Font.Bold = True            | 加粗     |
| Font.Italic = True          | 斜体     |
| Font.Underline = True       | 下划线   |
| Font.Color = RGB(255, 0, 0) | 字体颜色 |
| Font.Superscript = True     | 上标     |
| Font.Subscript = True       | 下标     |
| Font.StrikeThrough = True   | 删除线   |

### Office Tips

1. 在word中，
    - 打出三个“-”再按回车，生成一条长直线
    - 打出三个“=”再按回车，生成一条长双直线
    - 打出三个“*”再按回车，生成一条长虚线
    - 打出三个“#”再按回车，生成一条长隔行线
    - 打出三个“~”再按回车，生成一条长波浪线

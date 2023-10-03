import{_ as l}from"./plugin-vue_export-helper-c27b6911.js";import{r as d,o as s,c as a,a as n,b as e,e as r,d as t}from"./app-5a751137.js";const u={},o=t(`<h3 id="一键将word文档中的公式字体改为times-new-roman且非斜体" tabindex="-1"><a class="header-anchor" href="#一键将word文档中的公式字体改为times-new-roman且非斜体" aria-hidden="true">#</a> 一键将word文档中的公式字体改为Times New Roman且非斜体</h3><div class="language-VB line-numbers-mode" data-ext="VB"><pre class="language-VB"><code>Sub ChangeEquationFormat()
    Dim eq As OMath &#39;声明公式对象
    For Each eq In ActiveDocument.OMaths &#39;遍历文档中所有公式
        eq.Range.OMaths(1).ConvertToNormalText &#39;将公式输入格式改为Text
        eq.Range.Font.Name = &quot;Times New Roman&quot; &#39;将字体更改为Times New Roman
        eq.Range.Font.Italic = False &#39;不倾斜
    Next eq
End Sub
</code></pre><div class="line-numbers" aria-hidden="true"><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div></div></div><h3 id="一键将文档中所有出现的-xxx-内容进行格式修改" tabindex="-1"><a class="header-anchor" href="#一键将文档中所有出现的-xxx-内容进行格式修改" aria-hidden="true">#</a> 一键将文档中所有出现的“XXX”内容进行格式修改</h3><div class="language-VB line-numbers-mode" data-ext="VB"><pre class="language-VB"><code>Sub AddUnderlineAndBold()
    Dim findRange As Range
    Dim searchTerm As String
    
    searchTerm = &quot;XXX&quot; &#39;XXX为查找的内容
    
    Set findRange = ActiveDocument.Content
    
    With findRange.Find
        .Text = searchTerm
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            findRange.Font.Underline = True &#39;加粗
            findRange.Font.Bold = True &#39;下划线
            findRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub
</code></pre><div class="line-numbers" aria-hidden="true"><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div></div></div><h3 id="一键将word文档中的15磅斜体文字改为添加下划线并加粗" tabindex="-1"><a class="header-anchor" href="#一键将word文档中的15磅斜体文字改为添加下划线并加粗" aria-hidden="true">#</a> 一键将word文档中的15磅斜体文字改为添加下划线并加粗</h3><p>与“选择格式相似的文本”功能相同，但是MS Office for Mac 并没有此功能。</p><div class="language-VB line-numbers-mode" data-ext="VB"><pre class="language-VB"><code>Sub AddUnderlineAndBoldTo15ptItalicText()
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

</code></pre><div class="line-numbers" aria-hidden="true"><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div><div class="line-number"></div></div></div><p>全部参数：</p>`,8),c=n("thead",null,[n("tr",null,[n("th",null,"Font对象的属性"),n("th",null,"作用")])],-1),v={href:"http://Font.Name",target:"_blank",rel:"noopener noreferrer"},m=n("td",null,"字体格式",-1),b=n("tr",null,[n("td",null,"Font.Size = 12"),n("td",null,"字体大小")],-1),h=n("tr",null,[n("td",null,"Font.Bold = True"),n("td",null,"加粗")],-1),_=n("tr",null,[n("td",null,"Font.Italic = True"),n("td",null,"斜体")],-1),p=n("tr",null,[n("td",null,"Font.Underline = True"),n("td",null,"下划线")],-1),g=n("tr",null,[n("td",null,"Font.Color = RGB(255, 0, 0)"),n("td",null,"字体颜色")],-1),f=n("tr",null,[n("td",null,"Font.Superscript = True"),n("td",null,"上标")],-1),F=n("tr",null,[n("td",null,"Font.Subscript = True"),n("td",null,"下标")],-1),T=n("tr",null,[n("td",null,"Font.StrikeThrough = True"),n("td",null,"删除线")],-1),x=n("h3",{id:"office-tips",tabindex:"-1"},[n("a",{class:"header-anchor",href:"#office-tips","aria-hidden":"true"},"#"),e(" Office Tips")],-1),B=n("ol",null,[n("li",null,[e("在word中， "),n("ul",null,[n("li",null,"打出三个“-”再按回车，生成一条长直线"),n("li",null,"打出三个“=”再按回车，生成一条长双直线"),n("li",null,"打出三个“*”再按回车，生成一条长虚线"),n("li",null,"打出三个“#”再按回车，生成一条长隔行线"),n("li",null,"打出三个“~”再按回车，生成一条长波浪线")])])],-1);function w(S,R){const i=d("ExternalLinkIcon");return s(),a("div",null,[o,n("table",null,[c,n("tbody",null,[n("tr",null,[n("td",null,[n("a",v,[e("Font.Name"),r(i)]),e(' = "字体名称"')]),m]),b,h,_,p,g,f,F,T])]),x,B])}const A=l(u,[["render",w],["__file","2023-04-20-OfficeVBA.html.vue"]]);export{A as default};

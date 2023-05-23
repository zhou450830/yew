' 定义变量
Dim objFSO, objFolder, objFile, objWord, objDoc
Dim strFolderPath, strMacroFile, strStyleFile
Dim arrStyles(), arrFormats()

' 设置文件夹路径
strFolderPath = InputBox("请输入文件夹路径：")

' 获取文件夹对象
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolderPath)

' 获取word对象
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

' 循环处理文件夹内的doc、docx文件
For Each objFile In objFolder.Files
    If Right(objFile.Name, 4) = "docx" Or Right(objFile.Name, 3) = "doc" Then
        ' 打开文档
        Set objDoc = objWord.Documents.Open(objFile.Path)
        
        ' 调用宏代码
        strMacroFile = "WordVBASample.docm"
        objWord.Run MacroName:=strMacroFile & "!Main"

        ' 关闭文档
        objDoc.Close SaveChanges:=True
    End If
Next

' 退出word
objWord.Quit

' 宏代码部分
Sub Main()
    ' 定义变量
    Dim i, j, k, m, n, intLevel, intPrevLevel, intSectionNum
    Dim objRange, objPara, objStyle, objTOC, objTable
    Dim strStyleName, strReg, strReplace, strTOCStyle, strCellStyle, strLangName
    Dim arrTOCStyles(), arrCellFormats()
    
    ' 读入样式文件，定义文档样式
    strStyleFile = "WordVBASample.docm"
    Set objDoc = ActiveDocument
    Set objRange = objDoc.Range(0, 0)
    objRange.InsertFile strStyleFile
    Set objStyle = objDoc.Styles("Normal")
    objStyle.Font.Name = "宋体"
    objStyle.Font.Size = 14
    objStyle.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    objDoc.DefaultTabStop = CentimetersToPoints(1.25)

    ' 定义标题样式
    ReDim arrStyles(6)
    arrStyles(1) = Array("一级标题", "黑体", 22, wdAlignParagraphCenter)
    arrStyles(2) = Array("二级标题", "黑体", 16, wdAlignParagraphCenter)
    arrStyles(3) = Array("三级标题", "黑体", 14, wdAlignParagraphCenter)
    arrStyles(4) = Array("四级标题", "黑体", 12, wdAlignParagraphLeft)
    arrStyles(5) = Array("五级标题", "宋体", 12, wdAlignParagraphLeft)
    arrStyles(6) = Array("六级标题", "宋体", 12, wdAlignParagraphLeft)
    For i = 1 To 6
        strStyleName = arrStyles(i)(0)
        objDoc.Styles.Add strStyleName, wdStyleTypeParagraph
        Set objStyle = objDoc.Styles(strStyleName)
        objStyle.Font.Name = arrStyles(i)(1)
        objStyle.Font.Size = arrStyles(i)(2)
        objStyle.Font.Bold = True
        objStyle.ParagraphFormat.Alignment = arrStyles(i)(3)
    Next

    ' 定义正文样式
    objDoc.Styles.Add "正文", wdStyleTypeParagraph
    Set objStyle = objDoc.Styles("正文")
    objStyle.Font.Name = "宋体"
    objStyle.Font.Size = 12
    objStyle.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    objStyle.ParagraphFormat.Alignment = wdAlignParagraphJustify

    ' 定义列表样式
    objDoc.ListTemplates.Add
    objDoc.ListTemplates(1).ListLevels(1).NumberFormat = "%1."
    objDoc.ListTemplates(1).ListLevels(1).TrailingCharacter = wdTrailingTab
    objDoc.ListTemplates(1).ListLevels(1).NumberStyle = wdListNumberStyleArabic
    objDoc.ListTemplates(1).ListLevels(1).NumberPosition = CentimetersToPoints(0.63)
    objDoc.ListTemplates(1).ListLevels(1).Alignment = wdListLevelAlignLeft
    For i = 1 To 6
        objDoc.ListTemplates(1).ListLevels(i).LinkNumberedToPrevious = False
        objDoc.ListTemplates(1).ListLevels(i).Font.Name = arrStyles(i)(1)
        objDoc.ListTemplates(1).ListLevels(i).Font.Size = arrStyles(i)(2)
    Next

    ' 定义图片样式
    objDoc.Styles.Add "图片", wdStyleTypeParagraph
    Set objStyle = objDoc.Styles("图片")
    objStyle.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    objStyle.ParagraphFormat.Alignment = wdAlignParagraphCenter
    objStyle.ParagraphFormat.SpaceBefore = CentimetersToPoints(0.5)
    objStyle.ParagraphFormat.SpaceAfter = CentimetersToPoints(0.5)
    objStyle.ParagraphFormat.WidowControl = False
    objStyle.ParagraphFormat.KeepWithNext = True

    ' 定义表格样式
    objDoc.Styles.Add "表格", wdStyleTypeTable
    Set objStyle = objDoc.Styles("表格")
    objStyle.Table.AllowAutoFit = False
    objStyle.Table.LeftIndent = 0
    objStyle.Table.RightIndent = 0
    objStyle.Table.TopPadding = 0
    objStyle.Table.BottomPadding = 0
    objStyle.Table.LeftPadding = 0
    objStyle.Table.RightPadding = 0
    objStyle.Table.Spacing = 0
    objStyle.Table.Style = "网格型-Accent5"
    objStyle.Font.Name = "宋体"
    objStyle.Font.Size = 12

    ' 定义索引样式
    objDoc.Styles.Add "索引", wdStyleTypeParagraph
    Set objStyle = objDoc.Styles("索引")
    objStyle.ParagraphFormat.Alignment = wdAlignParagraphLeft
    objStyle.ParagraphFormat.SpaceBefore = CentimetersToPoints(0.5)
    objStyle.ParagraphFormat.SpaceAfter = CentimetersToPoints(0.25)

    ' 设置语言
    strLangName = "中文（中国）"
    objDoc.Content.LanguageID = wdChinesePRC
    
    ' 定义自动编号
    Set objRange = objDoc.Range(0, 0)
    objRange.Select
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListItem
    Set objTOC = objDoc.TablesOfContents.Add(objDoc.Range(0, 0), True)
    
    ' 设置目录样式
    strTOCStyle = "目录 1"
    Set objStyle = objDoc.Styles(strTOCStyle)
    objStyle.Font.Name = "黑体"
    objStyle.Font.Size = 14
    objStyle.Font.Bold = True
    objStyle.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    objStyle.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    ' 自动生成目录
    With objTOC
        .TabLeader = wdTabLeaderSpaces
        .UpperHeadingLevel = 1
        .LowerHeadingLevel = 6
        .IncludePageNumbers = True
        .RightAlignPageNumbers = True
        .UseFields = False
        .TableID = 1
        .OutlineLevels(1).Range.Style = strTOCStyle
        .OutlineLevels(2).Range.Style = arrStyles(2)(0)
        .OutlineLevels(3).Range.Style = arrStyles(3)(0)
        .OutlineLevels(4).Range.Style = arrStyles(4)(0)
        .OutlineLevels(5).Range.Style = arrStyles(5)(0)
        .OutlineLevels(6).Range.Style = arrStyles(6)(0)
    End With
End Sub

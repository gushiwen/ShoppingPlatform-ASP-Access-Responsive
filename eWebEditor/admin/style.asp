<!--#include file = "private.asp"-->
<%
'######################################
' eWebEditor V10.0 - Advanced online web based WYSIWYG HTML editor.
' Copyright (c) 2003-2015 eWebSoft.com
'
' For further information go to http://www.ewebeditor.net/
' This copyright notice MUST stay intact for use.
'######################################
%>
<%
If CheckLicense()=False Then
Call GoLicense()
End If
Dim sStyleID, sStyleName, sFixWidth, sSkin, sStyleUploadDir, sStyleWidth, sStyleHeight, sStyleMemo, sStyleStateFlag, sStyleDetectFromWord, sStyleInitMode, sStyleBaseUrl, sStyleUploadObject, sStyleAutoDir, sStyleBaseHref, sStyleContentPath, sStyleAutoRemote, sStyleShowBorder, sStyleAllowBrowse
Dim sAutoDetectLanguage, sDefaultLanguage, sSLTFlag, sSLTMode, sSLTCheckFlag, sSLTMinSize, sSLTOkSize, sSYWZFlag, sSYText, sSYFontColor, sSYFontSize, sSYFontName, sSYPicPath, sSLTSYObject, sSLTSYExt, sSYWZMinWidth, sSYShadowColor, sSYShadowOffset, sSYWZMinHeight, sSYWZPosition, sSYWZTextWidth, sSYWZTextHeight, sSYWZPaddingH, sSYWZPaddingV, sSYTPFlag, sSYTPMinWidth, sSYTPMinHeight, sSYTPPosition, sSYTPPaddingH, sSYTPPaddingV, sSYTPImageWidth, sSYTPImageHeight, sSYTPOpacity, sAdvApiFlag
Dim sStyleFileExt, sStyleFlashExt, sStyleImageExt, sStyleMediaExt, sStyleRemoteExt, sStyleLocalExt, sStyleFileSize, sStyleFlashSize, sStyleImageSize, sStyleMediaSize, sStyleRemoteSize, sStyleLocalSize
Dim sToolBarID, sToolBarName, sToolBarOrder, sToolBarButton
Dim sSBCode, sSBEdit, sSBText, sSBView, sSBSize
Dim sEnterMode, sAreaCssMode, sFileNameMode, sEncryptKey
Dim sPaginationMode, sPaginationKey, sPaginationAutoFlag, sPaginationAutoNum, sSpaceSize
Dim sMFUMode, sMFUBlockSize, sMFUEnable
Dim sCodeFormat, sTB2Flag, sTB2Code, sTB2Edit, sTB2Text, sTB2View, sTB2Max, sShowBlock
Dim sFileNameSameFix, sAutoDirOrderFlag, sAutoTypeDirImage, sAutoTypeDirFlash, sAutoTypeDirMedia, sAutoTypeDirAttach, sAutoTypeDirRemote, sAutoTypeDirLocal
Dim sWordImportInitMode, sWordImportAPI, sQuickFormatInitFontName, sQuickFormatInitFontSize
Dim sUIMinHeight, sSYValidNormal, sSYValidLocal, sSYValidRemote
Dim sFSOname
Dim sAutoDoneWordPaste, sAutoDoneExcelPaste, sAutoDoneQuickFormat, sAutoDonePasteOption
Dim nStyleID
sPosition = sPosition & "样式管理"
If sAction = "STYLEPREVIEW" Then
Call InitStyle()
Call ShowStylePreview()
Response.End
End If
Call Header()
Call ShowPosition()
Call Content()
Call Footer()
Sub Content()
Select Case sAction
Case "COPY"
Call InitStyle()
Call DoCopy()
Call ShowStyleList()
Case "STYLEADD"
Call ShowStyleForm("ADD")
Case "STYLESET"
Call InitStyle()
Call ShowStyleForm("SET")
Case "STYLEADDSAVE"
Call CheckStyleForm()
Call DoStyleAddSave()
Case "STYLESETSAVE"
Call CheckStyleForm()
Call DoStyleSetSave()
Case "STYLEDEL"
Call InitStyle()
Call DoStyleDel()
Call ShowStyleList()
Case "CODE"
Call InitStyle()
Call ShowStyleCode()
Case "TOOLBAR"
Call InitStyle()
Call ShowToolBarList()
Case "TOOLBARADD"
Call InitStyle()
Call DoToolBarAdd()
Call ShowToolBarList()
Case "TOOLBARMODI"
Call InitStyle()
Call DoToolBarModi()
Call ShowToolBarList()
Case "TOOLBARDEL"
Call InitStyle()
Call DoToolBarDel()
Call ShowToolBarList()
Case "BUTTONSET"
Call InitStyle()
Call InitToolBar()
Call ShowButtonList()
Case "BUTTONSAVE"
Call InitStyle()
Call InitToolBar()
Call DoButtonSave()
Case Else
Call ShowStyleList()
End Select
End Sub
Sub ShowPosition()
Response.Write "<table border=0 cellspacing=1 align=center class=navi>" & _
"<tr><th>" & sPosition & "</th></tr>" & _
"<tr><td align=center>[<a href='?'>所有样式列表</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='?action=styleadd'>新建一样式</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='#' onclick='history.back()'>返回前一页</a>]</td></tr>" & _
"</table><br>"
End Sub
Sub ShowMessage(str)
Response.Write "<table border=0 cellspacing=1 align=center class=list><tr><td>" & str & "</td></tr></table><br>"
End Sub
Sub ShowStyleList()
Call ShowMessage("<b class=blue>以下为当前所有样式列表：</b>")
Response.Write "<table border=0 cellpadding=0 cellspacing=1 class=list align=center>" & _
"<form action='?action=del' method=post name=myform>" & _
"<tr align=center>" & _
"<th width='10%'>样式名</th>" & _
"<th width='10%'>最佳宽度</th>" & _
"<th width='10%'>最佳高度</th>" & _
"<th width='45%'>说明</th>" & _
"<th width='25%'>管理</th>" & _
"</tr>"
Dim sManage, i, aCurrStyle
For i = 1 To Ubound(aStyle)
aCurrStyle = Split(aStyle(i), "|||")
sManage = "<a href='?action=stylepreview&id=" & i & "' target='_blank'>预览</a>|<a href='?action=code&id=" & i & "'>代码</a>|<a href='?action=styleset&id=" & i & "'>设置</a>|<a href='?action=toolbar&id=" & i & "'>工具栏</a>|<a href='?action=copy&id=" & i & "'>拷贝</a>|<a href='?action=styledel&id=" & i & "' onclick=""return confirm('提示：您确定要删除此样式吗？')"">删除</a>"
Response.Write "<tr align=center>" & _
"<td>" & outHTML(aCurrStyle(0)) & "</td>" & _
"<td>" & aCurrStyle(4) & "</td>" & _
"<td>" & aCurrStyle(5) & "</td>" & _
"<td align=left>" & outHTML(aCurrStyle(26)) & "</td>" & _
"<td>" & sManage & "</td>" & _
"</tr>"
Next
Response.Write "</table><br>"
Call ShowMessage("<b class=blue>提示：</b>你可以通过“拷贝”一样式以达到快速新建样式的目的。")
End Sub
Sub DoCopy()
Dim i, b, sNewID, sNewName
b = False
i = 0
Do While b = False
i = i + 1
sNewName = sStyleName & i
If StyleName2ID(sNewName) = -1 Then
b = True
End If
Loop
Dim nNewStyleID
nNewStyleID = Ubound(aStyle) + 1
Redim Preserve aStyle(nNewStyleID)
aStyle(nNewStyleID) = sNewName & Mid(aStyle(nStyleID), Len(sStyleName)+1)
Dim nToolbarNum, nNewToolbarID, aCurrToolbar
nToolbarNum = Ubound(aToolbar)
For i = 1 To nToolbarNum
aCurrToolbar = Split(aToolbar(i), "|||")
If aCurrToolbar(0) = sStyleID Then
nNewToolbarID = Ubound(aToolbar) + 1
Redim Preserve aToolbar(nNewToolbarID)
aToolbar(nNewToolbarID) = nNewStyleID & "|||" & aCurrToolbar(1) & "|||" & aCurrToolbar(2) & "|||" & aCurrToolbar(3)
End If
Next
Call WriteConfig()
Call GoUrl("?")
End Sub
Function StyleName2ID(str)
Dim i
StyleName2ID = -1
For i = 1 To UBound(aStyle)
If Lcase(Split(aStyle(i), "|||")(0)) = Lcase(str) Then
StyleName2ID = i
Exit Function
End If
Next
End Function
Sub ShowStyleForm(sFlag)
Dim s_Title, s_Action
Dim s_FormStateFlag, s_FormDetectFromWord, s_FormInitMode, s_FormBaseUrl, s_FormUploadObject, s_FormAutoRemote, s_FormShowBorder, s_FormSLTFlag, s_FormSLTMode, s_FormSLTCheckFlag, s_FormSYWZFlag, s_FormSLTSYObject, s_FormAllowBrowse, s_FormSYTPFlag, s_FormSYWZPosition, s_FormSYTPPosition, s_FormAdvApiFlag, s_FormMFUMode, s_FormMFUEnable, s_FormCodeFormat, s_FormShowBlock
Dim s_FormSBCode, s_FormSBEdit, s_FormSBView, s_FormSBText, s_FormSBSize
Dim s_FormAutoDetectLanguage, s_FormDefaultLanguage, s_FormEnterMode, s_FormAreaCssMode, s_FormFileNameMode
Dim s_FormPaginationMode, s_FormPaginationAutoFlag
Dim s_FormTB2Flag, s_FormTB2Code, s_FormTB2Edit, s_FormTB2Text, s_FormTB2View, s_FormTB2Max
Dim s_FormAutoDirOrderFlag, s_FormWordImportInitMode, s_FormWordImportAPI
Dim s_FormSYValidNormal, s_FormSYValidLocal, s_FormSYValidRemote
Dim s_FormAutoDoneWordPaste, s_FormAutoDoneExcelPaste, s_FormAutoDoneQuickFormat, s_FormAutoDonePasteOption
If sFlag = "ADD" Then
sStyleID = ""
sStyleName = ""
sFixWidth = ""
sSkin = "office2003"
sStyleUploadDir = "uploadfile/"
sStyleBaseHref = ""
sStyleContentPath = ""
sStyleWidth = "550"
sStyleHeight = "350"
sStyleMemo = ""
s_Title = "新增样式"
s_Action = "StyleAddSave"
sStyleFileExt = "rar|zip|exe|doc|xls|chm|hlp"
sStyleFlashExt = "swf"
sStyleImageExt = "gif|jpg|jpeg|bmp"
sStyleMediaExt = "rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov"
sStyleRemoteExt = "gif|jpg|bmp"
sStyleFileSize = "500"
sStyleFlashSize = "100"
sStyleImageSize = "100"
sStyleMediaSize = "100"
sStyleRemoteSize = "100"
sStyleStateFlag = "1"
sSBCode = "1"
sSBEdit = "1"
sSBText = "1"
sSBView = "1"
sSBSize = "1"
sAutoDetectLanguage = "1"
sDefaultLanguage = "zh-cn"
sEnterMode = "1"
sAreaCssMode = "0"
sFileNameMode = "0"
sStyleAutoRemote = "1"
sStyleShowBorder = "0"
sStyleAllowBrowse = "0"
sStyleUploadObject = "0"
sFSOname = ""
sStyleAutoDir = ""
sStyleDetectFromWord = "1"
sStyleInitMode = "EDIT"
sStyleBaseUrl = "1"
sSLTFlag = "0"
sSLTMode = "0"
sSLTCheckFlag = "0"
sSLTMinSize = "300"
sSLTOkSize = "120"
sSYWZFlag = "0"
sSYText = "版权所有..."
sSYFontColor = "000000"
sSYFontSize = "12"
sSYFontName = "宋体"
sSYPicPath = ""
sSLTSYObject = "0"
sSLTSYExt = "bmp|jpg|jpeg|gif"
sSYWZMinWidth = "100"
sSYShadowColor = "FFFFFF"
sSYShadowOffset = "1"
sStyleLocalExt = "gif|jpg|bmp|wmz"
sStyleLocalSize = "100"
sSYWZMinHeight = "100"
sSYWZPosition = "1"
sSYWZTextWidth = "66"
sSYWZTextHeight = "17"
sSYWZPaddingH = "5"
sSYWZPaddingV = "5"
sSYTPFlag = "0"
sSYTPMinWidth = "100"
sSYTPMinHeight = "100"
sSYTPPosition = "1"
sSYTPPaddingH = "5"
sSYTPPaddingV = "5"
sSYTPImageWidth = "88"
sSYTPImageHeight = "31"
sSYTPOpacity = "1"
sAdvApiFlag = "0"
sEncryptKey = ""
sPaginationMode = "1"
sPaginationKey = "{page}"
sPaginationAutoFlag = "0"
sPaginationAutoNum = "2000"
sSpaceSize = ""
sMFUMode = "0"
sMFUBlockSize = "200"
sMFUEnable = "1"
sCodeFormat = "2"
sTB2Flag = "1"
sTB2Code = "1"
sTB2Edit = "1"
sTB2Text = "1"
sTB2View = "1"
sTB2Max = "1"
sShowBlock = "0"
sFileNameSameFix = ""
sAutoDirOrderFlag = "0"
sAutoTypeDirImage = ""
sAutoTypeDirFlash = ""
sAutoTypeDirMedia = ""
sAutoTypeDirAttach = ""
sAutoTypeDirRemote = ""
sAutoTypeDirLocal = ""
sWordImportInitMode = "1"
sWordImportAPI = "1"
sQuickFormatInitFontName = ""
sQuickFormatInitFontSize = ""
sUIMinHeight = "300"
sSYValidNormal = "1"
sSYValidLocal = ""
sSYValidRemote = ""
sAutoDoneWordPaste = ""
sAutoDoneExcelPaste = ""
sAutoDoneQuickFormat = ""
sAutoDonePasteOption = ""
Else
sStyleName = inHTML(sStyleName)
sSkin = inHTML(sSkin)
sStyleUploadDir = inHTML(sStyleUploadDir)
sStyleBaseHref = inHTML(sStyleBaseHref)
sStyleContentPath = inHTML(sStyleContentPath)
sStyleMemo = inHTML(sStyleMemo)
sSYText = inHTML(sSYText)
sSYFontColor = inHTML(sSYFontColor)
sSYFontSize = inHTML(sSYFontSize)
sSYFontName = inHTML(sSYFontName)
sSYPicPath = inHTML(sSYPicPath)
s_Title = "设置样式"
s_Action = "StyleSetSave"
End If
s_FormStateFlag = InitCheckBox("d_stateflag", "1", sStyleStateFlag)
s_FormSBCode = InitCheckBox("d_sbcode", "1", sSBCode)
s_FormSBEdit = InitCheckBox("d_sbedit", "1", sSBEdit)
s_FormSBText = InitCheckBox("d_sbtext", "1", sSBText)
s_FormSBView = InitCheckBox("d_sbview", "1", sSBView)
s_FormSBSize = InitCheckBox("d_sbsize", "1", sSBSize)
s_FormTB2Flag = InitCheckBox("d_tb2flag", "1", sTB2Flag)
s_FormTB2Code = InitCheckBox("d_tb2code", "1", sTB2Code)
s_FormTB2Edit = InitCheckBox("d_tb2edit", "1", sTB2Edit)
s_FormTB2Text = InitCheckBox("d_tb2text", "1", sTB2Text)
s_FormTB2View = InitCheckBox("d_tb2view", "1", sTB2View)
s_FormTB2Max = InitCheckBox("d_tb2max", "1", sTB2Max)
s_FormSYValidNormal = InitCheckBox("d_syvalidnormal", "1", sSYValidNormal)
s_FormSYValidLocal = InitCheckBox("d_syvalidlocal", "1", sSYValidLocal)
s_FormSYValidRemote = InitCheckBox("d_syvalidremote", "1", sSYValidRemote)
s_FormAutoDoneWordPaste = InitCheckBox("d_autodonewordpaste", "1", sAutoDoneWordPaste)
s_FormAutoDoneExcelPaste = InitCheckBox("d_autodoneexcelpaste", "1", sAutoDoneExcelPaste)
s_FormAutoDoneQuickFormat = InitCheckBox("d_autodonequickformat", "1", sAutoDoneQuickFormat)
s_FormAutoDonePasteOption = InitCheckBox("d_autodonepasteoption", "1", sAutoDonePasteOption)
s_FormAutoDetectLanguage = InitSelect("d_autodetectlanguage", Split("不自动检测|自动检测", "|"), Split("0|1", "|"), sAutoDetectLanguage, "", "")
s_FormDefaultLanguage = InitSelect("d_defaultlanguage", Split("简体中文|繁体中文|英文|日语|西班牙语|俄语|德语|法语|意大利语|荷兰语|瑞典语|葡萄牙语|挪威语|丹麦语", "|"), Split("zh-cn|zh-tw|en|ja|es|ru|de|fr|it|nl|sv|pt|no|da", "|"), sDefaultLanguage, "", "")
s_FormEnterMode = InitSelect("d_entermode", Split("Enter输入<P>，Shift+Enter输入<BR>|Enter输入<BR>，Shift+Enter输入<P>", "|"), Split("1|2", "|"), sEnterMode, "", "")
s_FormAreaCssMode = InitSelect("d_areacssmode", Split("常规模式|Word导入模式", "|"), Split("0|1", "|"), sAreaCssMode, "", "")
s_FormFileNameMode = InitSelect("d_filenamemode", Split("所有：自动重命名|所有：原文件名|附件按原名，其它自动重命名", "|"), Split("0|1|2", "|"), sFileNameMode, "", "")
s_FormAutoRemote = InitSelect("d_autoremote", Split("自动上传|不自动上传", "|"), Split("1|0", "|"), sStyleAutoRemote, "", "")
s_FormShowBorder = InitSelect("d_showborder", Split("默认显示|默认不显示", "|"), Split("1|0", "|"), sStyleShowBorder, "", "")
s_FormShowBlock = InitSelect("d_showblock", Split("默认显示|默认不显示", "|"), Split("1|0", "|"), sShowBlock, "", "")
s_FormAllowBrowse = InitSelect("d_allowbrowse", Split("是,开启|否,关闭", "|"), Split("1|0", "|"), sStyleAllowBrowse, "", "")
s_FormMFUMode = InitSelect("d_mfumode", Split("ADO文件流(不支持大文件)|eWebEditorMFUServer COM组件(支持大文件)", "|"), Split("0|1", "|"), sMFUMode, "", "")
s_FormMFUEnable = InitSelect("d_mfuenable", Split("是,开启|否,关闭", "|"), Split("1|0", "|"), sMFUEnable, "", "")
s_FormCodeFormat = InitSelect("d_codeformat", Split("关闭|启用:缩进1空格|启用:缩进2空格|启用:缩进3空格|启用:缩进4空格|启用:缩进5空格|启用:缩进6空格|启用:缩进7空格|启用:缩进8空格", "|"), Split("0|1|2|3|4|5|6|7|8", "|"), sCodeFormat, "", "")
s_FormUploadObject = InitSelect("d_uploadobject", Split("无组件上传类|ASPUpload上传组件|SA-FileUp上传组件|LyfUpload上传组件", "|"), Split("0|1|2|3", "|"), sStyleUploadObject, "", "")
s_FormDetectFromWord = InitSelect("d_detectfromword", Split("启用强制纯文本粘贴|启用高级粘贴|不启用", "|"), Split("2|1|0", "|"), sStyleDetectFromWord, "", "")
s_FormInitMode = InitSelect("d_initmode", Split("代码模式|编辑模式|文本模式|预览模式", "|"), Split("CODE|EDIT|TEXT|VIEW", "|"), sStyleInitMode, "", "")
s_FormBaseUrl = InitSelect("d_baseurl", Split("相对路径|绝对根路径|绝对全路径|站外绝对全路径", "|"), Split("0|1|2|3", "|"), sStyleBaseUrl, "", "")
s_FormSLTFlag = InitSelect("d_sltflag", Split("不使用|使用|模拟使用,不生成小图,改大图显示宽高", "|"), Split("0|1|2", "|"), sSLTFlag, "", "")
s_FormSLTMode = InitSelect("d_sltmode", Split("大小图:显示小图,链到大图|大小图:显示大图|只生成小图", "|"), Split("0|1|2", "|"), sSLTMode, "", "")
s_FormSLTCheckFlag = InitSelect("d_sltcheckflag", Split("宽|高|宽或高", "|"), Split("0|1|2", "|"), sSLTCheckFlag, "", "")
s_FormSYWZFlag = InitSelect("d_sywzflag", Split("不使用|使用|前台用户控制", "|"), Split("0|1|2", "|"), sSYWZFlag, "", "")
s_FormSLTSYObject = InitSelect("d_sltsyobject", Split("AspJpeg V1.3+|AspJpeg V1.8+", "|"), Split("0|1", "|"), sSLTSYObject, "", "")
s_FormSYTPFlag = InitSelect("d_sytpflag", Split("不使用|使用|前台用户控制", "|"), Split("0|1|2", "|"), sSYTPFlag, "", "")
s_FormSYWZPosition = InitSelect("d_sywzposition", Split("左上|左中|左下|中上|中中|中下|右上|右中|右下", "|"), Split("1|2|3|4|5|6|7|8|9", "|"), sSYWZPosition, "", "")
s_FormSYTPPosition = InitSelect("d_sytpposition", Split("左上|左中|左下|中上|中中|中下|右上|右中|右下", "|"), Split("1|2|3|4|5|6|7|8|9", "|"), sSYTPPosition, "", "")
s_FormAdvApiFlag = InitSelect("d_advapiflag", Split("禁用|启用一般接口(明文cusdir)|启用高级接口(Session安全)", "|"), Split("0|1|2", "|"), sAdvApiFlag, "", "")
s_FormPaginationMode = InitSelect("d_paginationmode", Split("不启用|启用：标准分页符|启用：自定义分页符", "|"), Split("0|1|2", "|"), sPaginationMode, "", "")
s_FormPaginationAutoFlag = InitSelect("d_paginationautoflag", Split("不启用|部分启用,内容中已有分页时不启用|完全启用,内容中已有的分页会被替换", "|"), Split("0|1|2", "|"), sPaginationAutoFlag, "", "")
s_FormAutoDirOrderFlag = InitSelect("d_autodirorderflag", Split("文件类型目录/年月日目录/|年月日目录/文件类型目录/", "|"), Split("0|1", "|"), sAutoDirOrderFlag, "", "")
s_FormWordImportInitMode = InitSelect("d_wordimportinitmode", Split("选择优化模式|全部清除模式", "|"), Split("1|2", "|"), sWordImportInitMode, "", "")
s_FormWordImportAPI = InitSelect("d_wordimportapi", Split("界面有选项，初始为[自动处理]|界面有选项，初始为[微软Office]|界面有选项，初始为[金山WPS]|界面无选项，固定为[自动处理]|界面无选项，固定为[微软Office]|界面无选项，固定为[金山WPS]", "|"), Split("0|1|2|10|11|12", "|"), sWordImportAPI, "", "")
Response.Write "<table border=0 cellpadding=0 cellspacing=1 align=center class=form>" & _
"<form action='?action=" & s_Action & "&id=" & sStyleID & "' method=post name=myform onsubmit='return checkStyleSetForm(this)'>" & _
"<tr><th colspan=4>&nbsp;&nbsp;" & s_Title & "（鼠标移到输入框可看说明，带*号为必填项）</th></tr>" & _
"<tr><td width='15%'>样式名称：</td><td width='35%'><input type=text class=input size=20 name=d_name title='引用此样式的名字，不要加特殊符号' value=""" & sStyleName & """> <span class=red>*</span></td><td width='15%'>初始模式：</td><td width='35%'>" & s_FormInitMode & " <span class=red>*</span></td></tr>" & _
"<tr><td>限宽模式宽度：</td><td><input type=text class=input size=20 name=d_fixwidth title='留空表示不启用，可以填入如：500px' value=""" & sFixWidth & """></td><td>界面皮肤目录：</td><td><input type=text class=input size=15 name=d_skin title='存放界面皮肤文件的目录名，必须在skin下' value=""" & sSkin & """> <select size=1 id=d_skin_drop onchange='this.form.d_skin.value=this.value'><option>-系统自带-</option><option value='blue1'>blue1</option><option value='blue2'>blue2</option><option value='flat1'>flat1</option><option value='flat2'>flat2</option><option value='flat3'>flat3</option><option value='flat4'>flat4</option><option value='flat5'>flat5</option><option value='flat7'>flat6</option><option value='flat8'>flat7</option><option value='flat8'>flat8</option><option value='flat9'>flat9</option><option value='flat10'>flat10</option><option value='green1'>green1</option><option value='light1'>light1</option><option value='office2000'>office2000</option><option value='office2003'>office2003</option><option value='officexp'>officexp</option><option value='red1'>red1</option><option value='vista1'>vista1</option><option value='yellow1'>yellow1</option></select> <span class=red>*</span></td></tr>" & _
"<tr><td>最佳宽度：</td><td><input type=text class=input name=d_width size=20 title='最佳引用效果的宽度，数字型' value='" & sStyleWidth & "'> <span class=red>*</span></td><td>最佳高度：</td><td><input type=text class=input name=d_height size=20 title='最佳引用效果的高度，数字型' value='" & sStyleHeight & "'> <span class=red>*</span></td></tr>" & _
"<tr><td>显示状态栏及按钮：</td><td>" & s_FormStateFlag & "状态栏 " & s_FormSBCode & "代码 " & s_FormSBEdit & "编辑 " & s_FormSBText & "文本 " & s_FormSBView & "预览 " & s_FormSBSize & "缩放<span class=red>*</span></td><td>高级粘贴自动检测：</td><td>" & s_FormDetectFromWord & " <span class=red>*</span></td></tr>" & _
"<tr><td>非编辑模式工具栏：</td><td>" & s_FormTB2Flag & "工具栏 " & s_FormTB2Code & "代码 " & s_FormTB2Edit & "编辑 " & s_FormTB2Text & "文本 " & s_FormTB2View & "预览 " & s_FormTB2Max & "最大化<span class=red>*</span></td><td>代码格式化：</td><td>" & s_FormCodeFormat & " <span class=red>*</span></td></tr>" & _
"<tr><td>远程文件：</td><td>" & s_FormAutoRemote & " <span class=red>*</span></td><td>显示表格虚框：</td><td>" & s_FormShowBorder & " <span class=red>*</span></td></tr>" & _
"<tr><td>显示区块：</td><td>" & s_FormShowBlock & " <span class=red>*</span></td><td>Word导入初始模式：</td><td>" & s_FormWordImportInitMode & " <span class=red>*</span></td></tr>" & _
"<tr><td>一键排版初始值：</td><td>字体：<input type=text class=input name=d_quickformatinitfontname size=10 title='字体名称' value=""" & sQuickFormatInitFontName & """> 字号：<input type=text class=input name=d_quickformatinitfontsize size=10 title='字体大小' value=""" & sQuickFormatInitFontSize & """></td><td>界面最小高度：</td><td><input type=text class=input name=d_uiminheight size=20 title='数字型' value='" & sUIMinHeight & "'> <span class=red>*</span></td></tr>" & _
"<tr><td>自动语言检测：</td><td>" & s_FormAutoDetectLanguage & " <span class=red>*</span></td><td>默认语言：</td><td>" & s_FormDefaultLanguage & " <span class=red>*</span></td></tr>" & _
"<tr><td>回车换行模式：</td><td>" & s_FormEnterMode & " <span class=red>*</span></td><td>编辑区CSS模式：</td><td>" & s_FormAreaCssMode & " <span class=red>*</span></td></tr>" & _
"<tr><td>高级接口状态：</td><td>" & s_FormAdvApiFlag & " <span class=red>*</span></td><td>安全接口加密串：</td><td><input type=text class=input size=20 name=d_encryptkey title='启用高级Session安全接口时，需要加密串，只能是字母和数字，不能有特殊字符' value=""" & sEncryptKey & """><input type=button value='随机' onclick='CreateRndEncryptKey()'></td></tr>" & _
"<tr><td>一键处理模块：</td><td>" & s_FormAutoDoneWordPaste & "Word粘贴 " & s_FormAutoDoneExcelPaste & "Excel粘贴 " & s_FormAutoDonePasteOption & "选择性粘贴 " & s_FormAutoDoneQuickFormat & "一键排版</td><td>Word导入接口选项：</td><td>" & s_FormWordImportAPI & "</td></tr>" & _
"<tr><td>备注说明：</td><td colspan=3><input type=text name=d_memo size=90 title='此样式的说明，更有利于调用' value=""" & sStyleMemo & """></td></tr>" & _
"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;上传相关设置（相关设置说明详见用户手册）：</span></td></tr>" & _
"<tr><td>上传组件：</td><td>" & s_FormUploadObject & " <span class=red>*</span></td><td>自定义FSO对象名：</td><td><input type=text class=input size=20 name=d_fsoname value=""" & sFSOname & """></td></tr>" & _
"<tr><td>自动目录顺序：</td><td>" & s_FormAutoDirOrderFlag & " <span class=red>*</span></td><td>年月日自动目录：</td><td><input type=text class=input size=18 name=d_autodir title='留空则不启用此功能，可用关键字：{yyyy}、{mm}、{dd}' value=""" & sStyleAutoDir & """> <select size=1 id=d_autodir_drop onchange='this.form.d_autodir.value=this.value'><option>-常用格式选择-</option><option value=''>不启用</option><option value='{yyyy}/'>{yyyy}/</option><option value='{yyyy}/{mm}/'>{yyyy}/{mm}/</option><option value='{yyyy}/{mm}/{dd}/'>{yyyy}/{mm}/{dd}/</option><option value='{yyyy}/{mm}{dd}/'>{yyyy}/{mm}{dd}/</option><option value='{yyyy}{mm}/'>{yyyy}{mm}/</option><option value='{yyyy}{mm}/{dd}/'>{yyyy}{mm}/{dd}/</option><option value='{yyyy}{mm}{dd}/'>{yyyy}{mm}{dd}/</option></select></td></tr>" & _
"<tr><td>文件名保存模式：</td><td>" & s_FormFileNameMode & " <span class=red>*</span></td><td>文件名同名处理：</td><td><input type=text class=input size=18 name=d_filenamesamefix title='留空则为替换已存在文件，可用关键字：{name}、{sn}、{time}' value=""" & sFileNameSameFix & """> <select size=1 id=d_filenamesamefix_drop onchange='this.form.d_filenamesamefix.value=this.value'><option>-常用格式选择-</option><option value=''>替换已存在文件</option><option value='{name}-{sn}'>自动重命名：原名+序号</option><option value='{name}-{time}'>自动重命名：原名+自动时间</option></select></td></tr>" & _
"<tr><td>上传文件浏览：</td><td>" & s_FormAllowBrowse & " <span class=red>*</span></td><td>批量上传功能启用：</td><td>" & s_FormMFUEnable & " <span class=red>*</span></td></tr>" & _
"<tr><td>批量上传接口组件：</td><td>" & s_FormMFUMode & " <span class=red>*</span></td><td>批量上传分块大小：</td><td><input type=text class=input size=20 name=d_mfublocksize title='数字型，单位KB' value=""" & sMFUBlockSize & """>KB <span class=red>*</span></td></tr>" & _
"<tr><td>路径模式：</td><td>" & s_FormBaseUrl & " <span class=red>*</span> <a href='#baseurl'>说明</a></td><td>上传路径：</td><td><input type=text class=input size=20 name=d_uploaddir title='上传文件所存放路径，相对eWebEditor根目录文件的路径' value=""" & sStyleUploadDir & """> <span class=red>*</span></td></tr>" & _
"<tr><td>显示路径：</td><td><input type=text class=input size=20 name=d_basehref title='显示内容页所存放路径，必须以&quot;/&quot;开头' value=""" & sStyleBaseHref & """></td><td>内容路径：</td><td><input type=text class=input size=20 name=d_contentpath title='实际保存在内容中的路径，相对显示路径的路径，不能以&quot;/&quot;开头' value=""" & sStyleContentPath & """></td></tr>" & _
"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;允许上传文件类型及文件大小设置（文件大小单位为KB，0表示不允许）：</span></td></tr>" & _
"<tr><td>总上传空间限制：</td><td><input type=text class=input name=d_spacesize size=20 title='数字型，单位MB，不限制请留空' value='" & sSpaceSize & "'>MB</td><td></td><td></td></tr>" & _
"<tr><td>图片类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_imageext size=30 title='用于图片相关的上传' value='" & sStyleImageExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_imagesize size=10 title='数字型，单位KB' value='" & sStyleImageSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirimage size=20 title='空表示不启用，格式如：image/' value=""" & sAutoTypeDirImage & """></td></tr>" & _
"<tr><td>Flash类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_flashext size=30 title='用于插入Flash动画' value='" & sStyleFlashExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_flashsize size=10 title='数字型，单位KB' value='" & sStyleFlashSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirflash size=20 title='空表示不启用，格式如：flash/' value=""" & sAutoTypeDirFlash & """></td></tr>" & _
"<tr><td>媒体类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_mediaext size=30 title='用于插入媒体文件' value='" & sStyleMediaExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_mediasize size=10 title='数字型，单位KB' value='" & sStyleMediaSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirmedia size=20 title='空表示不启用，格式如：media/' value=""" & sAutoTypeDirMedia & """></td></tr>" & _
"<tr><td>附件类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_fileext size=30 title='用于插入附件' value='" & sStyleFileExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_filesize size=10 title='数字型，单位KB' value='" & sStyleFileSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirattach size=20 title='空表示不启用，格式如：attach/' value=""" & sAutoTypeDirAttach & """></td></tr>" & _
"<tr><td>远程类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_remoteext size=30 title='用于自动上传远程文件' value='" & sStyleRemoteExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_remotesize size=10 title='数字型，单位KB' value='" & sStyleRemoteSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirremote size=20 title='空表示不启用，格式如：remote/' value=""" & sAutoTypeDirRemote & """></td></tr>" & _
"<tr><td>本地类型：</td><td colspan=3>文件扩展名：<input type=text class=input name=d_localext size=30 title='用于自动上传本地文件' value='" & sStyleLocalExt & "'>&nbsp;&nbsp; 文件大小限制：<input type=text class=input name=d_localsize size=10 title='数字型，单位KB' value='" & sStyleLocalSize & "'>KB&nbsp;&nbsp; 自动类型目录：<input type=text class=input name=d_autotypedirlocal size=20 title='空表示不启用，格式如：local/' value=""" & sAutoTypeDirLocal & """></td></tr>" & _
"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;分页相关设置（前台显示页应作相应处理以识别分页符）：</span></td></tr>" & _
"<tr><td>分页符模式：</td><td>" & s_FormPaginationMode & " <span class=red>*</span></td><td>自定义分页符关键字：</td><td><input type=text class=input size=20 name=d_paginationkey title='' value=""" & sPaginationKey & """></td></tr>" & _
"<tr><td>提交内容自动分页：</td><td>" & s_FormPaginationAutoFlag & " <span class=red>*</span></td><td>自动分页字数：</td><td><input type=text class=input size=20 name=d_paginationautonum title='当启用自动分页时，将依此值进行自动分页' value=""" & sPaginationAutoNum & """></td></tr>" & _
"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;缩略图及水印相关设置：</span></td></tr>" & _
"<tr><td>图形处理组件：</td><td>" & s_FormSLTSYObject & "</td><td>处理图形扩展名：</td><td><input type=text name=d_sltsyext size=20 class=input value=""" & sSLTSYExt & """></td></tr>" & _
"<tr><td>缩略图使用状态：</td><td>" & s_FormSLTFlag & "</td><td>缩略图生成模式：</td><td>" & s_FormSLTMode & "</td></tr>" & _
"<tr><td>缩略图长度条件：</td><td>" & s_FormSLTCheckFlag & "大于<input type=text name=d_sltminsize size=10 class=input title='图形的长度只有达到此最小长度要求时才会生成缩略图，数字型' value='" & sSLTMinSize & "'>px</td><td>缩略图生成长度：</td><td><input type=text name=d_sltoksize size=20 class=input title='生成的缩略图长度值，数字型' value='" & sSLTOkSize & "'>px</td></tr>" & _
"<tr><td>水印有效模块：</td><td>" & s_FormSYValidNormal & "普通上传 " & s_FormSYValidLocal & "本地上传 " & s_FormSYValidRemote & "远程上传</td><td></td><td></td></tr>" & _
"<tr><td>文字水印使用状态：</td><td>" & s_FormSYWZFlag & "</td><td>文字水印启用条件：</td><td>宽:<input type=text name=d_sywzminwidth size=4 class=input title='图形的宽度只有达到此最小宽度要求时才会生成水印，数字型' value='" & sSYWZMinWidth & "'>px&nbsp; 高:<input type=text name=d_sywzminheight size=4 class=input title='图形的高度只有达到此最小高度要求时才会生成水印，数字型' value='" & sSYWZMinHeight & "'>px</td></tr>" & _
"<tr><td>文字水印内容：</td><td><input type=text name=d_sytext size=20 class=input title='当使用文字水印时的文字内容' value=""" & sSYText & """></td><td>文字水印字体颜色：</td><td><input type=text name=d_syfontcolor size=20 class=input title='当使用文字水印时文字的颜色' value=""" & sSYFontColor & """></td></tr>" & _
"<tr><td>文字水印阴影颜色：</td><td><input type=text name=d_syshadowcolor size=20 class=input title='当使用文字水印时的文字阴影颜色' value=""" & sSYShadowColor & """></td><td>文字水印阴影大小：</td><td><input type=text name=d_syshadowoffset size=20 class=input title='当使用文字水印时文字的阴影大小' value=""" & sSYShadowOffset & """>px</td></tr>" & _
"<tr><td>文字水印字体大小：</td><td><input type=text name=d_syfontsize size=20 class=input title='当使用文字水印时文字的字体大小' value=""" & sSYFontSize & """>px</td><td>文字水印字体名称：</td><td><input type=text name=d_syfontname size=20 class=input title='当使用文字水印时文字的字体名' value=""" & sSYFontName & """></td></tr>" & _
"<tr><td>文字水印位置：</td><td>" & s_FormSYWZPosition & "</td><td>文字水印边距：</td><td>左右:<input type=text name=d_sywzpaddingh size=4 class=input title='居左时作用为左边距，居右时作用为右边距，数字型' value='" & sSYWZPaddingH & "'>px&nbsp; 上下:<input type=text name=d_sywzpaddingv size=4 class=input title='居上时作用为上边距，居下时作用为下边柜，数字型' value='" & sSYWZPaddingV & "'>px</td></tr>" & _
"<tr><td>文字水印文字占位：</td><td>宽:<input type=text name=d_sywztextwidth size=4 class=input title='水印文字的占位宽度，由字数、字体大小等设置的效果确定，数字型' value='" & sSYWZTextWidth & "'>px&nbsp; 高:<input type=text name=d_sywztextheight size=4 class=input title='水印文字的占位高度，由字数、字体大小等设置的效果确定，数字型' value='" & sSYWZTextHeight & "'>px&nbsp; <input type=button value='检测宽高' onclick='doCheckWH(1)'></td><td></td><td></td></tr>" & _
"<tr><td>图片水印使用状态：</td><td>" & s_FormSYTPFlag & "</td><td>图片水印启用条件：</td><td>宽:<input type=text name=d_sytpminwidth size=4 class=input title='图形的宽度只有达到此最小宽度要求时才会生成水印，数字型' value='" & sSYTPMinWidth & "'>px&nbsp; 高:<input type=text name=d_sytpminheight size=4 class=input title='图形的高度只有达到此最小高度要求时才会生成水印，数字型' value='" & sSYTPMinHeight & "'>px</td></tr>" & _
"<tr><td>图片水印位置：</td><td>" & s_FormSYTPPosition & "</td><td>图片水印边距：</td><td>左右:<input type=text name=d_sytppaddingh size=4 class=input title='居左时作用为左边距，居右时作用为右边距，数字型' value='" & sSYTPPaddingH & "'>px&nbsp; 上下:<input type=text name=d_sytppaddingv size=4 class=input title='居上时作用为上边距，居下时作用为下边柜，数字型' value='" & sSYTPPaddingV & "'>px</td></tr>" & _
"<tr><td>图片水印图片路径：</td><td><input type=text name=d_sypicpath size=20 class=input title='当使用图片水印时图片的路径' value=""" & sSYPicPath & """></td><td>图片水印透明度：</td><td><input type=text name=d_sytpopacity size=20 class=input title='0至1间的数字，如0.5表示半透明' value=""" & sSYTPOpacity & """></td></tr>" & _
"<tr><td>图片水印图片占位：</td><td>宽:<input type=text name=d_sytpimagewidth size=4 class=input title='水印图片的宽度，数字型' value='" & sSYTPImageWidth & "'>px&nbsp; 高:<input type=text name=d_sytpimageheight size=4 class=input title='水印图片的高度，数字型' value='" & sSYTPImageHeight & "'>px&nbsp; <input type=button value='检测宽高' onclick='doCheckWH(2)'></td><td></td><td></td></tr>" & _
"<tr><td>水印宽高检测区：</td><td colspan=3><span id=tdPreview></span></td></tr>" & _
"<tr><td align=center colspan=4><input type=submit value='  提交  ' align=absmiddle>&nbsp;<input type=reset name=btnReset value='  重填  '></td></tr>" & _
"</form>" & _
"</table><br>"
Dim sMsg
sMsg = "<a name=baseurl></a><p><span class=blue><b>路径模式设置说明：</b></span><br>" & _
"<b>相对路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""UploadFile/...""或""../UploadFile/...""形式呈现，当使用此模式时，显示路径和内容路径必填，显示路径必须以""/""开头和结尾，内容路径设置中不能以""/""开头。<br>" & _
"<b>绝对根路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""/eWebEditor/UploadFile/...""这种形式呈现，当使用此模式时，显示路径和内容路径不必填。<br>" & _
"<b>绝对全路径：</b>指所有的相关上传或自动插入文件路径，编辑后都以""http://xxx.xxx.xxx/eWebEditor/UploadFile/...""这种形式呈现，当使用此模式时，显示路径和内容路径不必填。<br>" & _
"<b>站外绝对全路径：</b>当使用此模式时，上传路径必须是实际物理路径，如：""c:\xxx\""；显示路径为空；内容路径必须以""http""开头。</p>"
Call ShowMessage(sMsg)
End Sub
Sub InitStyle()
Dim b, aCurrStyle
b = False
sStyleID = Trim(Request("id"))
If IsNumeric(sStyleID) = True Then
nStyleID = Clng(sStyleID)
If nStyleID <= Ubound(aStyle) Then
aCurrStyle = Split(aStyle(nStyleID), "|||")
sStyleName = aCurrStyle(0)
sFixWidth = aCurrStyle(1)
sSkin = aCurrStyle(2)
sStyleUploadDir = aCurrStyle(3)
sStyleWidth = aCurrStyle(4)
sStyleHeight = aCurrStyle(5)
sStyleFileExt = aCurrStyle(6)
sStyleFlashExt = aCurrStyle(7)
sStyleImageExt = aCurrStyle(8)
sStyleMediaExt = aCurrStyle(9)
sStyleRemoteExt = aCurrStyle(10)
sStyleFileSize = aCurrStyle(11)
sStyleFlashSize = aCurrStyle(12)
sStyleImageSize = aCurrStyle(13)
sStyleMediaSize = aCurrStyle(14)
sStyleRemoteSize = aCurrStyle(15)
sStyleStateFlag = aCurrStyle(16)
sStyleDetectFromWord = aCurrStyle(17)
sStyleInitMode = aCurrStyle(18)
sStyleBaseUrl = aCurrStyle(19)
sStyleUploadObject = aCurrStyle(20)
sFSOname = aCurrStyle(21)
sStyleBaseHref = aCurrStyle(22)
sStyleContentPath = aCurrStyle(23)
sStyleAutoRemote = aCurrStyle(24)
sStyleShowBorder = aCurrStyle(25)
sStyleMemo = aCurrStyle(26)
sAutoDetectLanguage = aCurrStyle(27)
sDefaultLanguage = aCurrStyle(28)
sSLTFlag = aCurrStyle(29)
sSLTMinSize = aCurrStyle(30)
sSLTOkSize = aCurrStyle(31)
sSYWZFlag = aCurrStyle(32)
sSYText = aCurrStyle(33)
sSYFontColor = aCurrStyle(34)
sSYFontSize = aCurrStyle(35)
sSYFontName = aCurrStyle(36)
sSYPicPath = aCurrStyle(37)
sSLTSYObject = aCurrStyle(38)
sSLTSYExt = aCurrStyle(39)
sSYWZMinWidth = aCurrStyle(40)
sSYShadowColor = aCurrStyle(41)
sSYShadowOffset = aCurrStyle(42)
sStyleAllowBrowse = aCurrStyle(43)
sStyleLocalExt = aCurrStyle(44)
sStyleLocalSize = aCurrStyle(45)
sSYWZMinHeight = aCurrStyle(46)
sSYWZPosition = aCurrStyle(47)
sSYWZTextWidth = aCurrStyle(48)
sSYWZTextHeight = aCurrStyle(49)
sSYWZPaddingH = aCurrStyle(50)
sSYWZPaddingV = aCurrStyle(51)
sSYTPFlag = aCurrStyle(52)
sSYTPMinWidth = aCurrStyle(53)
sSYTPMinHeight = aCurrStyle(54)
sSYTPPosition = aCurrStyle(55)
sSYTPPaddingH = aCurrStyle(56)
sSYTPPaddingV = aCurrStyle(57)
sSYTPImageWidth = aCurrStyle(58)
sSYTPImageHeight = aCurrStyle(59)
sSYTPOpacity = aCurrStyle(60)
sAdvApiFlag = aCurrStyle(61)
sSBCode = aCurrStyle(62)
sSBEdit = aCurrStyle(63)
sSBText = aCurrStyle(64)
sSBView = aCurrStyle(65)
sEnterMode = aCurrStyle(66)
sAreaCssMode = aCurrStyle(67)
sFileNameMode = aCurrStyle(68)
sSLTMode = aCurrStyle(69)
sEncryptKey = aCurrStyle(70)
sStyleAutoDir = aCurrStyle(71)
sPaginationMode = aCurrStyle(72)
sPaginationKey = aCurrStyle(73)
sPaginationAutoFlag = aCurrStyle(74)
sPaginationAutoNum = aCurrStyle(75)
sSBSize = aCurrStyle(76)
sSLTCheckFlag = aCurrStyle(77)
sSpaceSize = aCurrStyle(78)
sMFUMode = aCurrStyle(79)
sMFUBlockSize = aCurrStyle(80)
sMFUEnable = aCurrStyle(81)
sCodeFormat = aCurrStyle(82)
sTB2Flag = aCurrStyle(83)
sTB2Code = aCurrStyle(84)
sTB2Max = aCurrStyle(85)
sShowBlock = aCurrStyle(86)
sFileNameSameFix = aCurrStyle(87)
sAutoDirOrderFlag = aCurrStyle(88)
sAutoTypeDirImage = aCurrStyle(89)
sAutoTypeDirFlash = aCurrStyle(90)
sAutoTypeDirMedia = aCurrStyle(91)
sAutoTypeDirAttach = aCurrStyle(92)
sAutoTypeDirRemote = aCurrStyle(93)
sAutoTypeDirLocal = aCurrStyle(94)
sWordImportInitMode = aCurrStyle(95)
sQuickFormatInitFontName = aCurrStyle(96)
sQuickFormatInitFontSize = aCurrStyle(97)
sUIMinHeight = aCurrStyle(98)
sSYValidNormal = aCurrStyle(99)
sSYValidLocal = aCurrStyle(100)
sSYValidRemote = aCurrStyle(101)
sAutoDoneWordPaste = aCurrStyle(102)
sAutoDoneExcelPaste = aCurrStyle(103)
sAutoDoneQuickFormat = aCurrStyle(104)
sWordImportAPI = aCurrStyle(105)
sAutoDonePasteOption = aCurrStyle(106)
sTB2Edit = aCurrStyle(107)
sTB2Text = aCurrStyle(108)
sTB2View = aCurrStyle(109)
b = True
End If
End If
If b = False Then
GoError "无效的样式ID号，请通过页面上的链接进行操作！"
End If
End Sub
Sub CheckStyleForm()
sStyleName = Trim(Request("d_name"))
sFixWidth = Trim(Request("d_fixwidth"))
sSkin = Trim(Request("d_skin"))
sStyleUploadDir = Trim(Request("d_uploaddir"))
sStyleWidth = Trim(Request("d_width"))
sStyleHeight = Trim(Request("d_height"))
sStyleFileExt = Trim(Request("d_fileext"))
sStyleFlashExt = Trim(Request("d_flashext"))
sStyleImageExt = Trim(Request("d_imageext"))
sStyleMediaExt = Trim(Request("d_mediaext"))
sStyleRemoteExt = Trim(Request("d_remoteext"))
sStyleFileSize = Trim(Request("d_filesize"))
sStyleFlashSize = Trim(Request("d_flashsize"))
sStyleImageSize = Trim(Request("d_imagesize"))
sStyleMediaSize = Trim(Request("d_mediasize"))
sStyleRemoteSize = Trim(Request("d_remotesize"))
sStyleStateFlag = Trim(Request("d_stateflag"))
sStyleDetectFromWord = Trim(Request("d_detectfromword"))
sStyleInitMode = Trim(Request("d_initmode"))
sStyleBaseUrl = Trim(Request("d_baseurl"))
sStyleUploadObject = Trim(Request("d_uploadobject"))
sFSOname = Trim(Request("d_fsoname"))
sStyleBaseHref = Trim(Request("d_basehref"))
sStyleContentPath = Trim(Request("d_contentpath"))
sStyleAutoRemote = Trim(Request("d_autoremote"))
sStyleShowBorder = Trim(Request("d_showborder"))
sStyleMemo = Trim(Request("d_memo"))
sSLTFlag = Trim(Request("d_sltflag"))
sSLTMinSize = Trim(Request("d_sltminsize"))
sSLTOkSize = Trim(Request("d_sltoksize"))
sSYWZFlag = Trim(Request("d_sywzflag"))
sSYText = Trim(Request("d_sytext"))
sSYFontColor = Trim(Request("d_syfontcolor"))
sSYFontSize = Trim(Request("d_syfontsize"))
sSYFontName = Trim(Request("d_syfontname"))
sSYPicPath = Trim(Request("d_sypicpath"))
sSLTSYObject = Trim(Request("d_sltsyobject"))
sSLTSYExt = Trim(Request("d_sltsyext"))
sSYWZMinWidth = Trim(Request("d_sywzminwidth"))
sSYShadowColor = Trim(Request("d_syshadowcolor"))
sSYShadowOffset = Trim(Request("d_syshadowoffset"))
sStyleAllowBrowse = Trim(Request("d_allowbrowse"))
sStyleLocalExt = Trim(Request("d_localext"))
sStyleLocalSize = Trim(Request("d_localsize"))
sSYWZMinHeight = Trim(Request("d_sywzminheight"))
sSYWZPosition = Trim(Request("d_sywzposition"))
sSYWZTextWidth = Trim(Request("d_sywztextwidth"))
sSYWZTextHeight = Trim(Request("d_sywztextheight"))
sSYWZPaddingH = Trim(Request("d_sywzpaddingh"))
sSYWZPaddingV = Trim(Request("d_sywzpaddingv"))
sSYTPFlag = Trim(Request("d_sytpflag"))
sSYTPMinWidth = Trim(Request("d_sytpminwidth"))
sSYTPMinHeight = Trim(Request("d_sytpminheight"))
sSYTPPosition = Trim(Request("d_sytpposition"))
sSYTPPaddingH = Trim(Request("d_sytppaddingh"))
sSYTPPaddingV = Trim(Request("d_sytppaddingv"))
sSYTPImageWidth = Trim(Request("d_sytpimagewidth"))
sSYTPImageHeight = Trim(Request("d_sytpimageheight"))
sSYTPOpacity = Trim(Request("d_sytpopacity"))
sAdvApiFlag = Trim(Request("d_advapiflag"))
sSBCode = Trim(Request("d_sbcode"))
sSBEdit = Trim(Request("d_sbedit"))
sSBText = Trim(Request("d_sbtext"))
sSBView = Trim(Request("d_sbview"))
sEnterMode = Trim(Request("d_entermode"))
sAreaCssMode = Trim(Request("d_areacssmode"))
sFileNameMode = Trim(Request("d_filenamemode"))
sSLTMode = Trim(Request("d_sltmode"))
sEncryptKey = Trim(Request("d_encryptkey"))
sStyleAutoDir = Trim(Request("d_autodir"))
sPaginationMode = Trim(Request("d_paginationmode"))
sPaginationKey = Trim(Request("d_paginationkey"))
sPaginationAutoFlag = Trim(Request("d_paginationautoflag"))
sPaginationAutoNum = Trim(Request("d_paginationautonum"))
sSBSize = Trim(Request("d_sbsize"))
sAutoDetectLanguage = Trim(Request("d_autodetectlanguage"))
sDefaultLanguage = Trim(Request("d_defaultlanguage"))
sSLTCheckFlag = Trim(Request("d_sltcheckflag"))
sSpaceSize = Trim(Request("d_spacesize"))
sMFUMode = Trim(Request("d_mfumode"))
sMFUBlockSize = Trim(Request("d_mfublocksize"))
sMFUEnable = Trim(Request("d_mfuenable"))
sCodeFormat = Trim(Request("d_codeformat"))
sTB2Flag = Trim(Request("d_tb2flag"))
sTB2Code = Trim(Request("d_tb2code"))
sTB2Edit = Trim(Request("d_tb2edit"))
sTB2Text = Trim(Request("d_tb2text"))
sTB2View = Trim(Request("d_tb2view"))
sTB2Max = Trim(Request("d_tb2max"))
sShowBlock = Trim(Request("d_showblock"))
sFileNameSameFix = Trim(Request("d_filenamesamefix"))
sAutoDirOrderFlag = Trim(Request("d_autodirorderflag"))
sAutoTypeDirImage = Trim(Request("d_autotypedirimage"))
sAutoTypeDirFlash = Trim(Request("d_autotypedirflash"))
sAutoTypeDirMedia = Trim(Request("d_autotypedirmedia"))
sAutoTypeDirAttach = Trim(Request("d_autotypedirattach"))
sAutoTypeDirRemote = Trim(Request("d_autotypedirremote"))
sAutoTypeDirLocal = Trim(Request("d_autotypedirlocal"))
sWordImportInitMode = Trim(Request("d_wordimportinitmode"))
sWordImportAPI = Trim(Request("d_wordimportapi"))
sQuickFormatInitFontName = Trim(Request("d_quickformatinitfontname"))
sQuickFormatInitFontSize = Trim(Request("d_quickformatinitfontsize"))
sUIMinHeight = Trim(Request("d_uiminheight"))
sSYValidNormal = Trim(Request("d_syvalidnormal"))
sSYValidLocal = Trim(Request("d_syvalidlocal"))
sSYValidRemote = Trim(Request("d_syvalidremote"))
sAutoDoneWordPaste = Trim(Request("d_autodonewordpaste"))
sAutoDoneExcelPaste = Trim(Request("d_autodoneexcelpaste"))
sAutoDoneQuickFormat = Trim(Request("d_autodonequickformat"))
sAutoDonePasteOption = Trim(Request("d_autodonepasteoption"))
End Sub
Sub DoStyleAddSave()
If StyleName2ID(sStyleName) <> -1 Then
GoError "此样式名已经存在，请用另一个样式名！"
End If
Dim nNewStyleID
nNewStyleID = Ubound(aStyle) + 1
Redim Preserve aStyle(nNewStyleID)
aStyle(nNewStyleID) = GetStyleDataString()
Call WriteConfig()
Call ShowMessage("<b><span class=red>样式增加成功！</span></b><li><a href='?action=toolbar&id=" & nNewStyleID & "'>设置此样式下的工具栏</a>")
End Sub
Sub DoStyleSetSave()
Dim n, s_OldStyleName
sStyleID = Trim(Request("id"))
If IsNumeric(sStyleID) = True Then
n = StyleName2ID(sStyleName)
If CStr(n) <> sStyleID And n <> -1 Then
GoError "此样式名已经存在，请用另一个样式名！"
End If
If Clng(sStyleID) < 1 And Clng(sStyleID)>UBound(aStyle) Then
GoError "无效的样式ID号，请通过页面上的链接进行操作！"
End If
s_OldStyleName = Split(aStyle(Clng(sStyleID)), "|||")(0)
aStyle(Clng(sStyleID)) = GetStyleDataString()
Else
GoError "无效的样式ID号，请通过页面上的链接进行操作！"
End If
Call WriteConfig()
Call ShowMessage("<b><span class=red>样式修改成功！</span></b><li><a href='?action=stylepreview&id=" & sStyleID & "' target='_blank'>预览此样式</a><li><a href='?action=toolbar&id=" & sStyleID & "'>设置此样式下的工具栏</a><li><a href='?action=styleset&id=" & sStyleID & "'>重新设置此样式</a>")
End Sub
Function GetStyleDataString()
GetStyleDataString = sStyleName & "|||" & sFixWidth & "|||" & sSkin & "|||" & sStyleUploadDir & "|||" & sStyleWidth & "|||" & sStyleHeight & "|||" & sStyleFileExt & "|||" & sStyleFlashExt & "|||" & sStyleImageExt & "|||" & sStyleMediaExt & "|||" & sStyleRemoteExt & "|||" & sStyleFileSize & "|||" & sStyleFlashSize & "|||" & sStyleImageSize & "|||" & sStyleMediaSize & "|||" & sStyleRemoteSize & "|||" & sStyleStateFlag & "|||" & sStyleDetectFromWord & "|||" & sStyleInitMode & "|||" & sStyleBaseUrl & "|||" & sStyleUploadObject & "|||" & sFSOname & "|||" & sStyleBaseHref & "|||" & sStyleContentPath & "|||" & sStyleAutoRemote & "|||" & sStyleShowBorder & "|||" & sStyleMemo & "|||" & sAutoDetectLanguage & "|||" & sDefaultLanguage & "|||" & sSLTFlag & "|||" & sSLTMinSize & "|||" & sSLTOkSize & "|||" & sSYWZFlag & "|||" & sSYText & "|||" & sSYFontColor & "|||" & sSYFontSize & "|||" & sSYFontName & "|||" & sSYPicPath & "|||" & sSLTSYObject & "|||" & sSLTSYExt & "|||" & sSYWZMinWidth & "|||" & sSYShadowColor & "|||" & sSYShadowOffset & "|||" & sStyleAllowBrowse & "|||" & sStyleLocalExt & "|||" & sStyleLocalSize & "|||" & sSYWZMinHeight & "|||" & sSYWZPosition & "|||" & sSYWZTextWidth & "|||" & sSYWZTextHeight & "|||" & sSYWZPaddingH & "|||" & sSYWZPaddingV & "|||" & sSYTPFlag & "|||" & sSYTPMinWidth & "|||" & sSYTPMinHeight & "|||" & sSYTPPosition & "|||" & sSYTPPaddingH & "|||" & sSYTPPaddingV & "|||" & sSYTPImageWidth & "|||" & sSYTPImageHeight & "|||" & sSYTPOpacity & "|||" & sAdvApiFlag  & "|||" & sSBCode  & "|||" & sSBEdit  & "|||" & sSBText  & "|||" & sSBView & "|||" & sEnterMode & "|||" & sAreaCssMode & "|||" & sFileNameMode & "|||" & sSLTMode & "|||" & sEncryptKey & "|||" & sStyleAutoDir & "|||" & sPaginationMode & "|||" & sPaginationKey & "|||" & sPaginationAutoFlag & "|||" & sPaginationAutoNum & "|||" & sSBSize & "|||" & sSLTCheckFlag & "|||" & sSpaceSize & "|||" & sMFUMode & "|||" & sMFUBlockSize & "|||" & sMFUEnable & "|||" & sCodeFormat & "|||" & sTB2Flag & "|||" & sTB2Code & "|||" & sTB2Max & "|||" & sShowBlock & "|||" & sFileNameSameFix & "|||" & sAutoDirOrderFlag & "|||" & sAutoTypeDirImage & "|||" & sAutoTypeDirFlash & "|||" & sAutoTypeDirMedia & "|||" & sAutoTypeDirAttach & "|||" & sAutoTypeDirRemote & "|||" & sAutoTypeDirLocal & "|||" & sWordImportInitMode & "|||" & sQuickFormatInitFontName & "|||" & sQuickFormatInitFontSize & "|||" & sUIMinHeight & "|||" & sSYValidNormal & "|||" & sSYValidLocal & "|||" & sSYValidRemote & "|||" & sAutoDoneWordPaste & "|||" & sAutoDoneExcelPaste & "|||" & sAutoDoneQuickFormat & "|||" & sWordImportAPI & "|||" & sAutoDonePasteOption & "|||" & sTB2Edit & "|||" & sTB2Text & "|||" & sTB2View
End Function
Sub DoStyleDel()
aStyle(Clng(sStyleID)) = ""
Call WriteConfig()
Call GoUrl("?")
End Sub
Sub ShowStylePreview()
Response.Write "<html><head>" & _
"<title>样式预览</title>" & _
"<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & _
"</head><body>" & _
"<input type=hidden name=content1  value=''>" & _
"<iframe ID='eWebEditor1' src='../ewebeditor.htm?id=content1&style=" & sStyleName & "' frameborder=0 scrolling=no width='" & sStyleWidth & "' HEIGHT='" & sStyleHeight & "'></iframe>" & _
"</body></html>"
End Sub
Sub ShowStyleCode()
Response.Write "<table border=0 cellspacing=1 align=center class=list>" & _
"<tr><th>样式（" & outHTML(sStyleName) & "）的最佳调用代码如下（其中XXX按实际关联的表单项进行修改）：</th></tr>" & _
"<tr><td><textarea rows=5 cols=65 style='width:100%'><IFRAME ID=""eWebEditor1"" SRC=""ewebeditor.htm?id=XXX&style=" & sStyleName & """ FRAMEBORDER=""0"" SCROLLING=""no"" WIDTH=""" & sStyleWidth & """ HEIGHT=""" & sStyleHeight & """></IFRAME></textarea></td></tr>" & _
"</table>"
End Sub
Sub ShowToolBarList()
Call ShowMessage("<b class=blue>样式（" & outHTML(sStyleName) & "）下的工具栏管理：</b>")
Dim s_AddForm, s_ModiForm, i, aCurrToolbar
Dim nMaxOrder
nMaxOrder = 0
For i = 1 To UBound(aToolbar)
aCurrToolbar = Split(aToolbar(i), "|||")
If aCurrToolbar(0) = sStyleID Then
If Clng(aCurrToolbar(3)) > nMaxOrder Then
nMaxOrder = Clng(aCurrToolbar(3))
End If
End If
Next
nMaxOrder = nMaxOrder + 1
s_AddForm = "<hr width='80%' align=center size=1><table border=0 cellpadding=4 cellspacing=0 align=center>" & _
"<form action='?id=" & sStyleID & "&action=toolbaradd' name='addform' method=post>" & _
"<tr><td>工具栏名：<input type=text name=d_name size=20 class=input value='工具栏" & nMaxOrder & "'> 排序号：<input type=text name=d_order size=5 value='" & nMaxOrder & "' class=input> <input type=submit name=b1 value='新增工具栏'></td></tr>" & _
"</form></table><hr width='80%' align=center size=1>"
Dim s_Manage
s_ModiForm = "<form action='?id=" & sStyleID & "&action=toolbarmodi' name=modiform method=post>" & _
"<table border=0 cellpadding=0 cellspacing=1 align=center class=form>" & _
"<tr align=center><th>ID</th><th>工具栏名</th><th>排序号</th><th>操作</th></tr>"
For i = 1 To UBound(aToolbar)
aCurrToolbar = Split(aToolbar(i), "|||")
If aCurrToolbar(0) = sStyleID Then
s_Manage = "<a href='?id=" & sStyleID & "&action=buttonset&toolbarid=" & i & "'>按钮设置</a>"
s_Manage = s_Manage & "|<a href='?id=" & sStyleID & "&action=toolbardel&delid=" & i & "'>删除</a>"
s_ModiForm = s_ModiForm & "<tr align=center>" & _
"<td>" & i & "</td>" & _
"<td><input type=text name='d_name" & i & "' value=""" & inHTML(aCurrToolbar(2)) & """ size=30 class=input></td>" & _
"<td><input type=text name='d_order" & i & "' value='" & aCurrToolbar(3) & "' size=5 class=input></td>" & _
"<td>" & s_Manage & "</td>" & _
"</tr>"
End If
Next
s_ModiForm = s_ModiForm & "<tr><td colspan=4 align=center><input type=submit name=b1 value='  修改  '></td></tr></table></form>"
Response.Write s_AddForm & s_ModiForm
End Sub
Sub DoToolBarAdd()
Dim s_Name, s_Order
s_Name = Trim(Request("d_name"))
s_Order = Trim(Request("d_order"))
If s_Name = "" Then
GoError "工具栏名不能为空！"
End If
If IsNumeric(s_Order) = False Then
GoError "无效的工具栏排序号，排序号必须为数字！"
End If
Dim nToolbarNum
nToolbarNum = Ubound(aToolbar) + 1
Redim Preserve aToolbar(nToolbarNum)
aToolbar(nToolbarNum) = sStyleID & "||||||" & s_Name & "|||" & s_Order
Call WriteConfig()
Response.Write "<script language=javascript>alert(""工具栏（" & outHTML(s_Name) & "）增加操作成功！"");</script>"
Call GoUrl("?action=toolbar&id=" & sStyleID)
End Sub
Sub DoToolBarModi()
Dim s_Name, s_Order, i, aCurrToolbar
For i = 1 To UBound(aToolbar)
aCurrToolbar = Split(aToolbar(i), "|||")
If aCurrToolbar(0) = sStyleID Then
s_Name = Trim(Request("d_name" & i))
s_Order = Trim(Request("d_order" & i))
If s_Name = "" Or IsNumeric(s_Order) = False Then 
aCurrToolbar(0) = ""
s_Name = ""
End If
aToolbar(i) = aCurrToolbar(0) & "|||" & aCurrToolbar(1) & "|||" & s_Name & "|||" & s_Order
End If
Next
Call WriteConfig()
Response.Write "<script language=javascript>alert('工具栏修改操作成功！');</script>"
Call GoUrl("?action=toolbar&id=" & sStyleID)
End Sub
Sub DoToolBarDel()
Dim s_DelID
s_DelID = Trim(Request("delid"))
If IsNumeric(s_DelID) = True Then
aToolbar(Clng(s_DelID)) = ""
Call WriteConfig()
Response.Write "<script language=javascript>alert('工具栏（ID：" & s_DelID & "）删除操作成功！');</script>"
Call GoUrl("?action=toolbar&id=" & sStyleID)
End If
End Sub
Sub InitToolBar()
Dim b, aCurrToolbar, nToolbarID
b = False
sToolBarID = Trim(Request("toolbarid"))
If IsNumeric(sToolBarID) = True Then
If Clng(sToolBarID) <= UBound(aToolbar) And Clng(sToolBarID) > 0 Then
aCurrToolbar = Split(aToolbar(Clng(sToolbarID)), "|||")
sToolBarName = aCurrToolbar(2)
sToolBarOrder = aCurrToolbar(3)
sToolBarButton = aCurrToolbar(1)
b = True
End If
End If
If b = False Then
GoError "无效的工具栏ID号，请通过页面上的链接进行操作！"
End If
End Sub
Sub ShowButtonList()
Call ShowMessage("<b class=blue>当前样式：<span class=red>" & outHTML(sStyleName) & "</span>&nbsp;&nbsp;当前工具栏：<span class=red>" & outHTML(sToolBarName) & "</span></b>")
%>
<script language='javascript' src='../js/buttons.js'></script>
<script language='javascript' src='../js/zh-cn.js'></script>
<table border=0 cellpadding=5 cellspacing=0 align=center>
<form action='?action=buttonsave&id=<%=sStyleID%>&toolbarid=<%=sToolBarID%>' method=post name=myform>
<tr align=center><td>可选按钮</td><td></td><td>已选按钮</td><td></td></tr>
<tr>
<td><DIV id=div1 style="BORDER-RIGHT: 1.5pt inset; PADDING-RIGHT: 0px; BORDER-TOP: 1.5pt inset; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; OVERFLOW: auto; BORDER-LEFT: 1.5pt inset; WIDTH: 250px; PADDING-TOP: 0px; BORDER-BOTTOM: 1.5pt inset; HEIGHT: 350px; BACKGROUND-COLOR: white"></DIV></td>
<td><input type=button name=b1 value=' → ' onclick='Add()'><br><br><input type=button name=b1 value=' ← ' onclick='Del()'></td>
<td><DIV id=div2 style="BORDER-RIGHT: 1.5pt inset; PADDING-RIGHT: 0px; BORDER-TOP: 1.5pt inset; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; OVERFLOW: auto; BORDER-LEFT: 1.5pt inset; WIDTH: 250px; PADDING-TOP: 0px; BORDER-BOTTOM: 1.5pt inset; HEIGHT: 350px; BACKGROUND-COLOR: white"></DIV></td>
<td><input type=button name=b3 value='↑' onclick='Up()'><br><br><br><input type=button name=b4 value='↓' onclick='Down()'></td>
</tr>
<input type=hidden name='d_button' value='<%=sToolBarButton%>'>
<tr><td colspan=4 align=right><input type=submit name=b value=' 保存设置 '></td></tr>
</form>
</table>
<script language=javascript>
initButtonOptions("<%=sSkin%>");
</script>
<%
Call ShowMessage("<b class=blue>提示：</b>你可以通过按“Ctrl”“Shift”来快速多选定，可以在指定项上“双击”快速增加或删除项。可以选定多个按钮同时上移或下移操作。")
End Sub
Sub DoButtonSave()
Dim s_Button, nToolBarID, aCurrToolbar
s_Button = Trim(Request("d_button"))
nToolBarID = Clng(sToolBarID)
aCurrToolbar = Split(aToolbar(nToolBarID), "|||")
aToolbar(nToolBarID) = aCurrToolbar(0) & "|||" & s_Button & "|||" & aCurrToolbar(2) & "|||" & aCurrToolbar(3)
Call WriteConfig()
Call ShowMessage("<b><span class=red>工具栏按钮设置保存成功！</span></b><li><a href='?action=stylepreview&id=" & sStyleID & "' target='_blank'>预览此样式</a><li><a href='?action=toolbar&id=" & sStyleID & "'>返回工具栏管理</a><li><a href='?action=buttonset&id=" & sStyleID & "&toolbarid=" & sToolBarID & "'>重新设置此工具栏下的按钮</a>")
End Sub
Sub GoUrl(url)
Response.Write "<script language=javascript>location.href=""" & url & """;</script>"
Response.End
End Sub
%>
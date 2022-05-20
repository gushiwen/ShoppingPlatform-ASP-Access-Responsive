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

sPosition = sPosition & "设置授权序列号"

Call Header()
Call Content()
Call Footer()


Sub Content()
	Select Case sAction
	Case "MODI"
		Call DoModi()
	Case Else
		Call ShowForm()
	End Select
End Sub


Sub ShowForm()
	%>
	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=form>
	<form action='?action=modi' method=post name=myform>
	<tr>
		<th>设置名称</th>
		<th>基本参数设置</th>
		<th>设置说明</th>
	</tr>
	<tr>
		<td width="15%">序列号：</td>
		<td width="55%"><input type=text class=input size=60 name=d_license value="<%=inHTML(sLicense)%>"></td>
		<td width="30%">多个以分号隔开，<a href="http://service.ewebeditor.net/" target="_blank">从客服中心获取</a></td>
	</tr>
	<tr><td align=center colspan=3><input type=submit name=bSubmit value="  提交  "></a>&nbsp;<input type=reset name=bReset value="  重填  "></td></tr>
	</form>
	</table>


	<%
End Sub

Sub DoModi()

	Dim sNewLicense
	sNewLicense = Trim(Request("d_license"))

	If InStr(sNewLicense, "'")>0 Or InStr(sNewLicense, """")>0 Then
		GoError "序列号不能含有特殊字符！"
	End If


	sLicense = sNewLicense

	Call WriteConfig()

	%>
	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=list>
	<tr>
		<td>授权序列号设置成功！</td>
	</tr>
	</table>
	<%
End Sub

%>
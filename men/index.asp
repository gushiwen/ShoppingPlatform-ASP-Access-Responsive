<%@ Language=VBScript %>
<%
	' 跳转到 main.asp
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", "main.asp"
%>

<%@ Language=VBScript %>
<%
	' 跳转到 main.html
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", "main.html"
%>

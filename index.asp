<%@ Language=VBScript %>
<%
	' 跳转到英文首页
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", "en/main.html"
%>

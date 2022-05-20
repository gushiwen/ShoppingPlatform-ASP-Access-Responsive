<!--#include file="goldweb_text.asp"-->
<%
' 搜索商品
'SAFARI javascript history.go(-1) back to form page not working, but full url page can
action=trim(request("action"))
keywords=trim(request("keywords"))

response.write "<script language='javascript'>"
response.write "location.href='search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"';"							
response.write "</script>"	
response.end
%>

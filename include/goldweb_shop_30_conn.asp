<%
' 连接数据库
DB="data/goldweb_shop.asp"
Response.Buffer=True
session.timeout=20
on error resume next
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(server.MapPath("include"))=false Then
  DB="../"&DB
  If fso.FolderExists(server.MapPath("../include"))=False Then
  DB="../"&DB
  End If 
End If
function jincheng (p)
jincheng=p-100000000000000
end function
set fso=Nothing
set conn=server.createobject("adodb.Connection")
connstr="provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DB)
'connstr= "driver={Microsoft Access Driver(*.mdb)};dbq=" & Server.MapPath(DB)
conn.Open connstr
copycolor="#A9A9A9"
%>
<!--#include file="goldweb_shop_info.asp" -->
<!--#include file="goldweb_functions.asp" -->
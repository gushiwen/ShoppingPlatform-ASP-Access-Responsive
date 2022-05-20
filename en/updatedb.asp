<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
response.write "Update DB started here" & "<br>"

'sql="update pay_type set PayTypeDefine='Nets支付', enPayTypeDefine='Nets Payment' where pay_type='52'"
'sql="update pay_type set PayTypeDefine='支票转账', enPayTypeDefine='Check Payment' where pay_type='53'"
'sql="insert into pay_type (PayType, PayTypeClass, PayTypeDefine, enPayTypeDefine, Display) values ('54', 'shop', '微信支付', 'Wechat Payment', True)"
'sql="update goldweb_OrderList set PayType='0'"
'sql="update goldweb_product set AddtoCart='1', ProdIdtext='商品编号'"
response.write "sql: " & sql & "<br>"

If sql<>"" Then 
	'conn.execute (sql)
End If 

conn.close
set conn=Nothing

response.write "Update DB finished here" & "<br>"
%>

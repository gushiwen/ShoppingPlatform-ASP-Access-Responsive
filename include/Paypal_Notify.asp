<!--#include file="goldweb_shop_30_conn.asp" -->
<!--#include file="goldweb_paid.asp" -->
<%
Dim pay_type
set rs_pay_online = conn.execute("select * from pay_type where PayTypeClass='online' and Display=true and instr(enPayTypeDefine,'Paypal')>0")
If Not rs_pay_online.eof	Then 
	pay_type = rs_pay_online("PayType")
End If 
rs_pay_online.close
set rs_pay_online=Nothing

Dim objHttp, tempStr
' post back to PayPal system to validate
tempStr = Request.Form & "&cmd=_notify-validate"
set objHttp = Server.createobject("MSXML2.ServerXMLHTTP") 
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send tempStr

'set fso = createobject("scripting.filesystemobject")
'set file = fso.opentextfile(server.mappath("test.txt"),8,true) '1读取 2写入 8追加
'file.write "pay_type: "& pay_type & vbCrLf
'file.write "objHttp.readystate: "& objHttp.readystate & vbCrLf
'file.write "objHttp.responseText: "& objHttp.responseText & vbCrLf & vbCrLf
'file.close
'set file = nothing
'set fso = nothing

' assign posted variables to local variables
' note: additional IPN variables also available -- see IPN documentation
'Item_name = Request.Form("item_name")
'Receiver_email = Request.Form("receiver_email")
Item_number = Request.Form("item_number")
'Invoice = Request.Form("invoice")
'Payment_status = Request.Form("payment_status")
'Payment_gross = Request.Form("payment_gross")
'Txn_id = Request.Form("txn_id")
'Payer_email = Request.Form("payer_email")

'Check ServerXMLHTTP coplete
if (objHttp.readystate=4 ) then
	' Check notification validation
	if (objHttp.status <> 200 ) then
	' HTTP error handling
	elseif (objHttp.responseText = "VERIFIED") then
		' check that Payment_status=Completed
		' check that Txn_id has not been previously processed
		' check that Receiver_email is an email address in your PayPal account
		' process payment
		'tempMemo = "" '"Paypal payment successful!"&"/nPaypal_invoice: "&Invoice&"/nPayment_gross: "&Payment_gross&"/nTxn_id: "&Txn_id&"/nPayer_email: "&Payer_email&"/n"

		Call UpdateDBAfterOnlinePayment(Item_number,pay_type)
		
	elseif (objHttp.responseText = "INVALID") then
	' log for manual investigation
	else 
	' error
	end If
End If 
set objHttp = Nothing

conn.close
set conn=nothing
%>

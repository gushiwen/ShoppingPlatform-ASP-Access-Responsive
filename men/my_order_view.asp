<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/Alipay_Payto.asp" -->
<%
'call aspsql()
'call checklogin()
OrderNum=request("OrderNum")
MobilePhone=request("MobilePhone")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>View Order Details-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0 ,minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/myorderviewstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

<%
If request.cookies("goldweb")("userid")<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and UserId='"&request.cookies("goldweb")("userid")&"'"
ElseIf request.cookies("goldweb")("userid")="" And MobilePhone<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and RecHomePhone='"&MobilePhone&"'"
Else	
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('You are not authorized to view this order.');"
	response.write "location.href='main.asp';"
	response.write "</script>"
	response.end
End If 

set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sqlinfo,conn,1,1
if rs.eof and rs.bof  then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('We can not find this order.');"
	response.write "location.href='main.asp';"
	response.write "</script>"
	response.end
else
%>
		<div class="order_details">
			<div class="details_title">Order Overview</div> 
			<div class="details_row">
				<%=rs("OrderNum")%>&nbsp;&nbsp;&nbsp;&nbsp;<%=DatePart("yyyy",rs("OrderTime"))%>-<%=Right("0"&DatePart("m",rs("OrderTime")),2)%>-<%=Right("0"&DatePart("d",rs("OrderTime")),2)%>&nbsp;&nbsp;&nbsp;&nbsp;<%=englobalpriceunit%><%=formatnumber(rs("OrderSum"),2)%>
			</div>
			<div class="details_row">
				Order Status: <% response.write conn.execute("select enStatusDefine from order_type where Status='" & rs("Status") & "'")("enStatusDefine") %>
			</div>
			<div class="details_row">
				Payment Status: <% response.write conn.execute("select enPayTypeDefine from pay_type where PayType='" & rs("PayType") & "'")("enPayTypeDefine") %>
			</div>
			<%
				If rs("PayType")="0" And rs("Status")<>"11" And rs("Status")<>"12" And rs("Status")<>"99" Then   
					ListPaymentType()
				End If 
			%>
		</div>


		<!--购物车内产品 -->
		<div class="content-box">
<%
 	Sum = 0
	set rs2=conn.execute("select A.ProdId,A.BuyPrice,A.ProdUnit,B.enProdName,B.enPriceUnit from goldweb_Order A left join goldweb_product B on A.ProdId=B.Prodid where A.OrderNum='"&OrderNum&"' order by A.ID")
	do while not rs2.eof 
		Sum = Sum + FormatNumber(rs2("BuyPrice"),2)*CInt(rs2("ProdUnit"))
%> 
			<ul class="item-content">
				<li class="td td-item">
					<div class="item-pic">
						<a href="product.asp?ProdId=<%=rs2("Prodid")%>"><img src="<%=conn.execute("select ImgPrev from goldweb_product where ProdId='"&rs2("Prodid")&"'")("ImgPrev")%>" width="80" height="80" border="0" onload="DrawImage(this,80,80)" ></a>
					</div>
					<div class="item-info">
						<div class="item-basic-info">
							<a href="product.asp?ProdId=<%=rs2("Prodid")%>" class="item-title J_MakePoint" data-point="tbcart.8.11"><%=rs2("enProdName")%></a>
						</div>
					</div>
				</li>
				<li class="td td-info">
					<span class="sku-line">
						Unit Price <%=rs2("enPriceUnit")%>&nbsp;<%=FormatNumber(rs2("BuyPrice"),2)%>
					</span>
				</li>
				<li class="td td-amount">
					<div class="sl">
						Quantity <%=rs2("ProdUnit")%>x
					</div>
				</li>
			</ul>

<%
	rs2.movenext
	loop
	rs2.close
    set rs2=Nothing
    
	userkou=FormatNumber(rs("thiskou"),2)
    fei=rs("fei")
%> 

			<div class="sub-total">
				<div class="sub-price">
					Sub Total: <%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum,2)%>
				</div>
				<%
					If userkou<>10 then
				%>
				<div class="discount-price">
					After -<%=100-10*userkou%>% Discount: <%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%>
				</div>
				<%
					End If 
				%>
				<div class="total-price">
					Total Amount with Delivery: <%=englobalpriceunit%>&nbsp;<%=FormatNumber((Sum*userkou/10+fei),2)%>
				</div>
			</div>

			<div class="clear"></div>
		</div>

<%
rs.movefirst
%>

		<!--客户详细资料-->
		<div class="customer_details">
			<div class="details_title">Customer Details</div> 
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("pei")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecName")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecHomePhone")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecMail")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecAddress")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecCountry")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecZipcode")%>" readonly="true" >
			</div>
			<% If Trim(rs("Notes"))<>"" Then %>
			<div class="details_row">
				<textarea class="textarea_common" disabled ><%=rs("Notes")%></textarea>
			</div>
			<% End If %>
			<div class="details_row">
				<textarea class="textarea_common" disabled ><%=rs("memo")%></textarea>
			</div>
			<%
				If request.cookies("goldweb")("userid")<>"" Then 
			%>
			<div class="details_row">
				<input class="btn_common" type="button" value="Return to Order List" onclick="document.location.href='my_order.asp';">
			</div>
			<%
				End If 
			%>
		</div>

<%
end If

rs.close
set rs=nothing
%>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
'列出支付方式
Sub ListPaymentType()

INTERFACE_URL="https://www.alipay.com/cooperate/gateway.do?"

' 列出在线支付方式
sql_pay_online = "select * from pay_type where PayTypeClass='online' and Display=true order by PayType desc"
set rs_pay_online=Server.Createobject("ADODB.RecordSet")
rs_pay_online.Open sql_pay_online,conn,1,1
if rs_pay_online.eof and rs_pay_online.bof  then

Else
  do while not rs_pay_online.eof
  if InStr(rs_pay_online("enPayTypeDefine"),"Paypal")>0 then
%>
<div class="details_row">
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="new">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="<%=rs_pay_online("Key1")%>">
<input type="hidden" name="item_name" value="<%=siteurl&" online order: "&rs("OrderNum")%>">
<input type="hidden" name="amount" value="<%=rs("OrderSum")%>">
<input type="hidden" name="item_number" value="<%=rs("OrderNum")%>">
<input type="hidden" name="currency_code" value="<%=rs_pay_online("Key5")%>">
<input type="hidden" name="notify_url" value="<%=FullSiteUrl%>/include/Paypal_Notify.asp">
<input type="hidden" name="return" value="<%=FullSiteUrl%>/en/Paypal_Return.asp">
<input type="hidden" name="cancel_return" value="<%=FullSiteUrl%>/en/Paypal_Cancel_Return.asp">
<input class="btn_pay" type="submit" value="  <%=rs_pay_online("enPayTypeDefine")%> (<%=englobalpriceunit%> <%=FormatNumber(rs("OrderSum"),2)%>)  ">
</form>
</div>
<%
  elseif InStr(rs_pay_online("enPayTypeDefine"),"Alipay")>0 then
  		'海外商家请求参数设置
		Dim Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key
		Dim Alipay_notify_url,Alipay_return_url
		Dim Alipay_subject,Alipay_out_trade_no,Alipay_currency,Alipay_total_fee
		Dim Alipay_Obj,Alipay_itemUrl
		Alipay_gateway = "https://mapi.alipay.com/gateway.do?"
		Alipay_service =	"create_forex_trade"
		Alipay_partner	=	rs_pay_online("Key1")	
		Alipay_input_charset = WebPageCharset
		Alipay_sign_type = "MD5"
		Alipay_key = rs_pay_online("Key2")
		Alipay_notify_url = FullSiteUrl&"/include/Alipay_Notify.asp"
		Alipay_return_url = FullSiteUrl&"/en/Alipay_Return.asp"
		Alipay_subject = siteurl&" online order: "&rs("OrderNum")
		Alipay_out_trade_no = rs("OrderNum")
		Alipay_currency = rs_pay_online("Key5")
		Alipay_total_fee = FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_online("key6"),2),2)

		'海外商家即时到账接口
		Set Alipay_Obj	= New creatAlipayItemURL
		Alipay_itemUrl=Alipay_Obj.creatAlipayItemURL(Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key,Alipay_notify_url,Alipay_return_url,Alipay_subject,Alipay_out_trade_no,Alipay_currency,Alipay_total_fee)

  		'国内商家请求参数设置
		'Dim Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key
		'Dim Alipay_notify_url,Alipay_return_url
		'Dim Alipay_out_trade_no,Alipay_subject,Alipay_payment_type,Alipay_total_fee,Alipay_seller_email
		'Dim Alipay_Obj,Alipay_itemUrl
		'Alipay_gateway = "https://mapi.alipay.com/gateway.do?"
		'Alipay_service =	"create_direct_pay_by_user"
		'Alipay_partner	=	rs_pay_online("Key1")	
		'Alipay_input_charset = WebPageCharset
		'Alipay_sign_type = "MD5"
		'Alipay_key = rs_pay_online("Key2")
		'Alipay_notify_url = FullSiteUrl&"/include/Alipay_Notify.asp"
		'Alipay_return_url = FullSiteUrl&"/ch/Alipay_Return.asp"
		'Alipay_out_trade_no = rs("OrderNum")
		'Alipay_subject = siteurl&" 网上订单："&rs("OrderNum")&" 的付款"
		'Alipay_payment_type = "1"
		'Alipay_total_fee = FormatNumber(rs("OrderSum"),2)
		'Alipay_seller_email = rs_pay_online("Key3") 

		'国内商家即时到账接口
		'Set Alipay_Obj	= New creatAlipayItemURL
		'Alipay_itemUrl=Alipay_Obj.creatAlipayItemURL(Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key,Alipay_notify_url,Alipay_return_url,Alipay_out_trade_no,Alipay_subject,Alipay_payment_type,Alipay_total_fee,Alipay_seller_email)
%>
    <div class="details_row">
    <form action="<%=Alipay_itemUrl%>" method="post" target="new">
      <input class="btn_pay" type="submit" value="  <%=rs_pay_online("enPayTypeDefine")%>(exchange to RenMinBi ￥<%=FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_online("key6"),2),2)%>)  ">
	</form>
	</div>
<%
  end If

  rs_pay_online.movenext
  Loop
end If

rs_pay_online.close
set rs_pay_online=nothing

' 列出银行转账方式
sql_pay_bank = "select * from pay_type where PayTypeClass='bank' and Display=true order by PayType"
set rs_pay_bank=Server.Createobject("ADODB.RecordSet")
rs_pay_bank.Open sql_pay_bank,conn,1,1

if rs_pay_bank.eof and rs_pay_bank.bof  then

Else
  response.write "<div class='details_row'>"
  response.write "<div id='pay_bank_1'><input  class='btn_pay' type='button' value='View Bank / Alipay Accounts for Transferring' onclick='javascript:document.getElementById(""pay_bank_1"").style.display=""none"";document.getElementById(""pay_bank_2"").style.display="""";'></div>"
  response.write "<div id='pay_bank_2' style='display:none;'>"
  response.write "<input  class='btn_pay' type='button' value='Close Bank / Alipay Accounts for Transferring' onclick='javascript:document.getElementById(""pay_bank_1"").style.display="""";document.getElementById(""pay_bank_2"").style.display=""none"";'><br>"
  response.write "<div class='text_bank'>"
  do while not rs_pay_bank.eof
	if InStr(rs_pay_bank("enPayTypeDefine"),"Alipay")>0 then
		response.write rs_pay_bank("enPayTypeDefine")&FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_bank("key6"),2),2)&"<br>Number: "&rs_pay_bank("Key1")&", Name: "&rs_pay_bank("Key2")&"<br><br>"
	Else
		response.write rs_pay_bank("enPayTypeDefine")&"<br>Number: "&rs_pay_bank("Key1")&", Name: "&rs_pay_bank("Key2")&"<br><br>"
	End If 
    rs_pay_bank.movenext
  Loop
  response.write "</div>"
  response.write "<input class='btn_pay' type='button' value='Remittance Report' onclick='javascript:window.location.href=""my_remittance.asp?OrderNum="&OrderNum&"&MobilePhone="&server.URLEncode(MobilePhone)&""";'>"
  response.write "</div>"
  response.write "</div>"
End If 
  
rs_pay_bank.close
set rs_pay_bank=nothing

End Sub
%>
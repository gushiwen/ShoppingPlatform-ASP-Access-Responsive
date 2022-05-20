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
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>View Order Details-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>

</head>
<body>
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; View Order Details
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>
<!--#include file="goldweb_usertree.asp" -->
</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:10px;padding-left:10px;padding-right:10px;" >
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr><td>
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
	response.write "location.href='main.html';"
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
	response.write "location.href='main.html';"
	response.write "</script>"
	response.end
else
%>

<!--订单概述-->
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
  <tr height="30"> 
    <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif">  &nbsp;&nbsp;<b>Order Overview</b>&nbsp;&nbsp;</td>
  </tr>  
  <tr height="24"> 
    <td align="left" width="30%"><b>&nbsp;Order Number </b></td>
    <td align="left" width="70%"><%=rs("OrderNum")%></td>
  </tr>
  <tr height="24"> 
    <td align="left" width="30%"><b>&nbsp;Order Amount </b></td>
    <td align="left" width="70%"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(rs("OrderSum"),2)%></td>
  </tr>
  <tr height="24"> 
    <td align="left"><b>&nbsp;Order Time </b></td>
    <td align="left"><%=rs("OrderTime")%></td>
  </tr>
  <tr height="24"> 
    <td align="left"><b>&nbsp;Order Status </b></td>
    <td align="left">
	  <%
	  response.write conn.execute("select enStatusDefine from order_type where Status='" & rs("Status") & "'")("enStatusDefine") 
	  %>
	</td>
  </tr>
  <tr height="24"> 
    <td align="left"><b>&nbsp;Payment Status </b></td>
    <td align="left">
	  <%
	  response.write "<div style='padding-top: 10px;'>"
	  response.write conn.execute("select enPayTypeDefine from pay_type where PayType='" & rs("PayType") & "'")("enPayTypeDefine") 
	  response.write "</div>"
	  If rs("PayType")="0" And rs("Status")<>"11" And rs("Status")<>"12" And rs("Status")<>"99" Then   
	    ListPaymentType()
      End If 
	  %>
	</td>
  </tr>
  <tr height="38"> 
	<td colspan="2" align="center">
	  &nbsp;
	</td>
  </tr>
</table>


<!--订购产品信息-->
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
<tr height="30"> 
  <td colspan="4" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif">  &nbsp;&nbsp;<b>Order Details</b>&nbsp;&nbsp;</td>
</tr>  
<tr height="30" style="font-weight:bold;"> 
  <td width="10%" align="center">Quantity</td>
  <td width="60%" align="left">Product Name</td>
  <td width="15%" align="center">Unit Price</td>
  <td width="15%" align="center">Extended Price</td>
</tr>
<%
 	Sum = 0
	set rs2=conn.execute("select A.ProdId,A.BuyPrice,A.ProdUnit,B.enProdName,B.enPriceUnit from goldweb_Order A left join goldweb_product B on A.ProdId=B.Prodid where A.OrderNum='"&OrderNum&"' order by A.ID")
	do while not rs2.eof 
		Sum = Sum + FormatNumber(rs2("BuyPrice"),2)*CInt(rs2("ProdUnit"))
%> 
<tr height="25"> 
<td align="center"><%=rs2("ProdUnit")%>x</td>
<td align="left"><a href="product.asp?ProdId=<%=rs2("Prodid")%>"><%=rs2("enProdName")%></a></td>
<td align="center"><%=rs2("enPriceUnit")%>&nbsp;<%=FormatNumber(rs2("BuyPrice"),2)%></td>
<td align="center"><%=rs2("enPriceUnit")%>&nbsp;<%=FormatNumber(FormatNumber(rs2("BuyPrice"),2)*CInt(rs2("ProdUnit")),2)%></td>
</tr>
<%
	rs2.movenext
	loop
	rs2.close
    set rs2=Nothing
    
	userkou=FormatNumber(rs("thiskou"),2)
    fei=rs("fei")
%> 
<tr height="25" valign="middle"> 
<td colspan="4" align="right">&nbsp;</td>
</tr>
<tr height="25" valign="middle"> 
<td colspan="3" align="right"><b>Sub Total:</b>&nbsp;</td>
<td align="center"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum,2)%></td>
</tr>

<%
If userkou<>10 then
%>
<tr height="25" valign="middle"> 
<td colspan="3" align="right"><b>Discount (-<%=100-10*userkou%>%):</b>&nbsp;</td>
<td align="center">- <%=englobalpriceunit%>&nbsp;<%=FormatNumber(FormatNumber(Sum,2)-FormatNumber(Sum*userkou/10,2),2)%></td>
</tr>
<tr height="25" valign="middle"> 
<td colspan="3" align="right"><b>Amount after Discount:</b>&nbsp;</td>
<td align="center"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%></td>
</tr>
<%
End If 
%>
<tr height="25" valign="middle"> 
<td colspan="3" align="right"><b>Total Amount (with Delivery Cost):</b>&nbsp;</td>
<td align="center"><%=englobalpriceunit%>&nbsp;<%=FormatNumber((Sum*userkou/10+fei),2)%></td>
</tr>
<tr height="35" valign="middle"> 
<td colspan="4" align="left">
&nbsp;
</td>
</tr>
</table>

<%
rs.movefirst
%>
<!--客户信息-->
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
  <tr height="30"> 
    <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif">  &nbsp;&nbsp;<b>Customer Details</b>&nbsp;&nbsp;</td>
  </tr>  
  <tr height="24"> 
    <td align="left" width="30%"><b>&nbsp;Delivery Cost </b></td>
    <td align="left" width="70%"><%=rs("pei")%></td>
  </tr>
  <tr height="24"> 
    <td align="left"><b>&nbsp;Expected Delivery Time </b></td>
    <td align="left"><%=rs("gettime")%></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Full Contact Name </b></td>
	<td align="left"><%=rs("RecName")%></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Mobile Phone Number </b></td>
	<td align="left"><%=rs("RecHomePhone")%></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Other Contact Number </b></td>
	<td align="left"><%=rs("RecCompPhone")%></td>
  </tr> 
  <tr height="24"> 
	<td align="left"><b>&nbsp;Email </b></td>
	<td align="left"><%=rs("RecMail")%></td>
  </tr>  
  <tr height="24"> 
	<td align="left"><b>&nbsp;Address </b></td>
	<td align="left"><%=rs("RecAddress")%></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Country </b></td> 
	<td align="left"><%=rs("RecCountry")%></td>
  </tr>  
  <tr height="24"> 
	<td align="left"><b>&nbsp;Post Code </b></td> 
	<td align="left"><%=rs("RecZipcode")%></td>
  </tr>  
  <tr height="120"> 
	<td align="left" ><b>&nbsp;Other Requirements </b></td>
	<td align="left"><textarea cols="40" rows="6" disabled style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;  overflow:auto;"><%=rs("Notes")%></textarea><br /></td>
  </tr>
  <tr height="120"> 
	<td align="left" ><b>&nbsp;System Message </b></td>
	<td align="left"><textarea cols="40" rows="6" disabled style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;  overflow:auto;"><%=rs("memo")%></textarea><br /></td>
  </tr>
  <tr height="38"> 
	<td colspan="2" align="center">
	  &nbsp;
	</td>
  </tr>
<%
  If request.cookies("goldweb")("userid")<>"" Then 
%>
  <tr height="38"> 
	<td colspan="2" align="center">
	  <input type="button" style="font-size:10px; padding:2px;" value="  Return to Order List  " onclick="document.location.href='my_order.asp';">
	</td>
  </tr>
<%
  End If 
%>
</table>

<%
end If

rs.close
set rs=nothing
%>

</td></tr>
</table>
</DIV>
</DIV>
<!--中间右边一起结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
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
<div style='padding-top: 5px;'>
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
<input type="submit" style="font-size:10px; padding:2px;" value="  <%=rs_pay_online("enPayTypeDefine")%> (<%=englobalpriceunit%> <%=FormatNumber(rs("OrderSum"),2)%>)  ">
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
    <div style='padding-top: 5px;'>
    <form action="<%=Alipay_itemUrl%>" method="post" target="new">
      <input type="submit" style="font-size:10px; padding:2px;" value="  <%=rs_pay_online("enPayTypeDefine")%>(exchange to RenMinBi ￥<%=FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_online("key6"),2),2)%>)  ">
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
  response.write "<div style='padding-top: 5px;'>"
  response.write "<div id='pay_bank_1'><input type='button' value='  View Bank Accounts / Alipay Accounts for Transferring  ' style='font-size:10px; padding:2px;' onclick='javascript:document.getElementById(""pay_bank_1"").style.display=""none"";document.getElementById(""pay_bank_2"").style.display="""";'></div>"
  response.write "<div id='pay_bank_2' style='display:none;'>"
  response.write "<input type='button' value='  Close Bank Accounts / Alipay Accounts for Transferring  ' style='font-size:10px; padding:2px;' onclick='javascript:document.getElementById(""pay_bank_1"").style.display="""";document.getElementById(""pay_bank_2"").style.display=""none"";'>&nbsp;&nbsp;<input type='button' value='  Remittance Report  ' style='font-size:10px; padding:2px;' onclick='javascript:window.location.href=""my_remittance.asp?OrderNum="&rs("OrderNum")&"&MobilePhone="&server.URLEncode(MobilePhone)&""";'><br><br>"
  response.write "<div style='padding-left:20px;'>"
  do while not rs_pay_bank.eof
	if InStr(rs_pay_bank("enPayTypeDefine"),"Alipay")>0 then
		response.write rs_pay_bank("enPayTypeDefine")&FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_bank("key6"),2),2)&"<br>Number: "&rs_pay_bank("Key1")&", Name: "&rs_pay_bank("Key2")&"<br><br>"
	Else
		response.write rs_pay_bank("enPayTypeDefine")&"<br>Number: "&rs_pay_bank("Key1")&", Name: "&rs_pay_bank("Key2")&"<br><br>"
	End If 
    rs_pay_bank.movenext
  Loop
  response.write "</div>"
  response.write "</div>"
  response.write "</div>"
End If 
  
rs_pay_bank.close
set rs_pay_bank=nothing

End Sub
%>
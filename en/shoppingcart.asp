<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Shopping Cart-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>

<script language="JavaScript" type="text/JavaScript">
function DoDeleteProduct(ProductId) {
document.shoppingcart.DeleteProduct.value=ProductId;
document.shoppingcart.action="shoppingcart.asp";
document.shoppingcart.submit();
 }

function DoModifyQuantity() {
document.shoppingcart.ModifyQuantity.value="yes";
document.shoppingcart.action="shoppingcart.asp";
document.shoppingcart.submit();
 }

function CheckForm()
{
if (document.shoppingcart.Post.value == 0) {
alert("Please arrange the delivery.");
document.shoppingcart.Post.focus();
return false;
}
if (document.shoppingcart.RecName.value.length <2 || document.shoppingcart.RecName.value.length>30 ) {
alert("Please check contact name.");
document.shoppingcart.RecName.focus();
return false;
}
if (document.shoppingcart.HomePhone.value.length < 6 || document.shoppingcart.HomePhone.value.length > 20) {
alert("Please check mobile phone number.");
document.shoppingcart.HomePhone.focus();
return false;
}
if (document.shoppingcart.RecMail.value.length <11 || document.shoppingcart.RecMail.value.length > 50) {
alert("Please check e-mail address.");
document.shoppingcart.RecMail.focus();
return false;
}
if (document.shoppingcart.RecMail.value.length > 0 && !document.shoppingcart.RecMail.value.match( /^.+@.+$/ ) ) {
 alert("Please check your e-mail address.");
 document.shoppingcart.RecMail.focus();
return false;
}
if (document.shoppingcart.address.value.length < 3 || document.shoppingcart.address.value.length > 150) {
alert("Please check your address.");
document.shoppingcart.address.focus();
return false;
}
if (document.shoppingcart.Country.value.length < 2 || document.shoppingcart.Country.value.length > 50) {
alert("Please check your country.");
document.shoppingcart.address.focus();
return false;
}
if (document.shoppingcart.ZipCode.value.length < 4 || document.shoppingcart.ZipCode.value.length >12) {
alert("Please check the post code.");
document.shoppingcart.ZipCode.focus();
return false;
}
document.shoppingcart.Submit.disabled=true;
return true;
}
</script>
</head>

<body>
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=enSiteName%> 
&gt; Shopping Cart
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr><td>

<%
'call aspsql()
userkou=10
buylist=request.cookies("goldweb")("cart")

If request.cookies("goldweb")("userid")<>"" Then 
set rs=conn.execute ("select * from goldweb_user where userid='"&request.cookies("goldweb")("userid")&"'")
if not (rs.eof and rs.bof) then
HomePhone=rs("HomePhone")
CompPhone=rs("CompPhone")
Country=rs("Country")
address=rs("Address")
Zipcode=rs("ZipCode")
RecMail=rs("UserMail")
RecName=rs("UserName")
UserKou=rs("UserKou")
UserType=rs("UserType")
end If
set rs=Nothing
End If 

' Delete product in shopping cart
if trim(request("DeleteProduct"))<>"" then 
buylist=replace(buylist,trim(request("DeleteProduct")),"XXXXXXXX")
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart")=buylist

Post=trim(request("Post"))
GetTime=trim(request("GetTime"))
RecName=trim(request("RecName"))
HomePhone=trim(request("HomePhone"))
CompPhone=trim(request("CompPhone"))
RecMail=trim(request("RecMail"))
address=trim(request("address"))
Country=trim(request("Country"))
Zipcode=trim(request("Zipcode"))
Notes=trim(request("Notes"))
end If

' Modify product in shopping cart
If Request("ModifyQuantity") <> "" Then
buylist = ""
buyid = Split(Request("ProdId"), ", ")
For I=0 To UBound(buyid)
	If IsNumeric(request(buyid(I))) Then
		if buylist="" then
			buylist = "'" & buyid(I) & "', '"&CInt(request(buyid(I)))&"'"
		else
			buylist = buylist & ", '" & buyid(I) & "', '"&CInt(request(buyid(I)))&"'"
		End if
	end if
Next
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart") = buylist

Post=trim(request("Post"))
GetTime=trim(request("GetTime"))
RecName=trim(request("RecName"))
HomePhone=trim(request("HomePhone"))
CompPhone=trim(request("CompPhone"))
RecMail=trim(request("RecMail"))
address=trim(request("address"))
Country=trim(request("Country"))
Zipcode=trim(request("Zipcode"))
Notes=trim(request("Notes"))
End If

if buylist="" then
	response.Cookies("goldweb").path="/"
	response.cookies("goldweb")("cart") = ""
	response.write "Your shopping cart is empty."
else
	' buylist '02337', '1', '02335', '1', '01610', '1' 
	buyprodlist=""
	sp_buylist=split(buylist,", ")
	for i=0 to ubound(sp_buylist) Step 2 
		If buyprodlist="" Then 
			buyprodlist = Trim(sp_buylist(i))
		Else
			buyprodlist = buyprodlist & ", " & Trim(sp_buylist(i))
		End If 
	Next

	' buyprodlist '02337', '02335', '01610' 
	Set rs=conn.execute("select * from goldweb_product where ProdId in ("&buyprodlist&") order by instr ('"&Replace(buyprodlist,"'","")&"',ProdId)")
	if rs.eof and rs.bof then
		response.Cookies("goldweb").path="/"
		response.cookies("goldweb")("cart") = ""
		response.write "Your shopping cart is empty."
		rs.close
		set rs=nothing
	else
%>

<form action="ordersubmit.asp" method="Post" name="shoppingcart" onSubmit="return CheckForm();">

<!--订单信息-->
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
<tr height="30" bgcolor="#F4F6FC" style="font-weight:bold;"> 
<td width="5%" align="center">Delete</td>
<td width="10%" align="center">Quantity</td>
<td width="55%" align="left">Product Name</td>
<td width="15%" align="center">Unit Price</td>
<td width="15%" align="center">Extended Price</td>
</tr>
<%
		Sum = 0
		for i=0 to ubound(sp_buylist) Step 2 
			if Trim(Replace(sp_buylist(i),"'",""))=rs("prodid") Then
				Quatity=Trim(Replace(sp_buylist(i+1),"'",""))

				 if not isNumeric(Quatity) then Quatity=1
				Quatity=CInt(Quatity)
				If Quatity <= 0 Then Quatity = 1
				Sum = Sum + FormatNumber(rs("enPriceList"),2)*Quatity
	  
				'判断ProdId="00001", Post=4 送货安排N.A., 数量输入框readonly="true"
				If rs("ProdId")="00001" Then
					Post=4
				End If 
%> 
<tr height="25"> 
<td align="center"> 
<input type="hidden" name="ProdId" value="<%=rs("ProdId")%>">
<a href="javascript:DoDeleteProduct('<%=rs("ProdId")%>');"><img border=0 src="../images/small/delete.gif"></a>
</td>
<td align="center"> 
<input type="Text" name="<%=rs("ProdId")%>" value="<%=Quatity%>" size="3" maxlength="4" class="form" onchange="javascript:DoModifyQuantity();" <%If rs("ProdId")="00001" Or rs("ProdId")="00003" Then response.write "readonly=""true"""%>>
</td>
<td align="left"><A href="product.asp?ProdId=<%=rs("ProdId")%>"><%=rs("enProdName")%></A></td>
<td align="center"><%=rs("enPriceUnit")%>&nbsp;<%=FormatNumber(rs("enPriceList"),2)%></td>
<td align="center"><%=rs("enPriceUnit")%>&nbsp;<%=FormatNumber(rs("enPriceList")*Quatity,2)%></td>
</tr>
<%
      			rs.MoveNext
				If rs.eof Then Exit For 
			end If
		next

		rs.close
		set rs=nothing
%> 
<tr height="25" valign="middle"> 
<td colspan="5" align="right">&nbsp;</td>
</tr>
<tr height="25" valign="middle"> 
<td colspan="3" align="left">
<%
If request.cookies("goldweb")("userid")="" Then 
	response.write "(<a href=""reg_member.asp""><b>Register</b></a> and <a href=""login.asp""><b>Login in</b></a> to Enjoy -5% Discount)"
Else 
	If UserType="1" Then response.write "&nbsp;&nbsp;<a href=""productorder.asp?Prodid=00001""><b>(Upgrade to Permanent VIP Member to Enjoy -10% Discount)</b></a>"
End If 
%>
</td>
<td colspan="1" align="right"><b>Sub Total:</b>&nbsp;</td>
<td align="center"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum,2)%></td>
</tr>

<%
If userkou<>10 then
%>
<tr height="25" valign="middle"> 
<td colspan="4" align="right"><b>Discount (-<%=100-10*userkou%>%):</b>&nbsp;</td>
<td align="center">- <%=englobalpriceunit%>&nbsp;<%=FormatNumber(FormatNumber(Sum,2)-FormatNumber(Sum*userkou/10,2),2)%></td>
</tr>
<tr height="25" valign="middle"> 
<td colspan="4" align="right"><b>Amount after Discount:</b>&nbsp;</td>
<td align="center"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%></td>
</tr>
<%
End If 
%>
<tr height="25" valign="middle" id="Tr_Amount" style="display:none;"> 
<td colspan="4" align="right"><b>Total Amount (with Delivery Cost):</b>&nbsp;</td>
<td align="center" id="Td_Amount"></td>
</tr>
<tr height="48" valign="middle"> 
<td colspan="5" align="center">
&nbsp;<input type="hidden" name="DeleteProduct" value="">
&nbsp;<input type="hidden" name="ModifyQuantity" value="">
&nbsp;<input type="hidden" name="Total" value="">
&nbsp;<input type="hidden" name="Pei" value="">
&nbsp;<input type="hidden" name="Fei" value="">
</td>
</tr>
</table>

<!--客户信息-->
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
  <tr height="30"> 
    <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif">  &nbsp;&nbsp;<b>Customer Details</b>&nbsp;&nbsp;<font color="red">* Required Fields</font></td>
  </tr>  
<%
if enmianyoufei<=Sum*userkou/10 then enfei1=0
enpei1=enpei1 & " cost " & englobalpriceunit & " " & enfei1 &""
if enpei2<>"" then enpei2=enpei2 & " cost " & englobalpriceunit & " " & enfei2 &""
if enpei3<>"" then enpei3=enpei3 & " cost " & englobalpriceunit & " " & enfei3 &""
if enpei4<>"" then enpei4=enpei4 & " cost " & englobalpriceunit & " " & enfei4 &""
if enpei5<>"" then enpei5=enpei5 & " cost " & englobalpriceunit & " " & enfei5 &""
if enpei6<>"" then enpei6=enpei6 & " cost " & englobalpriceunit & " " & enfei6 &""

response.write "<tr height='24'><td align='left' width='30%'><b>&nbsp;Arrange Delivery </b><br>&nbsp;"&enmianyoufei_msg&"</td><td align='left' width='70%'>"
response.write "<select size='1' name='Post' style='BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ' onchange='javascript:UpdateTotalAmount();'>"
response.write "<option selected value='0'>Please arrange delivery here </option>"
If Post=1 Then
  response.write "<option value='1' selected>"&enpei1&"</option>"
Else
  response.write "<option value='1'>"&enpei1&"</option>"
End If 
if enpei2<>"" then 
If Post=2 Then
  response.write "<option value='2' selected>"&enpei2&"</option>"
Else
  response.write "<option value='2'>"&enpei2&"</option>"
End If 
End If 
if enpei3<>"" then 
If Post=3 Then
  response.write "<option value='3' selected>"&enpei3&"</option>"
Else
  response.write "<option value='3'>"&enpei3&"</option>"
End If 
End If 
if enpei4<>"" then 
If Post=4 Then
  response.write "<option value='4' selected>"&enpei4&"</option>"
Else
  response.write "<option value='4'>"&enpei4&"</option>"
End If 
End If 
if enpei5<>"" then 
If Post=5 Then
  response.write "<option value='5' selected>"&enpei5&"</option>"
Else
  response.write "<option value='5'>"&enpei5&"</option>"
End If 
End If 
if enpei6<>"" then 
If Post=6 Then
  response.write "<option value='6' selected>"&enpei6&"</option>"
Else
  response.write "<option value='6'>"&enpei6&"</option>"
End If 
End If 
response.write "</select> <font color='red'>*</font></td></tr>"
%>

<script language="JavaScript" type="text/JavaScript">
// 函数定义在免邮费比较后面
function UpdateTotalAmount(){
if (document.shoppingcart.Post.value!=0){
  if (document.shoppingcart.Post.value==1)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei1,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei1,2)%>";
    document.shoppingcart.Pei.value="<%=enpei1%>";
    document.shoppingcart.Fei.value="<%=enfei1%>";
  }
  if (document.shoppingcart.Post.value==2)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei2,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei2,2)%>";
    document.shoppingcart.Pei.value="<%=enpei2%>";
    document.shoppingcart.Fei.value="<%=enfei2%>";
  }
  if (document.shoppingcart.Post.value==3)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei3,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei3,2)%>";
    document.shoppingcart.Pei.value="<%=enpei3%>";
    document.shoppingcart.Fei.value="<%=enfei3%>";
  }
  if (document.shoppingcart.Post.value==4)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei4,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei4,2)%>";
    document.shoppingcart.Pei.value="<%=enpei4%>";
    document.shoppingcart.Fei.value="<%=enfei4%>";
  }
  if (document.shoppingcart.Post.value==5)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei5,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei5,2)%>";
    document.shoppingcart.Pei.value="<%=enpei5%>";
    document.shoppingcart.Fei.value="<%=enfei5%>";
  }
  if (document.shoppingcart.Post.value==6)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei6,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei6,2)%>";
    document.shoppingcart.Pei.value="<%=enpei6%>";
    document.shoppingcart.Fei.value="<%=enfei6%>";
  }
  document.getElementById("Tr_Amount").style.display="";
} else {
  document.getElementById("Td_Amount").innerHTML="";
  document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10,2)%>";
  document.shoppingcart.Pei.value="";
  document.shoppingcart.Fei.value="";
  document.getElementById("Tr_Amount").style.display="none";
}
}
// 修改或删除购物车产品后，页面刷新时需要调用
UpdateTotalAmount(); 
</script>

  <tr height="24"> 
    <td align="left"><b>&nbsp;Expected Delivery Time </b></td>
    <td align="left"><input name="GetTime" size="30"  maxlength="50" value="<%=GetTime%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Full Contact Name </b></td>
	<td align="left"><input type="text" name="RecName" maxlength="30" size="30" value="<%=RecName%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Mobile Phone Number </b></td>
	<td align="left"><input type="text" name="HomePhone" size="30" maxlength="20" value="<%=HomePhone%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Other Contact Number </b></td>
	<td align="left"><input type="text" name="CompPhone"  size="30" maxlength="20" value="<%=CompPhone%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> </td>
  </tr> 
  <tr height="24"> 
	<td align="left"><b>&nbsp;Email </b></td>
	<td align="left"><input type="text" name="RecMail"  size="30" maxlength="50" value="<%=RecMail%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>  
  <tr height="24"> 
	<td align="left"><b>&nbsp;Address </b></td>
	<td align="left"><input name="address" maxlength="150" size="30" value="<%=Address%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>
  <tr height="24"> 
	<td align="left"><b>&nbsp;Country </b></td> 
	<td align="left"><input type="text" name="Country" size="30" maxlength="50" value="<%=Country%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>  
  <tr height="24"> 
	<td align="left"><b>&nbsp;Post Code </b></td> 
	<td align="left"><input type="text" name="ZipCode" size="30" maxlength="12" value="<%=Zipcode%>" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <font color="red">*</font></td>
  </tr>  
  <tr height="120"> 
	<td align="left" ><b>&nbsp;Other Requirements </b><br>&nbsp;Color, Size, Packing<br>&nbsp;Address for piano moving service</td>
	<td align="left"><textarea name="Notes" cols="24" rows="6" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;  overflow:auto;"></textarea><br /></td>
  </tr>
  <tr height="38"> 
	<td colspan="2" align="center">
	  <input type="submit" value="  Submit Order  " name="Submit" style="font-size: 10px; padding:2px;">&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
	   <input type="button" value="  Reset Form  " name="Reset" style="font-size:10px; padding:2px;" onclick="javascript:location.href='shoppingcart.asp';">&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
	</td>
  </tr>
</table>

</form>
<%
  end if
end if
%>

</td></tr>
</table>
</DIV>
</DIV>
<!--右边结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>
<%
conn.close
set conn=nothing
%>
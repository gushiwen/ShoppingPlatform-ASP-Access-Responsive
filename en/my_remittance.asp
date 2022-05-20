<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
OrderNum=request("OrderNum")
MobilePhone=request("MobilePhone")

if request("send")="ok" then 
	call send()
else

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
Else
	OrderSum=formatnumber(rs("OrderSum"),2)
	UserName=rs("RecName")
	HomePhone=rs("RecHomePhone")
	UserMail=rs("RecMail")
	Address=rs("RecAddress")
	Country=rs("RecCountry")
	ZipCode=rs("RecZipcode")
End If 

rs.close
set rs=Nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Remittance Report-<%=ensitename%>-<%=siteurl%></title>
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
&gt; Remittance Report
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>
<!--#include file="goldweb_usertree.asp" -->
</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >
	
<table width="90%" border="0" cellpadding="5" cellspacing="5" align="center">
<form name="add" method="post" action="my_remittance.asp">
<tr height="24">
  <td colspan="2">Please report to us after you make the payment.</td>
</tr>

<tr height="24">
  <td width="28%" align="right">Order No.: &nbsp;</td>
  <td width="72%" align="left"><input name="hk1" type="text" size="25" maxlength="14" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=OrderNum%>"></td>
</tr>
<tr height="24">
  <td align="right">Payment Amount: &nbsp;</td>
  <td align="left"><input name="hk2" type="text" size="25" maxlength="10" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=englobalpriceunit%> <%=OrderSum%>"></td>
</tr>
<tr height="24">
  <td align="right">From Bank Name: &nbsp;</td>
  <td align="left"><input name="hk3" type="text" size="25" maxlength="25" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana"> </td>
</tr>
<tr height="26">
  <td align="right">Paid by: &nbsp;</td>
  <td align="left">
	<%
	 ' 列出银行转账方式
	set rs_pay_bank=conn.execute("select * from pay_type where PayTypeClass='bank' and Display=true order by PayType")
	if rs_pay_bank.eof and rs_pay_bank.bof  then
	Else
	  do while not rs_pay_bank.eof
	    response.write "<input name=""hk4"" type=""radio"" value="""&rs_pay_bank("enPayTypeDefine")&"""> "&rs_pay_bank("enPayTypeDefine")&" <br>"
	    rs_pay_bank.movenext
	  Loop
	End If 
	rs_pay_bank.close
	set rs_pay_bank=nothing
	%>
  </td>
</tr>

<tr height="24">
  <td align="right">Full Contact Name: &nbsp;</td>
  <td align="left"><input name="hk5" type="text" size="25" maxlength="25" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=UserName%>"></td>
</tr>
<tr height="24">
  <td align="right">Mobile Number: &nbsp;</td>
  <td align="left"><input name="hk6" type="text" size="25" maxlength="25" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=HomePhone%>"></td>
</tr>
<tr height="24">
  <td align="right">Email: &nbsp;</td>
  <td align="left"><input name="hk7" type="text" size="25" maxlength="100" style="BORDER:#CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=UserMail%>"></td>
</tr>
<tr height="24">
  <td align="right">Address: &nbsp;</td>
  <td align="left"><input name="hk8" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=Address%>"></td>
</tr>
<tr height="24">
  <td align="right">Country: &nbsp;</td>
  <td align="left"><input name="hk9" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=Country%>"></td>
</tr>
<tr height="24">
  <td align="right">ZipCode: &nbsp;</td>
  <td align="left"><input name="hk10" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; FONT-FAMILY: verdana" value="<%=ZipCode%>"></td>
</tr>
<tr  height="80">
  <td align="right">Note: &nbsp;<br><font color=red>≤200 Characters &nbsp;</font></td>
  <td><textarea rows="4" name="hk11" cols="40" style="BORDER: #CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666;  overflow:auto;"></textarea></td>
</tr>
<tr height="58">
  <td colspan="2" align="center">
    <input type="submit" name="Submit" value="  Submit  " style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
    <input type="reset" name="Reset" value="  Reset  " style="font-size:10px; padding:2px;">
	<input name="send" type="hidden" value="ok">
	<input name="OrderNum" type="hidden" value="<%=OrderNum%>">
	<input name="MobilePhone" type="hidden" value="<%=MobilePhone%>">
  </td>
</tr>
</form>
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
End If 

conn.close
set conn=nothing

sub send()
	call goldweb_check_path()

	if request.form("hk1")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Order Number.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
	
	if request.form("hk2")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Payment Amount.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk3")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in From Bank Name.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk4")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Payment Method.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if llen(request.form("hk11"))>200  then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('The note is too long.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	dim MailTitle
	dim MailContent
	MailTitle = request.form("hk5") & " has reported remittance on " & siteurl
	MailContent="Order No.: " & request.form("hk1")  & "<br>"
	MailContent=MailContent & "Payment Amount: " & request.form("hk2")  & "<br>"
	MailContent=MailContent & "From Bank Name: " & request.form("hk3")  & "<br>"
	MailContent=MailContent & "Paid by: " & request.form("hk4")  & "<br>"
	MailContent=MailContent & "Full Contact Name: " & request.form("hk5")  & "<br>"
	MailContent=MailContent & "Mobile Number: " & request.form("hk6")  & "<br>"
	MailContent=MailContent & "Email: " & request.form("hk7")  & "<br>"
	MailContent=MailContent & "Address: " & request.form("hk8")  & "<br>"
	MailContent=MailContent & "Country: " & request.form("hk9")  & "<br>"
	MailContent=MailContent & "ZipCode: " & request.form("hk10")  & "<br>"
	MailContent=MailContent & "Note: " & request.form("hk11")  & "<br>"
	Call SendOutMail(MailTitle, MailContent, adm_mail, adm_mail, request.form("hk7"))

	sql="select * from hand"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
  	rs("name")="admin"
  	rs("neirong")=MailContent
  	rs("riqi")=DateAdd("h", TimeOffset, now())
	If request.cookies("goldweb")("userid")<>"" Then 
 		rs("fname")=request.cookies("goldweb")("userid")
	Else
 		rs("fname")="Guest"
	End If 
  	rs("zuti")="1"
	rs.update	
	rs.close
	set rs=nothing

	conn.close
	set conn=Nothing

	response.write "<script language='javascript'>"
	response.write "alert('The message has been sent sucessfully.');"
	response.write "</script>"
	response.write "<meta http-equiv=refresh content='0;URL=my_order_view.asp?OrderNum="&OrderNum&"&MobilePhone="&server.URLEncode(MobilePhone)&"'>"
end sub
%>
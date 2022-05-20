<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()
if request("del")<>"" then call delorder()
if request("cancel")<>"" then call cancelorder()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Order List-<%=ensitename%>-<%=siteurl%></title>
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
&gt; Order List
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
<tr bgcolor="#F4F6FC" height="50">
  <td align="center" width="15%"><b>Order No.</b></td>
  <td align="center" width="20%"><b>Date</b></td>
  <td align="center" width="12%"><b>Amount</b></td>
  <td align="left" width="40%">&nbsp;&nbsp;<b>Order Status / Payment Status</b></td>
  <td align="center" width="13%"><b>Operation</b></td>
</tr>
<%
sql = "select A.OrderNum,A.OrderTime,A.RecName,A.OrderSum,A.RecName,A.Status,A.PayType,B.enStatusDefine,C.enPayTypeDefine "&_
          " from (goldweb_OrderList A left join order_type B on A.Status = B.Status) left join pay_type C on A.PayType = C.PayType "&_
		  " where A.del<>true and A.UserId='"&request.cookies("goldweb")("userid")&"' order by A.OrderTime desc"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,1

if rs.eof and rs.bof  then
	response.write "<tr><td colspan=5 height=50 align=center>No orders</td></tr>"
	else
	do while not rs.eof
%>
<tr height="50">
  <td align="center"><a href="my_order_view.asp?OrderNum=<%=rs("OrderNum")%>"><%=rs("OrderNum")%></a></td>
  <td align="center"><%=rs("OrderTime")%></td>
  <td align="center"><%=englobalpriceunit%>&nbsp;<%=formatnumber(rs("OrderSum"),2)%></td>
  <td align="left">&nbsp;&nbsp;<%=rs("enStatusDefine")%><br>&nbsp;&nbsp;<%=rs("enPayTypeDefine")%></td>
  <td align="center">
    <%
	  response.write "<a href=""my_order_view.asp?OrderNum="&rs("OrderNum")&""">View</a>"
	  If rs("Status")="0" And rs("PayType")="0" Then 
	    response.write "&nbsp;&nbsp;<a href=""my_order.asp?cancel="&rs("OrderNum")&""">Cancel</a>"
      ElseIf rs("Status")="11" And rs("PayType")="0" then 
	    response.write "&nbsp;&nbsp;<a href=""my_order.asp?cancel="&rs("OrderNum")&""">Retrieve</a>"
	  End If 
	%>
  </td>
</tr>
<%
rs.movenext
loop
end If

rs.close
set rs=nothing
%>
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
sub cancelorder()
sql="select * from goldweb_OrderList where OrderNum='"&request("cancel")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
		tishi="The order no. is not exist."
	elseif rs("userid")<>request.cookies("goldweb")("userid") then
		tishi="You are not authorized to manage this order."
	Else	
		'Add  order status updation to system message
	    set rs2 = conn.execute("select Status, enStatusDefine from order_type Order by Status asc")
		Do While Not rs2.eof	
		if rs2("Status") = "11" then enStatusDefine1=rs2("enStatusDefine")
		if rs2("Status") = "0"  then enStatusDefine2=rs2("enStatusDefine")
		rs2.movenext
		if rs2.eof then	exit do
		Loop
		rs2.close
	   set rs2=nothing

		if rs("Status")="11" then
		   rs("Status")="0"
		   rs("Memo")  = rs("Memo") & vbCrLf & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine1 & """ to """ & enStatusDefine2 & """;"
		   rs.update
		   tishi="The order has been retrieved sucessfully!"
		elseif rs("Status")="0" then
		   rs("Status")="11"
		   rs("Memo")  = rs("Memo") & vbCrLf & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine2 & """ to """ & enStatusDefine1 & """;"
		   tishi="The order has been cancelled successfully!"
		   rs.update
		else
	  	  tishi="The order can not be auto-cancelled!"
		end if
	end if
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('"&tishi&"');"
response.write "location.href='my_order.asp';"
response.write "</script>"
end sub

sub delorder()
sql="select * from goldweb_OrderList where OrderNum='"&request("Del")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3

	if rs.eof and rs.bof then
		tishi="The order no. is not exist."
	elseif rs("userid")<>request.cookies("goldweb")("userid") then
		tishi="You are not authorized to manage this order."
	else
		conn.execute("update goldweb_OrderList set del=true where OrderNum='"&request("Del")&"'")
		tishi="The order has been deleted sucessfully!"
	end if
  		response.write "<script language='javascript'>"
		response.write "alert('"&tishi&"');"
		response.write "location.href='my_order.asp';"
		response.write "</script>"
rs.close
set rs=nothing
end Sub

conn.close
set conn=nothing
%>

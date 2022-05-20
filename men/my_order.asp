<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 订单列表
call checklogin()
if request("del")<>"" then call delorder()
if request("cancel")<>"" then call cancelorder()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Order List-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/myorderstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="list-box">
			<div class="list_title">Order List</div> 

<%
sql = "select A.OrderNum,A.OrderTime,A.RecName,A.OrderSum,A.RecName,A.Status,A.PayType,B.enStatusDefine,C.enPayTypeDefine "&_
          " from (goldweb_OrderList A left join order_type B on A.Status = B.Status) left join pay_type C on A.PayType = C.PayType "&_
		  " where A.del<>true and A.UserId='"&request.cookies("goldweb")("userid")&"' order by A.OrderTime desc"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,1

if rs.eof and rs.bof  then
	response.write "<div class='list_row'><div class='no-item'>No order</div></div>"
else
	do while not rs.eof
%>

			<div class="list_row">
				<a class="btn_view" href="my_order_view.asp?OrderNum=<%=rs("OrderNum")%>"><%=rs("OrderNum")%>&nbsp;&nbsp;&nbsp;&nbsp;<%=DatePart("yyyy",rs("OrderTime"))%>-<%=Right("0"&DatePart("m",rs("OrderTime")),2)%>-<%=Right("0"&DatePart("d",rs("OrderTime")),2)%>&nbsp;&nbsp;&nbsp;&nbsp;<%=englobalpriceunit%><%=formatnumber(rs("OrderSum"),2)%></a>
				<%
	  				If rs("Status")="0" And rs("PayType")="0" Then 
	    				response.write "<a class=""btn_opt"" href=""my_order.asp?cancel="&rs("OrderNum")&""">Cancel</a>"
      				ElseIf rs("Status")="11" And rs("PayType")="0" then 
	    				response.write "<a class=""btn_opt"" href=""my_order.asp?cancel="&rs("OrderNum")&""">Retrieve</a>"
	  				End If 
				%>
			</div>

<%
rs.movenext
loop
end If

rs.close
set rs=nothing
%>

		</div>

		<!--#include file="goldweb_down.asp"-->
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

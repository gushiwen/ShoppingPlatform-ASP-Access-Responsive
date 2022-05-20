<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 账户管理
call checklogin()

set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
userid=rs("UserID")
totallogin=rs("totallogin")
usertype=rs("usertype")
userkou=rs("userkou")
totalsum=FormatNumber(rs("totalsum"),2)
jifen=CInt(rs("jifen"))

if usertype="1" then
usertypetext=enusertype1 & "<br><br><a class=""btn_VIP"" href=""productorder.asp?Prodid=00001"">Upgrade to VIP Member -10% Discount</a>"
elseif usertype="2" then
usertypetext=enusertype2
elseif usertype="3" then
usertypetext=enusertype3
elseif usertype="4" then
usertypetext=enusertype4
elseif usertype="5" then
usertypetext=enusertype5
elseif usertype="6" then
usertypetext=enusertype6
else
usertypetext=""
end if

IP=Request.serverVariables("REMOTE_ADDR")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Account Management-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/usercenterstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
			<div class="details_title">Account Management</div> 
			<div class="details_row">
				<div class="text">
					<% response.write "Account ID: " & UserID & ", Discount -" & (10-userkou)*10 & "%" %>
				</div>
			</div>
			<div class="details_row">
				<div class="text">
					<% response.write "Account Type: " & usertypetext %>
				</div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Order Management" onclick="document.location.href='my_order.asp';">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="My Advertisements" onclick="document.location.href='my_adv.asp';">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Favorite Products" onclick="document.location.href='my_fav.asp';">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Personal Information" onclick="document.location.href='my_info.asp';">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Password Setting" onclick="document.location.href='my_psw.asp';">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Log out" onclick="document.location.href='logout.asp';">
			</div>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
call checklogin()
'call aspsql()

set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
userid=rs("UserID")
totallogin=rs("totallogin")
usertype=rs("usertype")
userkou=rs("userkou")
totalsum=FormatNumber(rs("totalsum"),2)
jifen=CInt(rs("jifen"))

if usertype="1" then
usertypetext=enusertype1 & "&nbsp;&nbsp;<a href=""productorder.asp?Prodid=00001""><b>(Upgrade to Permanent VIP Member to Enjoy -10% Discount)</b></a>"
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
'response.Cookies("goldweb").path="/"
'response.cookies("goldweb")("userkou")=userkou
IP=Request.serverVariables("REMOTE_ADDR")
'if isnull(rs("UserQuestion"))=true or rs("UserQuestion")="" then
'response.write "<script language='javascript'>"
'response.write "alert('For account safty, please do Retrieve Password Setting first.\n\nYou can retrieve your password once you forgot your password.');"
'response.write "location.href='my_psw_set.asp';"
'response.write "</script>"
'response.end
'end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Account Management-<%=ensitename%>-<%=siteurl%></title>
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
&gt; Account Management
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

<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
<%
response.write "<tr height=""24""><td>Hello <font color=red> "&UserID&"</font>, welcome to Account Management Page.</td></tr>"
response.write "<tr height=""24""><td>You've logged in for <font color=red>"&totallogin&"</font> time(s).</td></tr>"
response.write "<tr height=""24""><td>Your logging IP is <font color=red>"&IP&"</font></td></tr>"
response.write "<tr height=""24""><td>Your acccount type is <font color=red>"&usertypetext&"</font></td></tr>"
response.write "<tr height=""24""><td>You have spent <font color=red>S$"&FormatNumber(totalsum,2)&" </font> totally for online orders.</td></tr>"
if userkou<>10 then response.write "<tr height=""24""><td>You can enjoy <font color=red>-"&(10-userkou)*10&"% </font> discount when shopping with us.</td></tr>"
'response.write "<tr height=""24""><td>Your reward point is <font color=red>"&jifen&"</font>.</td></tr>"
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
conn.close
set conn=nothing
%>

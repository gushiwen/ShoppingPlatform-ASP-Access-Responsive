<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
' 修改密码
call checklogin()
if request("edit")="ok" then call edit()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Password Setting-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/loginstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

		<div class="account_details">
		<form name="myinfo" action="my_psw.asp" method="post">
			<div class="details_title">Password Setting</div> 
			<input type="hidden" name="userid" value="<%=rs("UserID")%>">
			<div class="details_row">
				<input class="input_common" type="password" name="oldpassword" maxlength="16" placeholder="Old Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw1" maxlength="16" placeholder="New Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw2" maxlength="16" placeholder="Repeat Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" name="Submit" value="Submit Form" >
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Retrieve password setting" onclick="document.location.href='my_psw_set.asp';" >
			</div>
			<div class="details_row">
				<input class="btn_common" type="reset" name="Reset" value="Reset Form" >
			</div>
			<input type="hidden" name="edit" value="ok">
        </form>
		</div>

<%
rs.close
set rs=nothing
%>	  

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
sub edit()
call goldweb_check_path()
oldpassword=trim(request("oldpassword"))
Pw1=trim(request("pw1"))
Pw2=trim(request("pw2"))

if oldpassword="" or Pw1="" or pw2=""  then 
response.write "<script language='javascript'>"
response.write "alert('Please fill in the password.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if Pw1<>pw2 then 
response.write "<script language='javascript'>"
response.write "alert('You have input two different passwords.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if llen(pw1)<6 then 
response.write "<script language='javascript'>"
response.write "alert('The password should be at least 6 characters.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
if rs("userpassword")<>md5(oldpassword) then 
response.write "<script language='javascript'>"
response.write "alert('The old password is incorrect.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if ucase(request.cookies("goldweb")("userid"))<>ucase(request.form("userid")) then
response.write "<script language='javascript'>"
response.write "alert('You are not authorized to do this.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("userpassword")=md5(pw1)
rs.update
rs.close
set rs=nothing

response.write "<script language='javascript'>"
response.write "alert('Your password has been changed.');"
response.write "location.href='logout.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>

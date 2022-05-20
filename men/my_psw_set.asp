<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
' 取回密码设置
call checklogin()
if request("set")="ok" then call setq()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Retrieve Password Setting-<%=ensitename%>-<%=siteurl%></title>
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
        <form name="myinfo" action="my_psw_set.asp" method="post">
			<div class="details_title">Retrieve Password Setting</div> 
			<input type="hidden" name="userid" value="<%=rs("UserID")%>">
			<div class="details_row">
				<input class="input_common" type="password" name="password" maxlength="16" placeholder="Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserQuestion" value="<%=rs("UserQuestion")%>" maxlength="50" placeholder="Question to retrieve password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserAnswer1" value="" maxlength="50" placeholder="Answer to the question" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserAnswer2" value="" maxlength="50" placeholder="Repeat the answer" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" name="Submit" value="Submit Form" >
			</div>
			<div class="details_row">
				<input class="btn_common" type="reset" name="Reset" value="Reset Form" >
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="javascript: location.href='javascript:history.go(-1)';">
			</div>
			<input type="hidden" name="set" value="ok">
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
sub setq()
call goldweb_check_path()
Password=trim(request("Password"))
UserQuestion=trim(request("UserQuestion"))
UserAnswer1=trim(request("UserAnswer1"))
UserAnswer2=trim(request("UserAnswer2"))

if Password="" then 
response.write "<script language='javascript'>"
response.write "alert('Please fill in the password.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if UserQuestion="" or UserAnswer1="" or UserAnswer2=""  then 
response.write "<script language='javascript'>"
response.write "alert('Please fill in the question and answer.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if UserAnswer1<>UserAnswer2 then 
response.write "<script language='javascript'>"
response.write "alert('You have entered two different answers.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if len(UserAnswer1)<3 then 
response.write "<script language='javascript'>"
response.write "alert('The answer should be at least 3 characters.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
if rs("userpassword")<>md5(Password) then 
response.write "<script language='javascript'>"
response.write "alert('The password is incorrect.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if ucase(request.cookies("goldweb")("userid"))<>ucase(request("userid")) then
response.write "<script language='javascript'>"
response.write "alert('You are not authorized to do this.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("UserQuestion")=UserQuestion
rs("UserAnswer")=md5(UserAnswer1)
rs.update
rs.close
set rs=nothing

response.write "<script language='javascript'>"
response.write "alert('Retrieve Password Setting is done.');"
response.write "location.href='my_psw.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>

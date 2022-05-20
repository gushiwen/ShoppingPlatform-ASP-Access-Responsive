<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
'call aspsql()
call goldweb_check_path()

if request.cookies("goldweb")("userid")<>"" then
	conn.close
	set conn=Nothing
	response.write "<meta http-equiv='refresh' content='0;URL=user_center.asp'>"
end if

set rs=conn.execute("select * from goldweb_user where UserId='"&request.form("userid")&"'")
if rs.eof and rs.bof then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Invalid Account ID.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if md5(request.form("UserAnswer"))<>rs("UserAnswer") then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	session("gpw_error")=session("gpw_error")+1
	response.write "<script language='javascript'>"
	response.write "alert('Wrong answer for "&session("gpw_error")&" times.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if request.form("save")="ok" then

  if request.form("pw1")<> request.form("pw2") then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('You have input two different passwords.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
  else
	conn.execute ("update goldweb_user set UserPassword='"&md5(request.form("pw1"))&"' where userid='"&request.form("userid")&"'")
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Congratulations! You have set your password sucessfully!');"
	response.write "location.href='login.asp';"
	response.write "</script>"
	response.end
  end if
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Third Step to Retrieve Password-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/loginstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

<script language="JavaScript">
function checkform(){
	if (document.getpsw.pw1.value.length ==0){
		alert("Please input the password.");
		document.getpsw.pw1.focus();
		return false;
	}
	if (document.getpsw.pw2.value.length ==0){
		alert("Please input the repeat password.");
		document.getpsw.pw2.focus();
		return false;
	}
	if (document.getpsw.pw1.value.length !=document.getpsw.pw2.value.length){
		alert("You have input two different passwords.");
		document.getpsw.pw1.focus();
		return false;
	}
	if (document.getpsw.pw1.value.length <6){
		alert("Password should be at least 6 characters.");
		document.getpsw.pw1.focus();
		return false;
	}
	document.getpsw.submit();
}	
</script>

	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
		<form method="post" name="getpsw" action="getpsw-3.asp">
			<div class="details_title">Third Step to Retrieve Password</div> 
			<div class="details_row">
				<input class="input_common" type="password" name="pw1" maxlength="16" placeholder="New Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw2" maxlength="16" placeholder="Repeat Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Next" onclick="javascript: checkform();">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Back" onclick="javascript: location.href='javascript:history.go(-1)';">
			</div>
			<input type="hidden" name="save" value="ok">
			<input type="hidden" name="UserAnswer" value="<%=request.form("UserAnswer")%>">
			<input type="hidden" name="userid" value="<%=request.form("userid")%>">
        </form>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call goldweb_check_path()

if request.cookies("goldweb")("userid")<>"" then
	conn.close
	set conn=nothing
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

if isnull(rs("UserAnswer"))=true or rs("UserAnswer")="" or isnull(rs("UserQuestion"))=true or rs("UserQuestion")=""  then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Sorry, you have not set password protect question and answer.\n\nPlease contact web admin to reset for you.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end If

UserQuestion=rs("UserQuestion")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Second Step to Retrieve Password-<%=ensitename%>-<%=siteurl%></title>
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
	if (document.getpsw.UserAnswer.value.length ==0){
		alert("Please input the Answer.");
		document.getpsw.UserAnswer.focus();
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
			<div class="details_title">Second Step to Retrieve Password</div> 
			<div class="details_row">
				<%=UserQuestion%>
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserAnswer" maxlength="50" placeholder="Answer" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Next" onclick="javascript: checkform();">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Back" onclick="javascript: location.href='javascript:history.go(-1)';">
			</div>
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

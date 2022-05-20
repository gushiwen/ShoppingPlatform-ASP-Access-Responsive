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

randomize
yzm=Int(8999*rnd()+1000)
yzm_a=Int(yzm/1000)
yzm_b=Int((yzm-yzm_a*1000)/100)
yzm_c=Int((yzm-yzm_a*1000-yzm_b *100)/10)
yzm_d=Int(yzm-yzm_a*1000-yzm_b *100-yzm_c*10)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>First Step to Retrieve Password-<%=ensitename%>-<%=siteurl%></title>
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
	if (document.getpsw.userid.value.length ==0){
		alert("Please fill in your account ID.");
		document.getpsw.userid.focus();
		return false;
	}
	if (document.getpsw.yzm.value.length==0){
		alert("Please fill in the verify code.");
		document.getpsw.yzm.focus();
		return false;
	}
	if (document.getpsw.yzm.value!="<%=yzm%>"){
		alert("Please fill in the correct verify code.");
		document.getpsw.yzm.focus();
		return false;
	}
	document.getpsw.submit();
}	
</script>

	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
		<form method="post" name="getpsw" action="getpsw-2.asp">
			<div class="details_title">First Step to Retrieve Password</div> 
			<div class="details_row">
				<input class="input_common" type="text"  name="userid" maxlength="16" placeholder="Account ID" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_verify" type="text" name="yzm" maxlength="16" placeholder="Verify Code" autocomplete="off">
				<div class="verifycode"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif"></div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Next" onclick="javascript: checkform();">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Back" onclick="javascript: location.href='javascript:history.go(-1)';">
			</div>
        </form>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

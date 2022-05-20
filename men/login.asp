<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
If request.cookies("goldweb")("userid")<>"" Then 
	conn.close
	set conn=Nothing
	response.write "<meta http-equiv='refresh' content='0;URL=user_center.asp'>"
End If 

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

		<title>Member Login-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/loginstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

		<script type="text/javascript">
		// 个人用户登陆
		function  CheckPLoginForm() 
		{
			if(document.ploginform.userid.value=="")
			{
				alert("Please input your account ID");
				return false;
			}
			if(document.ploginform.password.value=="") 
			{
				alert("Please input password!");
				return false; 
			}	
			if(document.ploginform.verifycode.value != "<%=yzm%>")
			{
				alert("Please input correct verify code!");
				return false;
			}

			document.ploginform.submit();
		}
		</script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
		<form name="ploginform" action="indexlogin.asp" method="post" onSubmit="return CheckPLoginForm();">
			<div class="details_title">Member Login</div> 
			<div class="details_row">
				<input class="input_common" type="text"  name="userid" maxlength="16" placeholder="Account ID" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="password" maxlength="16" placeholder="Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_verify" type="text" name="verifycode" maxlength="12" placeholder="Verify Code" autocomplete="off">
				<div class="verifycode"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif"></div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Log in">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Forgot Password" onclick="document.location.href='getpsw.asp';">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Member Register" onclick="document.location.href='reg_member.asp';">
			</div>
			<input type="hidden" name="login" value="ok">
			<input type="hidden" name="cook" value="0">
			<input type="hidden" name="referer" value="<%=Request.ServerVariables("HTTP_REFERER")%>">
        </form>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

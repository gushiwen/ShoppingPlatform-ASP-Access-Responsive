<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>First Step to Retrieve Password-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>

<%
'call aspsql()
call goldweb_check_path()

if request.cookies("goldweb")("userid")<>"" then
	conn.close
	set conn=Nothing
	response.write "<meta http-equiv='refresh' content='0;URL=user_center.asp'>"
end if

randomize
yzm=int(8999*rnd()+1000)
%>
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
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=enSiteName%> 
&gt; First Step to Retrieve Password
</DIV>
<!--页面位置开始-->


<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >

<form method="post" name="getpsw" action="getpsw-2.asp">
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
	<tr height="25">
		<td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> First: Input Account ID</td>
	</tr>
	<tr height="25">
		<td width="28%" align="right">Account ID:&nbsp;</td>
		<td width="72%" align="left"><input type="text" name="userid" maxlength="16" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana ; overflow:auto;"></td>
	</tr>
	<tr height="25">
		<td align="right">Verify Code:&nbsp;</td>
		<td align="left"><input type="text" name="yzm" maxlength="16" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana ; overflow:auto;">
<%
a=int(yzm/1000)
b=int((yzm-a*1000)/100)
c=int((yzm-a*1000-b*100)/10)
d=int(yzm-a*1000-b*100-c*10)
response.write "<img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&a&".gif><img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&b&".gif><img  align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&c&".gif><img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&d&".gif>"
%>
		</td>
	</tr>
	<tr height="58">
		<td align="right">&nbsp;</td>
		<td align="left">
			<input type="button" value="  Next  " onclick="javascript: checkform();" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
			<input type="button" value="  Back  " onclick="javascript: location.href='javascript:history.go(-1)';" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
		</td>
	</tr>
</table>
</form>

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
<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
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
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Second Step to Retrieve Password-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
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
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--????????????-->
<div class="body">

<!-- ??????????????????-->
<DIV class="nav blueT">Current: <%=enSiteName%>
&gt; Second Step to Retrieve Password
</DIV>
<!--??????????????????-->

<!--????????????-->
<DIV id=homepage_left>

<!--??????????????????-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--??????????????????-->

</DIV>
<!--????????????-->

<!--????????????????????????-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >

<form method="post" name="getpsw" action="getpsw-3.asp">
<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
	<tr height="25">
		<td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> Second: Please answer the question</td>
	</tr>
	<tr height="25">
		<td width="28%" align="right">Question:&nbsp;</td>
		<td width="72%" align="left"><%=UserQuestion%></td>
	</tr>
	<tr height="25">
		<td align="right">Answer:&nbsp;</td>
		<td align="left"><input type="text" name="UserAnswer" maxlength="50" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana ; overflow:auto;"></td>
	</tr>
	<tr height="58">
		<td align="right">&nbsp;</td>
		<td align="left">
			<input type="hidden" name="userid" value="<%=request.form("userid")%>">
			<input type="button" value="  Next  " onclick="javascript: checkform();" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
			<input type="button" value="  Back  " onclick="javascript: location.href='javascript:history.go(-1)';"" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
		</td>
	</tr>
</table>
</form>

</DIV>
</DIV>
<!--????????????????????????-->

<div class="clear"></div>

</div>
<!--????????????-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>

<%
conn.close
set conn=nothing
%>
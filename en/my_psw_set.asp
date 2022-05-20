<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
'call aspsql()
call checklogin()
if request("set")="ok" then call setq()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Retrieve Password Setting-<%=ensitename%>-<%=siteurl%></title>
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
&gt; Retrieve Password Setting
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>
<!--#include file="goldweb_usertree.asp" -->
</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;height:180px;" >

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

	 <table width="80%" border="0" cellpadding="5" cellspacing="5" align="center">
        <form name="myinfo" action="my_psw_set.asp" method="post">
		<tr height="25"> 
          <td width="28%" align="right">Account ID: &nbsp;</td>
          <td width="72%" align="left"><%=rs("UserID")%><input type="hidden" name="userid" value="<%=rs("UserID")%>"></td>
        </tr>
        <tr height="24">
          <td align="right">Password: &nbsp;</td>
          <td align="left"><input size="35" type="password" name="Password" maxlength="16" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana"></td>
        </tr>
        <tr height="24"> 
          <td align="right">Question: &nbsp;</td>
          <td align="left"><input size="35" type="text" name="UserQuestion" value="<%=rs("UserQuestion")%>" maxlength="50" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana"> &nbsp; <img src="../images/admin/memo.gif" alt="Question to retrieve password"></td>
        </tr>
        <tr height="24"> 
          <td align="right">Answer: &nbsp;</td>
          <td align="left"><input size="35" type="text" name="UserAnswer1" value="" maxlength="50" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana "> &nbsp; <img src="../images/admin/memo.gif" alt="Answer to retrieve question"></td>
        </tr>
        <tr height="24"> 
          <td align="right">Repeat Answer: &nbsp;</td>
          <td align="left"><input size="35" type="text" name="UserAnswer2" value="" maxlength="50" style="height=20; BORDER: darkgray 1px solid; FONT-SIZE: 8pt; COLOR: #666666; FONT-FAMILY: verdana"> &nbsp; <img src="../images/admin/memo.gif" alt="Confirm the answer"></td>
        </tr>
	    <tr height="58">
		  <td colspan="2" align="center">
		    <input type="submit" name="Submit" value="  Submit  " style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="reset" name="Reset" value="  Reset  " style="font-size:10px; padding:2px;">
			<input type="hidden" name="set" value="ok">
	      </td>
		</tr>
        </form>
      </table>
<%
rs.close
set rs=nothing
%>	
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
end If

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

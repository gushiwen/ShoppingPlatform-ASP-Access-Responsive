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
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Member Login-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script type="text/javascript">
// 个人用户登陆
function  CheckPLoginForm() 
{
    if(document.ploginform.userid.value=="")
    {
		alert("Please input your account name!");
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
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; Member Login
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

	 <table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
	 <form name="ploginform" action="indexlogin.asp" method="post">
	    <tr height="25"> 
          <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> <font color=red>Required</font></td>
        </tr>
		<tr height="25">
          <td width="28%" align="right">Account ID: &nbsp;</td>
          <td width="72%" align="left"><input name="userid" type="text" size="30" maxlength="16" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;"></td>
        </tr>
        <tr height="25">
          <td align="right">Password: &nbsp;</td>
          <td align="left"><input name="password" type="password" size="31" maxlength="16" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;"></td>
        </tr>

        <tr height="25"> 
          <td align="right">Verify Code:&nbsp;</td>
          <td align="left"><input name="verifycode" type="text" size="30" maxlength="12" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;">&nbsp;&nbsp;<img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif">
	      </td> 
        </tr>

	    <tr height="58">
		  <td align="right">&nbsp;</td>
		  <td align="left">
		    <input type="button" value="  Log in  " onclick="javascript: CheckPLoginForm();" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="button" value="  register  " onclick="document.location.href='reg_member.asp';" style="font-size:10px; padding:2px;">&nbsp;&nbsp;&nbsp;&nbsp; 
			<a href="getpsw.asp">Forgot Password</a>
			<input type="hidden" name="login" value="ok">
			<input type="hidden" name="cook" value="0">
			<input type="hidden" name="referer" value="<%=Request.ServerVariables("HTTP_REFERER")%>">
	      </td>
		</tr>
        </form>
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

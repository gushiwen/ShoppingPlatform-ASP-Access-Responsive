<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()
if request("edit")="ok" then call edit()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Personal Infomation-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script type="text/javascript">
//判断表单输入正误
function Checkmodify()
{
	if (document.myinfo.UserName.value.length < 2 || document.myinfo.UserName.value.length >30) {
		alert("Contact Name should be 2-30 characters.");
		document.myinfo.UserName.focus();
		return false;
	}
	if (document.myinfo.HomePhone.value.length <6 || document.myinfo.HomePhone.value.length >20) {
		alert("Telephone Number should be 6-20 numbers.");
		document.myinfo.HomePhone.focus();
		return false;
	}
	if (document.myinfo.UserMail.value.length <10 || document.myinfo.UserMail.value.length >50) {
		alert("Please input a valid email address.");
		document.myinfo.UserMail.focus();
		return false;
	}
	if (document.myinfo.UserMail.value.length > 0 && !document.myinfo.UserMail.value.match( /^.+@.+$/ ) ) {
	    alert("Please input a valid email address.");
		document.myinfo.UserMail.focus();
		return false;
	}
	if (document.myinfo.Address.value.length <3 || document.myinfo.Address.value.length >150) {
		alert("Contact address should be 3-100 characters.");
		document.myinfo.Address.focus();
		return false;
	}
	if (document.myinfo.Country.value.length <3 || document.myinfo.Country.value.length >50) {
		alert("Country should be 3-50 characters.");
		document.myinfo.Address.focus();
		return false;
	}
	if (document.myinfo.ZipCode.value.length <4 || document.myinfo.ZipCode.value.length >12) {
		alert("Post code should be 4-10 characters.");
		document.myinfo.ZipCode.focus();
		return false;
	}
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
&gt; Personal Infomation
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>
<!--#include file="goldweb_usertree.asp" -->
</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

<table width="90%" border="0" cellpadding="5" cellspacing="5" align="center">
  <form name="myinfo" action="my_info.asp" method="post" onSubmit="return Checkmodify();">
          <tr height="25"> 
            <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> <font color=red>Required</font></td>
          </tr>
          <tr height="24"> 
            <td width="40%" align="left"><b>&nbsp;Account ID: </b></td>
            <td width="60%" align="left"><%=rs("UserID")%></td>
          </tr>

            <tr height="24" valign="top"> 
              <td align="left"><b>&nbsp;Account Type (with Discount): </b></td>
              <td align="left"> 
            <%
			    If rs("UserType")="1" Then response.write enusertype1
			    If rs("UserType")="2" Then response.write enusertype2
			    If rs("UserType")="3" Then response.write enusertype3
			    If rs("UserType")="4" Then response.write enusertype4
			    If rs("UserType")="5" Then response.write enusertype5
			    If rs("UserType")="6" Then response.write enusertype6
				If rs("UserKou")<>10 Then response.write " (-"&(100-(rs("UserKou"))*10)&"%)" ' 修改时折扣与类型无关
			    If rs("UserType")="1" Then response.write "<br><a href=""productorder.asp?Prodid=00001""><b>(Upgrade to Permanent VIP Member to Enjoy -10% Discount)</b></a>"
			%>
		      </td>
            </tr>

          <tr height="24">
            <td align="left"><b>&nbsp;Contact Name: </b></td>
            <td align="left"><input name="UserName" value="<%=rs("UserName")%>" type="text" size="25" maxlength="30" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
            </td> 
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Telephone Number: </b></td>
            <td align="left"><input name="HomePhone" value="<%=rs("HomePhone")%>" type="text" size="25" maxlength="20" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			</td> 
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Email: </b></td>
            <td align="left"><input name="UserMail" value="<%=rs("UserMail")%>" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
            </td> 
          </tr>

            <tr height="24"> 
              <td align="left"><b>&nbsp;Address: </b></td>
              <td align="left"><input name="Address" value="<%=rs("Address")%>" type="text" size="25" maxlength="150" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
 			</td> 
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Country: </b></td>      
              <td align="left"><input name="Country" value="<%=rs("Country")%>" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			  </td>
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Post Code: </b></td>
              <td align="left"><input name="ZipCode" value="<%=rs("ZipCode")%>" type="text" size="25" maxlength="12" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			 </td> 
            </tr>

          <tr height="6">
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr height="25"> 
            <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> <font color=red>Optional</font></td>
          </tr>  

            <tr height="24">
              <td align="left"><b>&nbsp;Company Name: </b></td>
              <td align="left"><input name="Compname" value="<%=rs("CompName")%>" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
              </td> 
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Alternative Phone Number: </b></td>
              <td align="left"><input name="CompPhone" value="<%=rs("CompPhone")%>" type="text" size="25" maxlength="20" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			  </td>
            </tr>

            <tr height="26"> 
              <td align="left"><b>&nbsp;Gender: </b></td>
              <td align="left"> <input name="Sex" type="radio" value="1" <%if rs("Sex")="1" then response.write "checked"%>><img border=0 src=../images/small/boy.gif> Male &nbsp;&nbsp;
                <input type="radio" name="Sex" value="0" <%if rs("Sex")="0" then response.write "checked"%>><img border=0 src=../images/small/girl.gif> Female 
		      </td>
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Birthday: </b>(Format 1980-01-01)</td>
              <td align="left"><input type="text" name="Birthday" value="<%=rs("Birthday")%>" size="25" maxlength="10" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			 </td> 
            </tr>

            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Nearest Branch: </b></td>
              <td align="left">
				<select name="BranchId" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
					<option value=''>Please choose branch</option>
	<%
		Set rs_branch = conn.Execute("select BranchId,enBranchName from branch order by BranchId asc") 
		While Not rs_branch.eof
			response.write "<option value='"&rs_branch("BranchId")&"' "
			If rs_branch("BranchId") = rs("BranchId") Then response.write "selected "
			response.write ">"&rs_branch("enBranchName")&"</option>"
			rs_branch.movenext
		Wend 
		rs_branch.close
		Set rs_branch = nothing
	%>
				</select>
			 </td> 
            </tr>

            <tr height="68"> 
              <td align="left"><b>&nbsp;Memo: </b><br>
                &nbsp;Introduction or Request</td>
              <td align="left"><textarea name="Memo" cols="20" rows="4" style="BORDER: #CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666;  overflow:auto;"><%=rs("Memo")%></textarea>
			  </td>
            </tr>

          <tr height="38"> 
            <td align="center" colspan="2">
              <input type="submit" value="  Submit form  " name="Submit" style="font-size:10px; padding:2px;"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			  <input type="reset" value="  Reset form  " name="Reset" style="font-size:10px; padding:2px;">
			  <input type="hidden" name="edit" value="ok">
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
'修改资料
sub edit()
call goldweb_check_path()

'Required
UserName=Trim(request.form("UserName"))
HomePhone=Trim(request.form("HomePhone"))
UserMail=Trim(request.form("UserMail"))
Address=Trim(request.form("Address"))
Country=Trim(request.form("Country"))
ZipCode=Trim(request.form("ZipCode"))

'Optional
CompName=Trim(request.form("CompName"))
CompPhone=Trim(request.form("CompPhone")) 'Mobile Number
Sex=request.form("Sex")
Birthday=Trim(request.form("Birthday"))
BranchId=Trim(request.form("BranchId"))
Memo=Trim(request.form("Memo"))

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("UserName")=UserName
rs("HomePhone")=HomePhone
rs("UserMail")=UserMail
rs("Address")=Address
rs("Country")=Country
rs("ZipCode")=ZipCode

rs("CompName")=CompName
rs("CompPhone")=CompPhone
rs("Sex")=Sex
rs("Birthday")=Birthday
rs("BranchId")=BranchId
rs("Memo")=Memo
rs.update
rs.close
set rs = Nothing
conn.close
set conn=nothing
response.write "<script language='javascript'>"
response.write "alert('Your personal information has been changed.');"
response.write "location.href='my_info.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>

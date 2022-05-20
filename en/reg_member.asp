<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->
<%
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

<title>Member Register-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script language="JavaScript">
function JM_sh(ob){
if (ob.style.display=="none"){ob.style.display=""}else{ob.style.display="none"};
}

function fucPWDchk(str) 
{ 
var strSource ="0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"; 
var ch; 
var i; 
var temp; 

for (i=0;i<=(str.length-1);i++) 
{ 

ch = str.charAt(i); 
temp = strSource.indexOf(ch); 
if (temp==-1) 
{ 
return 0; 
} 
} 
if (strSource.indexOf(ch)==-1) 
{ 
return 0; 
} 
else 
{ 
return 1; 
} 
} 
function jtrim(str) 
{ while (str.charAt(0)==" ") 
{str=str.substr(1);} 
while (str.charAt(str.length-1)==" ") 
{str=str.substr(0,str.length-1);} 
return(str); 
} 

//判断表单输入正误
function Checkreg()
{
	if (document.ADDUser.UserId.value.length < 4 || document.ADDUser.UserId.value.length >16) {
		alert("Account ID should be 4-16 characters.");
		document.ADDUser.UserId.focus();
		return false;
	}
	if (document.ADDUser.pw1.value.length <6 || document.ADDUser.pw1.value.length >16) {
		alert("Password should be 6-16 characters.");
		document.ADDUser.pw1.focus();
		return false;
	}
	if (!fucPWDchk(document.ADDUser.pw1.value)){
		alert("Password should be numbers, capital or small letters.");
		document.ADDUser.pw1.focus();
		return false;
		}
	if (document.ADDUser.pw1.value != document.ADDUser.pw2.value) {
		alert("You have entered  two different passwords.");
		document.ADDUser.pw2.focus();
		return false;
	}
	if (document.ADDUser.UserName.value.length < 2 || document.ADDUser.UserName.value.length >30) {
		alert("Contact Name should be 2-30 characters.");
		document.ADDUser.UserName.focus();
		return false;
	}
	if (document.ADDUser.HomePhone.value.length <6 || document.ADDUser.HomePhone.value.length >20) {
		alert("Telephone Number should be 6-20 numbers.");
		document.ADDUser.HomePhone.focus();
		return false;
	}
	if (document.ADDUser.UserMail.value.length <10 || document.ADDUser.UserMail.value.length >50) {
		alert("Please input a valid email address.");
		document.ADDUser.UserMail.focus();
		return false;
	}
	if (document.ADDUser.UserMail.value.length > 0 && !document.ADDUser.UserMail.value.match( /^.+@.+$/ ) ) {
	    alert("Please input a valid email address.");
		document.ADDUser.UserMail.focus();
		return false;
	}
	if (document.ADDUser.Address.value.length <3 || document.ADDUser.Address.value.length >150) {
		alert("Contact address should be 3-100 characters.");
		document.ADDUser.Address.focus();
		return false;
	}
	if (document.ADDUser.Country.value.length <2 || document.ADDUser.Country.value.length >50) {
		alert("Country should be 3-50 characters.");
		document.ADDUser.Address.focus();
		return false;
	}
	if (document.ADDUser.ZipCode.value.length <4 || document.ADDUser.ZipCode.value.length >12) {
		alert("Post code should be 4-10 characters.");
		document.ADDUser.ZipCode.focus();
		return false;
	}
	if (document.ADDUser.VerifyCode.value != "<%=yzm%>") {
		alert("Please input correct verify code.");
		document.ADDUser.VerifyCode.focus();
		return false;
	}
	document.ADDUser.Submit.disabled=true;
	return true;
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
&gt; Member Register
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
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr><td>

<table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
 <form name="ADDUser" method="POST" action="reg_save.asp" onSubmit="return Checkreg();">
          <tr height="25"> 
            <td colspan="2" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> <font color=red>Required</font></td>
          </tr>
          <tr height="24">
            <td width="40%" align="left"><b>&nbsp;Account ID (4-16 characters): </b><br>&nbsp;Numbers, letters only</td>
            <td width="60%" align="left"><input name="UserId" value="" type="text" size="24" maxlength="16" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "  onkeyup="value=value.replace(/[^a-zA-Z0-9]/g,'')">
			</td>
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Password (6-16 characters): </b></td>
            <td align="left"><input name="pw1" type="password" size="25" maxlength="16" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
            </td> 
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Repeat Password: </b></td>
            <td align="left"><input name="pw2" type="password" size="25" maxlength="16" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
            </td> 
          </tr>

            <%
			  '账户类型所在行
			  NormalAccounts=""
			  SpecialAccounts=""
			  BasicKou=9 ' 比它大的放NormalAccounts，比它小的放SpecialAccounts
			  If enusertype1<>"" and  instr(enusertype1,"VIP")=0 Then
			    If kou1>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1
				  If kou1<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou1*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			  If enusertype2<>"" and  instr(enusertype2,"VIP")=0 Then
			    If kou2>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""2""> "&enusertype2
				  If kou2<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou2*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""2""> "&enusertype2&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			  If enusertype3<>"" and  instr(enusertype3,"VIP")=0 Then
			    If kou3>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""3""> "&enusertype3
				  If kou3<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou3*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""3""> "&enusertype3&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			  If enusertype4<>"" and  instr(enusertype4,"VIP")=0 Then
			    If kou4>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""4""> "&enusertype4
				  If kou4<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou4*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""4""> "&enusertype4&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			  If enusertype5<>"" and  instr(enusertype5,"VIP")=0 Then
			    If kou5>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""5""> "&enusertype5
				  If kou5<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou5*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""5""> "&enusertype5&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			  If enusertype6<>"" and  instr(enusertype6,"VIP")=0 Then
			    If kou6>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<input type=""radio"" name=""UserType"" value=""6""> "&enusertype6
				  If kou6<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou6*10)&"%)"
				  NormalAccounts=NormalAccounts&" &nbsp;&nbsp;"
				Else
				  SpecialAccounts=SpecialAccounts&"<input type=""radio"" name=""UserType"" value=""6""> "&enusertype6&" (Special) &nbsp;&nbsp;"
				End If 
			  End If 
			%>
            <tr height="58"> 
              <td align="left"><b>&nbsp;Account Type (with Discount): </b><br>&nbsp;<font color="red">Can not change; Evaluation Required</font></td>
              <td align="left"> 
                 <%
				   If NormalAccounts<>"" Then response.write NormalAccounts
				   If NormalAccounts<>"" And SpecialAccounts<>"" Then response.write "<br>"
				   If SpecialAccounts<>"" Then response.write "<font color=""red"">"&SpecialAccounts&"</font>"
				 %>
		      </td>
            </tr>

          <tr height="24">
            <td align="left"><b>&nbsp;Contact Name: </b><br>&nbsp;Numbers, letters, spaces only</td>
            <td align="left"><input name="UserName" value="" type="text" size="25" maxlength="30" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "  onkeyup="value=value.replace(/[^a-zA-Z0-9 ]/g,'')">
            </td> 
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Mobile Phone Number: </b></td>
            <td align="left"><input name="HomePhone" value="" type="text" size="25" maxlength="20" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			</td> 
          </tr>
          <tr height="24"> 
            <td align="left"><b>&nbsp;Email: </b></td>
            <td align="left"><input name="UserMail" value="" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
            </td> 
          </tr>

            <tr height="24"> 
              <td align="left"><b>&nbsp;Address: </b></td>
              <td align="left"><input name="Address" value="" type="text" size="25" maxlength="150" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
 			</td> 
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Country: </b></td>      
              <td align="left"><input name="Country" value="Singapore" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			  </td>
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Post Code: </b></td>
              <td align="left"><input name="ZipCode" value="" type="text" size="25" maxlength="12" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			 </td> 
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Verify Code: </b></td>
              <td align="left"><input name="VerifyCode" value="" type="text" size="25" maxlength="12" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">&nbsp;&nbsp;<font color="#FF0000">Keyin</font>&nbsp;<img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif">
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
              <td align="left"><input name="CompName" value="" type="text" size="25" maxlength="50" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
              </td>
            </tr>
            <tr height="24">
              <td align="left"><b>&nbsp;Other Contact Number: </b></td>
              <td align="left"><input name="CompPhone" value="" type="text" size="25" maxlength="20" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			  </td>
            </tr>

            <tr height="26"> 
              <td align="left"><b>&nbsp;Gender: </b></td>
              <td align="left"> 
			    <input type="radio" name="Sex" value="1" checked><img border=0 src="../images/small/boy.gif"> Male &nbsp;&nbsp;
                <input type="radio" name="Sex" value="0"><img border=0 src="../images/small/girl.gif"> Female 
		      </td>
            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Birthday: </b>(Format 1980-01-01)</td>
              <td align="left"><input type="text" value="" name="Birthday"  size="25" maxlength="10" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
			 </td> 

            </tr>
            <tr height="24"> 
              <td align="left"><b>&nbsp;Nearest Branch: </b></td>
              <td align="left">
				<select name="BranchId" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; ">
					<option value=''>Please choose branch</option>
	<%
		Set rs_branch = conn.Execute("select BranchId,enBranchName from branch order by BranchId asc") 
		While Not rs_branch.eof
			response.write "<option value='"&rs_branch("BranchId")&"'>"&rs_branch("enBranchName")&"</option>"
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
              <td align="left"><textarea name="Memo" cols="20" rows="4" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;  overflow:auto;"></textarea>
			  </td>
            </tr>

          <tr height="38"> 
            <td align="center" colspan="2">
              <input type="submit" value="  Submit form  " name="Submit" style="font-size:10px; padding:2px;"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			  <input type="reset" value="  Reset form  " name="Reset" style="font-size:10px; padding:2px;">
			  <input type="hidden" name="adduser" value="true">
		    </td>  
          </tr>
  </form>
</table>

</td></tr>
</table>
</DIV>
</DIV>
<!--右边结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>
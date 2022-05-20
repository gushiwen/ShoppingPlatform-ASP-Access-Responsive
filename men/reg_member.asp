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
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Member Register-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/regmemberstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

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
			{ 
				while (str.charAt(0)==" ") 
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
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
		<form name="ADDUser" method="POST" action="reg_save.asp" onSubmit="return Checkreg();">
			<div class="details_title">Member Register</div> 
			<div class="details_row">
				<input class="input_common" type="text"  name="UserId" value="" maxlength="16" placeholder="Account ID (Numbers, Letters Only)" autocomplete="off" onkeyup="value=value.replace(/[^a-zA-Z0-9]/g,'')">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw1" maxlength="16" placeholder="Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw2" maxlength="16" placeholder="Repeat Password" autocomplete="off">
			</div>

            <%
			  '账户类型所在行
			  NormalAccounts=""
			  SpecialAccounts=""
			  BasicKou=9 ' 比它大的放NormalAccounts，比它小的放SpecialAccounts
			  If enusertype1<>"" and  instr(enusertype1,"VIP")=0 Then
			    If kou1>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1
				  If kou1<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou1*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1&" (Special)</li>"
				End If 
			  End If 
			  If enusertype2<>"" and  instr(enusertype2,"VIP")=0 Then
			    If kou2>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""2""> "&enusertype2
				  If kou2<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou2*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""2""> "&enusertype2&" (Special)</li>"
				End If 
			  End If 
			  If enusertype3<>"" and  instr(enusertype3,"VIP")=0 Then
			    If kou3>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""3""> "&enusertype3
				  If kou3<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou3*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""3""> "&enusertype3&" (Special)</li>"
				End If 
			  End If 
			  If enusertype4<>"" and  instr(enusertype4,"VIP")=0 Then
			    If kou4>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""4""> "&enusertype4
				  If kou4<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou4*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""4""> "&enusertype4&" (Special)</li>"
				End If 
			  End If 
			  If enusertype5<>"" and  instr(enusertype5,"VIP")=0 Then
			    If kou5>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""5""> "&enusertype5
				  If kou5<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou5*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""5""> "&enusertype5&" (Special)</li>"
				End If 
			  End If 
			  If enusertype6<>"" and  instr(enusertype6,"VIP")=0 Then
			    If kou6>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""6""> "&enusertype6
				  If kou6<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou6*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""6""> "&enusertype6&" (Special)</li>"
				End If 
			  End If 
			%>
			<div class="details_row">
				<ul class="ul_usertype">
				<%
				   If NormalAccounts<>"" Then response.write NormalAccounts
				   If SpecialAccounts<>"" Then response.write SpecialAccounts
				%>
				</ul>
			</div>

			<div class="details_row">
				<input class="input_common" type="text" name="UserMail" maxlength="50" placeholder="Email Address" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_verify" type="text" name="VerifyCode" maxlength="12" placeholder="Verify Code" autocomplete="off">
				<div class="verifycode"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif"></div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<input type="hidden" name="adduser" value="true">
        </form>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

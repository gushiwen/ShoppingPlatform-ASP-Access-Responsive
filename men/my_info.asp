<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
call checklogin()
if request("edit")="ok" then call edit()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Personal Infomation-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0 ,minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/myinfostyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

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
		<!--#include file="goldweb_top.asp"-->

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

		<!--客户详细资料-->
		<div class="account_details">
		<form name="myinfo" action="my_info.asp" method="post" onSubmit="return Checkmodify();">
			<div class="details_title">Personal Infomation</div> 

			<div class="details_row">
				<input class="input_common" type="text" name="UserName" value="<%=rs("UserName")%>" maxlength="30" placeholder="Full Name Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="HomePhone" value="<%=rs("HomePhone")%>" maxlength="20" placeholder="Mobile Number Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserMail" value="<%=rs("UserMail")%>" maxlength="50" placeholder="Email Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="Address" value="<%=rs("Address")%>" maxlength="150" placeholder="Detailed Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="Country" value="<%=rs("Country")%>" maxlength="50" placeholder="Country Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="ZipCode" value="<%=rs("ZipCode")%>" maxlength="12" placeholder="Post Code Required" autocomplete="off">
			</div>

			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<input type="hidden" name="edit" value="ok">
		 </form>
		</div>



		<!--#include file="goldweb_down.asp"-->
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

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("UserName")=UserName
rs("HomePhone")=HomePhone
rs("UserMail")=UserMail
rs("Address")=Address
rs("Country")=Country
rs("ZipCode")=ZipCode
rs.update
rs.close
set rs = Nothing
conn.close
set conn=nothing
response.write "<script language='javascript'>"
response.write "alert('Personal information has been changed.');"
response.write "location.href='my_info.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>

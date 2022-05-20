<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
OrderNum=request("OrderNum")
MobilePhone=request("MobilePhone")

if request("send")="ok" then 
	call send()
else

If request.cookies("goldweb")("userid")<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and UserId='"&request.cookies("goldweb")("userid")&"'"
ElseIf request.cookies("goldweb")("userid")="" And MobilePhone<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and RecHomePhone='"&MobilePhone&"'"
Else	
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('You are not authorized to view this order.');"
	response.write "location.href='main.asp';"
	response.write "</script>"
	response.end
End If 

set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sqlinfo,conn,1,1
if rs.eof and rs.bof  then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('We can not find this order.');"
	response.write "location.href='main.asp';"
	response.write "</script>"
	response.end
Else
	OrderSum=formatnumber(rs("OrderSum"),2)
	UserName=rs("RecName")
	HomePhone=rs("RecHomePhone")
	UserMail=rs("RecMail")
	Address=rs("RecAddress")
	Country=rs("RecCountry")
	ZipCode=rs("RecZipcode")
End If 

rs.close
set rs=Nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Remittance Report-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0 ,minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/myinfostyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>

	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="account_details">
		<form name="add" method="post" action="my_remittance.asp">
			<div class="details_title">Remittance Report</div> 

			<div class="details_row">
				<input class="input_common" type="text" name="hk1" maxlength="14" value="<%=OrderNum%>" placeholder="Order No. Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="hk2" maxlength="10" value="<%=englobalpriceunit%> <%=OrderSum%>" placeholder="Payment Amount Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="hk3" maxlength="25" placeholder="From Bank Name Required" autocomplete="off">
			</div>

			<div class="details_row">
				<select class="input_common" name="hk4">
					<option selected value="">Payment by Required</option>
<%
	 ' 列出银行转账方式
	set rs_pay_bank=conn.execute("select * from pay_type where PayTypeClass='bank' and Display=true order by PayType")
	if rs_pay_bank.eof and rs_pay_bank.bof  then
	Else
	  do while not rs_pay_bank.eof
	    response.write "<option value="""&rs_pay_bank("enPayTypeDefine")&"""> "&rs_pay_bank("enPayTypeDefine")&"</option>"
	    rs_pay_bank.movenext
	  Loop
	End If 
	rs_pay_bank.close
	set rs_pay_bank=nothing
	%>
				</select>
			</div>

			<div class="details_row">
				<textarea class="textarea_common" name="hk11" placeholder="Notes ≤200 Characters"></textarea>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="document.location.href='my_order_view.asp?OrderNum=<%=OrderNum%>&MobilePhone=<%=server.URLEncode(MobilePhone)%>';">
			</div>
			<input name="send" type="hidden" value="ok">
			<input name="OrderNum" type="hidden" value="<%=OrderNum%>">
			<input name="MobilePhone" type="hidden" value="<%=MobilePhone%>">
			<input name="hk5" type="hidden" value="<%=UserName%>">
			<input name="hk6" type="hidden" value="<%=HomePhone%>">
			<input name="hk7" type="hidden" value="<%=UserMail%>">
			<input name="hk8" type="hidden" value="<%=Address%>">
			<input name="hk9" type="hidden" value="<%=Country%>">
			<input name="hk10" type="hidden" value="<%=ZipCode%>">
		 </form>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
End If

conn.close
set conn=nothing

sub send()
	call goldweb_check_path()

	if request.form("hk1")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Order Number.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
	
	if request.form("hk2")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Payment Amount.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk3")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in From Bank Name.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk4")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please choose Payment By.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if llen(request.form("hk11"))>200  then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('The note is too long.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	dim MailTitle
	dim MailContent
	MailTitle = request.form("hk5") & " has reported remittance on " & siteurl
	MailContent="Order No.: " & request.form("hk1")  & "<br>"
	MailContent=MailContent & "Payment Amount: " & request.form("hk2")  & "<br>"
	MailContent=MailContent & "From Bank Name: " & request.form("hk3")  & "<br>"
	MailContent=MailContent & "Payment by: " & request.form("hk4")  & "<br>"
	MailContent=MailContent & "Full Name: " & request.form("hk5")  & "<br>"
	MailContent=MailContent & "Mobile Number: " & request.form("hk6")  & "<br>"
	MailContent=MailContent & "Email: " & request.form("hk7")  & "<br>"
	MailContent=MailContent & "Address: " & request.form("hk8")  & "<br>"
	MailContent=MailContent & "Country: " & request.form("hk9")  & "<br>"
	MailContent=MailContent & "Post Code: " & request.form("hk10")  & "<br>"
	MailContent=MailContent & "Note: " & request.form("hk11")  & "<br>"
	Call SendOutMail(MailTitle, MailContent, adm_mail, adm_mail, request.form("hk7"))

	sql="select * from hand"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
  	rs("name")="admin"
  	rs("neirong")=MailContent
  	rs("riqi")=DateAdd("h", TimeOffset, now())
	If request.cookies("goldweb")("userid")<>"" Then 
 		rs("fname")=request.cookies("goldweb")("userid")
	Else
 		rs("fname")="Guest"
	End If 
  	rs("zuti")="1"
	rs.update	
	rs.close
	set rs=nothing

	conn.close
	set conn=Nothing

	response.write "<script language='javascript'>"
	response.write "alert('The message has been sent.');"
	response.write "</script>"
	response.write "<meta http-equiv=refresh content='0;URL=my_order_view.asp?OrderNum="&OrderNum&"&MobilePhone="&server.URLEncode(MobilePhone)&"'>"
end sub
%>
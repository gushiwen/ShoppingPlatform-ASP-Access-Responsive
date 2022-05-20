<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
'call aspsql()
call goldweb_check_path()

'Required
UserId=Trim(request.form("UserId"))
User_Password=request.form("pw1")
UserPassword=md5(User_Password)
UserType=request.form("UserType")
UserMail=Trim(request.form("UserMail"))
VerifyCode=Trim(request.form("VerifyCode"))

if InStr(UserId, "@")>0 then
	response.write "<script language='javascript'>"
	response.write "alert('Account ID can only be letters or numbers.');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end If

If reg_check=0 Then '会员注册是否需要审查
  status="1"
Else 
  status="0"
End If 

' 注册时根据类型初始化折扣，以后修改时折扣与类型无关
if UserType="1" then
UserKou=kou1
elseif UserType="2" then
UserKou=kou2
elseif UserType="3" then
UserKou=kou3
elseif UserType="4" then
UserKou=kou4
elseif UserType="5" then
UserKou=kou5
elseif UserType="6" then
UserKou=kou6
end If

if instr(lcase(UserId),"admin")>0 or instr(lcase(UserId),"administrator")>0  or instr(lcase(UserId),"guest")>0 then
	response.write "<script language='javascript'>"
	response.write "alert('Account ID is invalid.');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if

set rs = conn.execute ("SELECT * FROM goldweb_user where UserId= '" & UserId & "'")
if Not(rs.Bof and rs.eof) Then
    rs.close
    set rs=Nothing
    conn.close
    set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('This Account ID has been taken already!');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
Else
    rs.close
    set rs=Nothing

	sqlinfo = "select * from goldweb_user"
	set rs=Server.Createobject("ADODB.RecordSet")
	rs.Open sqlinfo,conn,1,3
	rs.AddNew
	rs("UserId")=UserId
	rs("UserPassword")=UserPassword
	rs("UserType")=UserType
	rs("UserMail")=UserMail

	rs("AdvDays")=30
	rs("userkou")=userkou
	rs("status")=status
	rs("SignDate")=DateAdd("h", TimeOffset, now())
	rs("IP")=Request.serverVariables("REMOTE_ADDR")
	rs.Update
	rs.close
	set rs = nothing

	'邮件通知开始
	'发信功能打开时继续执行
	if reg_mailyesorno=1 Then 
	dim MailTitle
	dim MailContent
	MailTitle = UserId&" has registered on "&siteurl&" successfully."
	MailContent = "Dear "&UserId&",<br><br>You have successfully registered on "&siteurl&"<br>Account ID: "&UserId&"<br><br>Now you can login to our website to order products online." & "<br><br>"
	MailContent = MailContent & "Your sincerely<br>"&ensitename&"<br>"&"Tel: "&adm_tel&"<br>"&FullSiteUrl

	Call SendOutMail(MailTitle, MailContent, mailname, UserMail, adm_mail)
End  If 
'邮件通知结束

If status="1" Then
	'登陆
	Response.cookies("goldweb").path="/"
	response.cookies("goldweb")("userid")=UserId
	'response.cookies("goldweb")("usertype")=UserType
	response.cookies("goldweb")("userkou")=UserKou

	'Initialise Session Variables for Forum Auto Login
	Session("USER") = UserId
	Session("PASSWORD") = User_Password
	Session("EMAIL") = UserMail

	response.write "<script language='javascript'>"
	response.write "alert('You have registered successfully.');"
	response.write "</script>"
	'使用cookies保存用户信息，因此，必须使用meta转向，若使用redirect转向可能cookies还未写入就已经转向了
	response.write "<meta http-equiv=refresh content='0;URL=user_center.asp'>"
Else
	response.write "<script language='javascript'>"
	response.write "alert('Your registry is pending for approval now.');"
	response.write "</script>"
	'使用cookies保存用户信息，因此，必须使用meta转向，若使用redirect转向可能cookies还未写入就已经转向了
	response.write "<meta http-equiv=refresh content='0;URL=main.asp'>"
End If

end if
%>
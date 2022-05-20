<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
' 登陆
call goldweb_check_path()

Userid=trim(request.form("userid"))
Password=trim(request.form("password"))
'cook=cint(request.form("cook")) 
referer=trim(request.form("referer")) 
userkou=10

if request.form("Login")<>"ok" then 
	conn.close
	set conn=nothing
	response.redirect "main.html"
End If 
if Userid = "" or Password ="" then 
	conn.close
	set conn=nothing
	response.redirect  "main.html"
End If 
if Userid = request.cookies("goldweb")("userid") then 
	conn.close
	set conn=nothing
	response.redirect "user_center.asp"
End If 

sql = "select * from goldweb_user where userid='"&Userid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,3
if (rs.bof and rs.eof) Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('This Account ID has not been registered.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end If

if rs("Status")<>"1" Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('This Account ID is pending for approval.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end If

if  rs("UserPassword")<> md5(Password) Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	session("login_error")=session("login_error")+1
	response.write "<script language='javascript'>"
	response.write "alert('Please enter the correct password.\n\nYou have entered wrongly for "&session("login_error")&" time(s).');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
Else
	rs("LastLogin")=DateAdd("h", TimeOffset, now())
	rs("IP")=Request.serverVariables("REMOTE_ADDR")
	rs("TotalLogin")=rs("TotalLogin")+1
    userkou=rs("userkou")
    usermail=rs("usermail")
	rs.update

	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing

	response.cookies("goldweb").path="/"
	response.cookies("goldweb")("userid")=userid
	response.cookies("goldweb")("userkou")=userkou

	'Initialise Session Variables for Forum Auto Login
	Session("USER") = userid
	Session("PASSWORD") = Password
	Session("EMAIL") = usermail

	'if request.form("cook")<>"0" then response.cookies("goldweb").expires=now+cook
	
	If referer<>"" And InStr(referer, "reg_member.asp")=0 And InStr(referer, "login.asp")=0 Then 'Go to previous page after login
	  response.write "<meta http-equiv='refresh' content='0;URL=" & referer & "'>"
	Else 
	  response.write "<meta http-equiv='refresh' content='0;URL=user_center.asp'>"
	End If 
end If
%>

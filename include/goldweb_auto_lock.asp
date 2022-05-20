<%
if lockip<>"0" then
if ip<>"" then
ip=checktext(ip)
lockip=split(cstr(ip),"@")
for N=0 to UBound(lockip)
if instr(Request.serverVariables("REMOTE_ADDR"),lockip(n))>0 then
response.redirect "shop_error.asp?error=99"
response.end
end if
next
end if
end if


'连续登陆X次密码出错，锁定IP
if session("login_error")>=10 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from shopsetup"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('您涉嫌非法猜解会员密码，已被系统限制访问。');"
response.write "location.href='main.asp';"
response.write "</script>"
response.end
end if

'使用“取回密码”时连续出错X次，锁定IP
if session("gpw_error")>=10 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from shopsetup"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('您涉嫌非法猜解会员密码，您已被系统限制访问。');"
response.write "location.href='main.asp';"
response.write "</script>"
response.end
end if

'一个会话期内（20分钟）连续发布留言X次，锁定IP
if session("book_error")>=5 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from shopsetup"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)<0 then rs("ip")=rs("ip")&"@"&userip
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('您涉嫌在本站留言本恶意灌水，您已被系统限制访问。');"
response.write "location.href='main.asp';"
response.write "</script>"
response.end
end if
%>

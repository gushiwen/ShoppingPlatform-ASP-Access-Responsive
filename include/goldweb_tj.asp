<!--#include file="goldweb_shop_30_conn.asp"-->

<%
' 流量统计
'where=Request.ServerVariables("HTTP_REFERER")
where=Request.QueryString("where")

'response.write "alert('test+"&where&"');"

IP = Request.ServerVariables("REMOTE_ADDR")
response.Cookies("goldweb_user_ip").path="/"
response.cookies("goldweb_user_ip")=ip

'upgrade visiting basic data
set rs=conn.Execute ("select * from count_total")
if rs("zzday") <> DateValue(DateAdd("h", TimeOffset, now())) Then ' date()
  conn.Execute("Update count_total set zzday=#" & DateValue(DateAdd("h", TimeOffset, now())) & "#,yesterday=today,today=1,total=total+1")
else
  conn.Execute("Update count_total set today=today+1,total=total+1")
end if

'upgrade online visitor
set rs=server.createobject("adodb.recordset")
sql="select * from count_online where ip='"&IP&"' and datediff('h',time,'" & DateAdd("h", TimeOffset, now()) & "')<1" ' now()
rs.open sql,conn,1,3
if (rs.eof and rs.bof) then 
    rs.addnew
    rs("time")=DateAdd("h", TimeOffset, now())
    rs("IP")=ip
    rs.update
end if
rs.close
set rs=nothing

'upgrade visiting record
set rs=server.createobject("adodb.recordset")
sql="select * from count_shop where  ip='"&IP&"' and #" & DateValue(DateAdd("h", TimeOffset, now())) & "#=day and datediff('h',times,'" & TimeValue(DateAdd("h", TimeOffset, now())) & "')<2" ' date() time()
rs.open sql,conn,1,3
    if (rs.eof and rs.bof) then
    rs.addnew
    rs("day")=DateValue(DateAdd("h", TimeOffset, now())) 
    rs("times")=TimeValue(DateAdd("h", TimeOffset, now()))
    rs("IP")=IP
    rs("where")=lleft(where,200)
    rs.update
    end if
rs.close
set rs=Nothing

conn.close
set conn=nothing
%>
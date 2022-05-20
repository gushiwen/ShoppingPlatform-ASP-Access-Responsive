<%
if request.form("save")="ok" And request.Form("yzm")=request.Form("yzmcheck") then 
call goldweb_check_path()
set rs=conn.execute("select * from book_setup")
jianju=clng(rs("book_jianju"))		'间距
view=cstr(rs("view1"))			'评论是否需要审核
mailyes=clng(rs("mailyes1"))		'是否必填邮箱

prodid=trim(request("prodid"))
name=trim(request.form("name"))
mail=trim(request.form("mail"))
nr=request.form("nr")
jibie=request.form("jibie")

    if mailyes=0 then		'邮箱为必填时检查邮箱是否合法
	if Instr(mail,".")<=0 or Instr(mail,"@")<=0 or len(mail)<10 or len(mail)>50 then
	response.write "<script language='javascript'>"
	response.write "alert('Please input an valid email address.');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

    end if

	set rs=Server.CreateObject("ADODB.RecordSet")
	sql="select * from goldweb_pinglun"
	rs.open sql,conn,1,3

			rs.Addnew
			rs("name")=name
			rs("prodid")=prodid
			rs("mail")=mail
			rs("nr")=nr
			rs("jibie")=jibie
			rs("view")=view
			rs("IP")=Request.serverVariables("REMOTE_ADDR")
			rs.Update
		rs.close
		set rs=nothing
	response.write "<script language='javascript'>"
	if view="0" then
	response.write "alert('Submit successfully. The comment will be displayed soon.');"
	else
	response.write "alert('Submit successfully.');"
	end if
	response.write "location.href='product.asp?prodid="&prodid&"';"
	response.write "</script>"
	response.end
end if
%>

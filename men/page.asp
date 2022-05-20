<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 公司信息
id=request("id")
If not isNumeric(id) Then
	conn.close
	set conn=nothing
	response.redirect "main.asp"
	response.end
end if

title="entitle"&cstr(id)
body="enbody"&cstr(id)
Set rsp = conn.Execute("select * from page")
title=rsp(title) 
body=rsp(body)
rsp.close
set rsp=Nothing

if body="" then 
	conn.close
	set conn=nothing
	response.redirect "main.asp"
	response.end
End If 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title><%=title%>-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/companystyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="list-box">
			<div class="list_title"><%=title%></div> 

			<div class="list_row">
				<div class='text_common'>
					<%= body%>
				</div>
			</div>

			<div class="list_row">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="document.location.href='company.asp';">
			</div> 

		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

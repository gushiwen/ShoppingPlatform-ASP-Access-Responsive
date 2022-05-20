<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=FullSiteUrl%>"+"/en/main.html");</script>
		<link href="<%=FullSiteUrl%>/en/main.html" rel="canonical" />

		<title>Company Information-<%=ensitename%>-<%=siteurl%></title>
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
			<div class="list_title">Company Information</div> 

			<%
				' 信息列表
				Set rsp = conn.Execute("select * from page") 

				for lie=1 to 5
					for hang=1 to help_hang
						ID=lie*10+hang
						ID=cstr(id)
						title="entitle"&id
						url="url"&id
						response.write "<div class='list_row'>"
						if rsp(url)<>"" then
							response.write "<a class='btn_submit' href='"&rsp(url)&"'>"&rsp(title)&"</a>"
						else
							response.write "<a class='btn_submit' href='page.asp?id="&id&"'>"&rsp(title)&"</a>"
						end if
						response.write "</div>"
					next
				next

				rsp.close
				set rsp=Nothing
			%>

		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

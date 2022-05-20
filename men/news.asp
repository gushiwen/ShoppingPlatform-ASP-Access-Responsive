<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 新闻详细
NewsId=Request("NewsId")
If not isNumeric(NewsId) Then
	conn.close
	set conn=nothing
	response.redirect "news_home.asp"
	response.end
end If

set rs=Server.CreateObject("ADODB.RecordSet")
sql= "select * from News where online=true and enNewsTitle<>'' and NewsId="&request("NewsId")
rs.open sql,conn,1,1
if rs.bof and rs.eof Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.redirect "news_home.asp"
	response.end
Else
	Title= rs("enNewsTitle")
	NewsContain= rs("enNewsContain")
	NewsClass= rs("enNewsClass")
	NewsSource= rs("enSource")
	PubDate= rs("PubDate")
	Cktimes= rs("cktimes")
	conn.execute "UPDATE news SET ckTimes ="&Cktimes+1&" WHERE NewsId="&NewsId
	rs.close
	set rs=Nothing
end If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title><%=Title%>-<%=ensitename%>-<%=SiteUrl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=Title%>,<%=ensitename%>">
		<meta name="description" content="<%=Title%>,<%=ensitename%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/newshomestyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="top-box">
		</div>

		<div class="list-box">
			<div class="list_title"><%=Title%></div> 
			<div class="list_row">
				<p>Source: <%= NewsSource%></p>
			</div> 
			<div class="list_row">
				<p>Published: <%= DateValue(PubDate)%></p>
			</div> 
			<div class="list_row">
				<div class="text_common">
					<%= NewsContain%>
				</div> 
			</div> 
			<div class="list_row">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="location.href='javascript:history.go(-1)';">
			</div> 
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>
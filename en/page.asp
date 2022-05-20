<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
id=request("id")
If not isNumeric(id) Then
	conn.close
	set conn=nothing
	response.redirect "main.html"
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
	response.redirect "main.html"
	response.end
End If 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=title%>-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>

</head>
<body>
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; <%=title%>
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

</DIV>
<!--左边结束-->


<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:10px;padding-left:10px;padding-right:10px;" >
<%= body%>
</DIV>
</DIV>
<!--中间右边一起结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>
<%
conn.close
set conn=nothing
%>

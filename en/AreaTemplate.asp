<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=FullSiteUrl%>"+"/men/main.asp");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=FullSiteUrl%>/men/main.asp" />
<link href="<%=FullSiteUrl%>/men/main.asp" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>{AreaName} <%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="{AreaName} <%=ensitekeywords%>">
<meta name="description" content="{AreaName} <%=ensitedescription%>">

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

<!--左边开始-->
<div class="homepage_left1">
<!--#include file="goldweb_treeindex.asp"-->
</div>
<!--左边结束-->

<!--右边开始-->
<div class="homepage_center">

<!--右边正中间开始-->
<div class="homepage_center1 mt6">
<!--#include file="indexgundong.asp"-->
<!--#include file="indexshop.asp"-->
</div>
<!--右边正中间结束-->

<!--小右边开始-->
<div class="homepage_right1">
<!--#include file="indexuser.asp"-->
</div>
<!--小右边结束-->

</div>
<!--右边结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>
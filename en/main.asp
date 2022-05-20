<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />
<link href="../style/camera.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../js/common.js"></script>

</head>
<body>
<!--页面开始-->
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--BODY START-->
<div class="body">

<!--LEFT START-->
<div class="homepage_left1">
<!--#include file="goldweb_treeindex.asp"-->
</div>
<!--LEFT END-->

<!--RIGHT START-->
<div class="homepage_center">

<!--RIGHT MIDDLE START-->
<div class="homepage_center1 mt6">
<!--#include file="indexgundong.asp"-->
<!--#include file="indexshop.asp"-->
</div>
<!--RIGHT MIDDLE END-->

<!--RIGHT RIGHT START-->
<div class="homepage_right1">
<!--#include file="indexuser.asp"-->
</div>
<!--RIGHT RIGHT END-->

</div>
<!--RIGHT END-->

<div class="clear"></div>

<div class="clear mt6">
<a href="<%= iadvurl4%>" target="_blank"><img src="<%= iadv4%>" width="980" height="60"  border="0" style="cursor:pointer" /></a>
</div>

</div>
<!--BODY END-->

<!--#include file="goldweb_down.asp"-->
<!--#include file="goldweb_guanggao.asp"-->
</div>
<!--页面结束-->

<!-- JavaScript libs are placed at the end of the document so the pages load faster -->
<script type="text/javascript" src="../js/jquery.min.js"></script>
<script type="text/javascript" src="../js/jquery.easing.1.3.js"></script> 
<script type="text/javascript" src="../js/camera.min.js"></script> 
<script>
		jQuery(function(){
			
			jQuery('#camera_wrap_4').camera({
				height: '195',
				loader: 'bar',
				pagination: false,
				thumbnails: false,
				hover: false,
				opacityOnGrid: false,
				imagePath: '../images/adv/'
			});

		});
</script>

</body>
</html>
<%
conn.close
set conn=nothing
%>

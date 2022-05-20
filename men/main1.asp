<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=FullSiteUrl%>"+"/en/main.html");</script>
		<link href="<%=FullSiteUrl%>/en/main.html" rel="canonical" />

		<title><%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>" />
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/mainstyle.css" rel="stylesheet" type="text/css" />
		<script src="../AmazeUI-2.4.2/assets/js/jquery.min.js" type="text/javascript"></script>
		<script src="../AmazeUI-2.4.2/assets/js/amazeui.min.js" type="text/javascript"></script>
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->
		
		<!--轮播 -->
		<%
			Set ggrs = conn.Execute("select * from adv") 
			pic1= ggrs("pic1")
			pic1_lnk= ggrs("pic1_lnk")
			tit1= ggrs("tit1")
			pic2= ggrs("pic2")
			pic2_lnk= ggrs("pic2_lnk")
			tit2= ggrs("tit2")
			pic3= ggrs("pic3")
			pic3_lnk= ggrs("pic3_lnk")
			tit3= ggrs("tit3")
			pic4= ggrs("pic4")
			pic4_lnk=ggrs("pic4_lnk")
			tit4= ggrs("tit4")
			ggrs.close
			set ggrs=nothing
		%>
		<div>
			<div class="am-slider am-slider-default scoll" data-am-flexslider id="demo-slider-0">
			<ul class="am-slides">
				<li ><a href="<%=pic1_lnk%>"><img src="<%=pic1%>" /></a></li>
				<li ><a href="<%=pic2_lnk%>"><img src="<%=pic2%>" /></a></li>
				<li ><a href="<%=pic3_lnk%>"><img src="<%=pic3%>" /></a></li>
				<li ><a href="<%=pic4_lnk%>"><img src="<%=pic4%>" /></a></li>
			</ul>
			</div>
			<div class="clear"></div>	
		</div>
		
		<!--小导航 -->
		<div class="shopNav">
				<div class="am-g am-g-fixed smallnav">
					<div class="am-u-sm-3">
						<a href="company.asp"><img src="../mimages/companysmall.jpg" />
							<div class="title">Company</div>
						</a>
					</div>
					<div class="am-u-sm-3">
						<a href="news_home.asp"><img src="../mimages/newssmall.jpg" />
							<div class="title">Newsletter</div>
						</a>
					</div>
					<div class="am-u-sm-3">
						<a href="http://www.forum.ol.sg" target=_blank><img src="../mimages/forumsmall.jpg" />
							<div class="title">Forum</div>
						</a>
					</div>
					<div class="am-u-sm-3">
						<a href="../mch/main.asp"><img src="../mimages/navsmall.jpg" />
							<div class="title">中文版</div>
						</a>
					</div>
				</div>
				<div class="clear"></div>
		</div>

		<!--推荐产品-->
		<div class="shopMain">
			<%
				Set rs=Server.CreateObject("ADODB.Recordset")
				sql="select * from goldweb_class where MidSeq=1 order by larseq"
				rs.open sql,conn,1,3
				do while not rs.eof
					larcode=rs("enLarCode")
						If rs("LarSeq") Mod 5 <> 0 Then
							classseqid = rs("LarSeq") Mod 5
						Else
							classseqid = 5
						End If 
			%>
				<div class="shopTitle ">
					<a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>"><%=rs("enLarCode")%></a>
				</div>
				<div class="am-g am-g-fixed flood method3 ">
					<ul class="am-thumbnails ">
					<%
						Set rsm=Server.CreateObject("ADODB.Recordset")
						sqlm="select top 4 * from goldweb_product where enlarcode='"&rs("enLarCode")&"' and online=true and Remark='1' order by TJDate desc"
						rsm.open sqlm,conn,1,3
						do while not rsm.eof
					%>
						<li>
							<a href="product.asp?ProdId=<%=rsm("ProdId")%>" title="<%=rsm("enProdName")%>">
								<div class="pro-image-frame "><div class="pro-image "><img src="<%=rsm("ImgPrev")%>" /></div></div>
								<div class="pro-title "><%=rsm("enProdName")%></div>
							</a>
						</li>
					<%
							rsm.movenext
						loop
						rsm.close
						set rsm=Nothing
					%>
					</ul>
				</div>
				<div class="clear "></div>
			<%
					rs.movenext
				loop
				rs.close
				set rs=Nothing
			%>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

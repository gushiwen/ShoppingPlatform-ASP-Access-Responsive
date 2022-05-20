<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
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
		<link href="../mstyle/categorystyle.css" rel="stylesheet" type="text/css" />
		<script src="../AmazeUI-2.4.2/assets/js/jquery.min.js" type="text/javascript"></script>
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<!--主体-->
		<div id="nav" class="navfull">
			<div class="area clearfix">
				<div class="category-content" id="guide_2">
					<div class="long-title"><span class="all-goods">Category</span><span id="meauBack"><a href="javascript:history.back()">Back</a></span></div>
					<div class="category">
										<ul class="category-list">
										<%
											larnumber=0
											sql="select * from goldweb_class where MidSeq=1 order by larseq"
   											Set rs=Server.CreateObject("ADODB.Recordset")
    										rs.open sql,conn,1,3
    										do while not rs.eof
												larnumber=larnumber+1
										%>
											<li class="<%If larnumber=1 Then response.write "selected" %>">
												<div class="category-info">
													<h3 class="category-name b-category-name"><a><%=rs("enLarCode")%></a></h3>
												</div>
												<div class="menu-item menu-in top">
													<div class="area-in">
														<div class="area-bg">
															<div class="menu-srot">
																<div class="sort-side">
																	<ul class="ul-sort">
																		<li class="li-sort"><a class="larcode" title="<%=rs("enLarCode")%>" href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>"><span><%=rs("enLarCode")%></span></a></dt>
																		<%
  																			Set rsm=Server.CreateObject("ADODB.Recordset")
  																			sqlm="select * from goldweb_class where enlarcode='"&rs("enLarCode")&"' and midseq<>1 order by midseq"
  																			rsm.open sqlm,conn,1,3
  																			do while not rsm.eof
																		%>
																		<li class="li-sort"><a class="midcode" title="<%=rsm("enMidCode")%>" href="productlist.asp?LarCode=<%=Server.URLEncode(rsm("enLarCode"))%>&MidCode=<%=Server.URLEncode(rsm("enMidCode"))%>&ClassId=<%=rsm("ClassId")%>"><span><%=rsm("enMidCode")%></span></a></dd>
																		<%
  																				rsm.movenext
  																			loop
  																			rsm.close
 																			set rsm=Nothing
																		%>    
																	</ul>
																</div>
															</div>
														</div>
													</div>
												</div>
											<b class="arrow"></b>	
											</li>
										<%
    											rs.movenext
    										loop
    										rs.close
    										set rs=Nothing
										%>
										</ul>
					</div>
				</div>

			</div>
		</div>
		<script type="text/javascript">
			$(document).ready(function() {
		$("li").click(function() {		
		     	$(this).addClass("selected").siblings().removeClass("selected");
	       })
		})
		</script>
		<div class="clear"></div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

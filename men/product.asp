<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
id=trim(request("prodid"))
if id="" then response.redirect "main.asp"
Set rsprod=conn.execute ("select * from goldweb_product where Online=true and ProdId='"&id&"'")
if (rsprod.bof and rsprod.eof) then
response.redirect "main.asp"
Else
Model=rsprod("enModel")
ProdName=rsprod("enProdName")	
LarCode=rsprod("enLarCode")	
MidCode=rsprod("enMidCode")	
ProdId=rsprod("ProdId")
ImgPrev=rsprod("ImgPrev")
PriceOrigin=comma(rsprod("enPriceOrigin"))
PriceList=comma(rsprod("enPriceList"))
ClickTimes=rsprod("ClickTimes")
conn.execute "UPDATE goldweb_product SET ClickTimes ="&ClickTimes+1&" WHERE ProdId ='"&id&"'"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title><%=ProdName%>-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=Prodname%>,<%=MidCode%>,<%=LarCode%>,<%=ensitename%>">
		<meta name="description" content="<%=Prodname%>,<%=ensitename%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/productstyle.css" type="text/css" rel="stylesheet" />
		<script src="../mjs/jquery-1.7.min.js" type="text/javascript"></script>
		<script src="../mjs/jquery.flexslider.js" type="text/javascript"></script>
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="productMain">

			<script type="text/javascript">
				$(function() {});
				$(window).load(function() {
					$('.flexslider').flexslider({
						animation: "slide",
						start: function(slider) {
							$('body').removeClass('loading');
						}
					});
				});
			</script>
			<div class="scoll">
				<section class="slider">
					<div class="flexslider">
						<ul class="slides">
						<%
							Set more_pic=Server.CreateObject("ADODB.Recordset")
							sqlmp="select A.ID,A.ProdId,A.FilePath,B.ProdName from more_pic A,goldweb_product B where A.ProdId=B.ProdId and A.ProdId='"&id&"' order by A.Num asc"
							more_pic.Open sqlmp,conn,1,1

							response.write "<li><img src="&rsprod("ImgPrev")&" /></li>"

							If Not(more_pic.bof and more_pic.eof) then
								do while not more_pic.eof
									response.write "<li><img src="&more_pic("FilePath")&" /></li>"
									more_pic.movenext
								Loop
							End If 
							more_pic.close
							set more_pic=Nothing
						%>
						</ul>
					</div>
				</section>
			</div>

			<!--规格属性-->
			<!--名称-->
			<div class="tb-detail-hd">
				<h1><%= ProdName%></h1>
			</div>

			<!--价格-->
			<% 
				If CSng(rsprod("enPriceList")) > 0 Or CSng(rsprod("enPriceOrigin")) >0 Then 
					response.write "<div class=""tb-detail-price"">"
					If CSng(rsprod("enPriceList")) > 0 then response.write "<li class=""price iteminfo_price""><dd><em>"&rsprod("enPriceUnit")&"</em><b class=""sys_item_price"">"&comma(rsprod("enPriceList"))&"</b></dd></li>"     
					If CSng(rsprod("enPriceOrigin")) >0 Then response.write "<li class=""price iteminfo_mktprice""><dd><em>"&rsprod("enPriceUnit")&"</em><b class=""sys_item_mktprice"">"&comma(rsprod("enPriceOrigin"))&"</b></dd></li>"
					response.write "<div class=""clear""></div>"
					response.write "</div>"
				End If 
			%>

			<!--收藏、购物车按钮-->
			<div class="pay">
				<li>
					<div class="clearfix tb-btn tb-btn-buy theme-login">
						<a href="my_fav.asp?ProdId=<%=rsprod("ProdId")%>">Add to Favorites</a>
					</div>
				</li>
				<li>
					<div class="clearfix tb-btn tb-btn-basket theme-login">
						<script type="text/javascript">
							if(getCookie('cart').indexOf('<%=rsprod("ProdId")%>')>-1)
							{
								document.write('<a href=\'shoppingcart.asp\'>In Shopping Cart</a>');
							}
							else
							{
								if('<%=rsprod("AddtoCart")%>'=='1')
								{
									if(<%=reg%>==1)
									{
										if(getCookie('userid')=='')
										{
											document.write('<a href=\'login.asp\'>Add to Cart</a>');
										}
										else
										{
											document.write('<a href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\'>Add to Cart</a>');
										}
									}
									else
									{
										document.write('<a href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\'>Add to Cart</a>');
									}
								}
							}
						</script>
					</div>
				</li>
				<div class="clear"></div>
			</div>

			<%If Trim(rsprod("enProdDisc"))<>"" Then %>
			<div class="content-box">
				<div class="content-title">Introduction</div>
				<div class="content-detail">
					<%=rsprod("enProdDisc")%>
				</div>
				<div class="clear"></div>
			</div>
			<%End If %>

			<%If Trim(rsprod("enMemoSpec"))<>"" Then %>
			<div class="content-box">
				<div class="content-title">Specs and Details</div>
				<div class="content-detail">
					<%=rsprod("enMemoSpec")%>
				</div>
				<div class="clear"></div>
			</div>
			<%End If %>

			<div class="btn_box">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="location.href='javascript:history.go(-1)';">
			</div>
		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>
<%
conn.close
set conn=Nothing
%>

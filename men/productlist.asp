<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
Page=request("page")
If not isNumeric(page) then
response.redirect "main.asp"
response.end
end if
pagecount=page

LarCode=trim(request("LarCode"))
MidCode=Trim(request("MidCode"))
ClassId=Trim(request("ClassId"))
action=trim(request("action"))
listorder=trim(request("listorder"))

sqlprod="select A.ProdId,A.enProdName,A.ImgPrev,A.enPriceUnit,A.enPriceOrigin,A.enPriceList,A.AddtoCart from goldweb_product A left join goldweb_class B on A.enLarCode=B.enLarCode and A.enMidCode=B.enMidCode where A.online=true "

if LarCode<>"" then 
title = LarCode & "-" & title
sqlprod = sqlprod & "and  A.enLarCode='"&LarCode&"'"
End If

if MidCode<>"" then 
title = MidCode & "-" & title
sqlprod = sqlprod & "and  A.enMidCode='"&MidCode&"'"
End if

if action="tuijian" then 
title="Recommend Products" & "-" & title
sqlprod = sqlprod & " and  A.Remark='1' order by  A.Tjdate desc"
elseif action="new" then 
title="New Products" & "-" & title
sqlprod = sqlprod & " order by  A.AddDate desc"
elseif action="tejia" then 
title="Special Products" & "-" & title
sqlprod = sqlprod & " and A.tejia='1' order by  A.AddDate desc"
elseif action="hot" then
title="Hot Products" & "-" & title
sqlprod = sqlprod & " order by  A.ClickTimes desc"
else
sqlprod = sqlprod & "order by A.AddDate desc"
'sqlprod = sqlprod & "order by B.LarSeq, B.MidSeq, A.enModel,  A.AddDate desc"
end If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title><%=title%><%=ensitename%>-<%=SiteUrl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=LarCode%>,<%=MidCode%>,<%=ensitename%>">
		<meta name="description" content="<%=LarCode%>,<%=MidCode%>,<%=ensitename%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/productliststyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/jquery-1.7.min.js" type="text/javascript"></script>
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="listMain">
			
				<div class="am-g am-g-fixed">
					<div class="am-u-sm-12 am-u-md-12">

						<div class="theme-popover">
								<p>
								<% 
									If Trim(LarCode)<>"" Then response.write LarCode
									If Trim(MidCode)<>"" Then response.write " &gt; "&MidCode
								%>
								</p>
                        </div>

						<%
							Set rsprod=Server.CreateObject("ADODB.RecordSet") 
							rsprod.open sqlprod,conn,1,1
							n=0
							if rsprod.bof and rsprod.eof then
								response.write "<div class='theme-popover'>Coming soon...</div>"
							else
								pages=per_page_num
								rsprod.pagesize=per_page_num
								allPages = rsprod.pageCount	'总页数
								If not isNumeric(page) then page=1
								if isEmpty(page) or clng(page) < 1 then
									page = 1
								elseif clng(page) >= allPages then
									page = allPages 
								end if
								rsprod.AbsolutePage=page
						%> 
						<div class="product-content">
							<ul class="am-avg-sm-2 am-avg-md-3 am-avg-lg-4 boxes">
							<%
								Do While Not rsprod.eof And pages>0
							%> 
								<li>
									<div class="i-pic limit">
										<a href="product.asp?ProdId=<%=rsprod("ProdId")%>" title="<%=rsprod("enProdName")%>">
											<div class="pro-image-frame"><div class="pro-image"><img src="<%=rsprod("ImgPrev")%>" /></div></div>
											<div class="pro-title"><%=rsprod("enProdName")%></div>
										</a>
										<p class="price fl">
										<% 
											If CSng(rsprod("enPriceList")) > 0 Then 
										%>
												<b><%=rsprod("enPriceUnit")%></b><strong><%=comma(rsprod("enPriceList"))%></strong>

						<script type="text/javascript">
							if(getCookie('cart').indexOf('<%=rsprod("ProdId")%>')>-1)
							{
								document.write('<input class="cartbtn" type="image" src="../mimages/incart.png"  onclick="javascript:location.href=\'shoppingcart.asp\';">');
							}
							else
							{
								if('<%=rsprod("AddtoCart")%>'=='1')
								{
									if(<%=reg%>==1)
									{
										if(getCookie('userid')=='')
										{
											document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'login.asp\';">');
										}
										else
										{
											document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';">');
										}
									}
									else
									{
										document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';">');
									}
								}
							}
						</script>
										<%
											End If 
										%>
											<input class="cartbtn"  type="image" src="../mimages/addtofavorites.png"  onclick="javascript:location.href='my_fav.asp?ProdId=<%=rsprod("ProdId")%>';">
										</p>
									</div>
								</li>
							<%
								pages = pages - 1
								rsprod.movenext
								if rsprod.eof then exit do
								Loop
							%>
							</ul>

							<!--分页 -->
							<ul class="am-pagination am-pagination-centered">
							<%
								if page = 1 then
									response.write "<li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;First&nbsp;&nbsp;&nbsp;</a></li><li class=""am-disabled""><a href=""#"">Previous</a></li>"
								else
									response.write "<li class=""am-active""><a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page=1>&nbsp;&nbsp;&nbsp;First&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&page-1&">Previous</a></li>"
								end if
								if page = allpages then
									response.write "<li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
								else
									response.write "<li class=""am-active""><a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&page+1&">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&allpages&">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
								end if
							%>
								<p>Total <%=RSprod.RecordCount%> Products, Page <%=page%>/<%=allpages%></p>
							</ul>
						</div>
						<div class="clear"></div>
						<%
							end if
							rsprod.close
							set rsprod=nothing
						%>

					</div>
				</div>

		</div>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

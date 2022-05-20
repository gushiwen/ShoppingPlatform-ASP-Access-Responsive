<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
Page=request("page")
action=trim(request("action"))
keywords=trim(request("keywords"))

if action="topsearch" Then 'From top search
	if keywords="" Then
		conn.close
		set conn=Nothing
		response.write "<script language='javascript'>"
		response.write "location.href='javascript:history.go(-1)';"							
		response.write "</script>"	
		response.end
	else
		Set rs=Server.CreateObject("ADODB.RecordSet") 
		sql="select * from goldweb_key where keywords='"&Replace(keywords,"'","''")&"'"
		rs.open sql,conn,1,3
		if not (rs.eof and rs.bof) then
			rs("num")=rs("num")+1
			rs("keydate")=now()
		else
			rs.addnew
			rs("keywords")=keywords
			rs("keydate")=now()
		end if
		rs.update
		rs.close
		set rs=Nothing

		sqlNameKeys=""
		sqlDiscKeys=""
		SpKeywords = Split(Replace(keywords,"'","''"), " ")
		For I=0 To UBound(SpKeywords)
			If SpKeywords(I) <> "" then
				if sqlNameKeys="" then
					sqlNameKeys="A.enProdName like '%" & SpKeywords(I) & "%'"
					sqlDiscKeys="A.enProdDisc like '%" & SpKeywords(I) & "%'"
				Else
					sqlNameKeys=sqlNameKeys & " and A.enProdName like '%" & SpKeywords(I) & "%'"
					sqlDiscKeys=sqlDiscKeys & " and A.enProdDisc like '%" & SpKeywords(I) & "%'"
				End If
			End If 
		Next
		If sqlNameKeys<>"" Then 
		sqlNameKeys="(" & sqlNameKeys & ")"
		sqlDiscKeys="(" & sqlDiscKeys & ")"
		sqlprod="select A.ProdId,A.enProdName,A.ImgPrev,A.enPriceUnit,A.enPriceOrigin,A.enPriceList,A.AddtoCart from goldweb_product A left join goldweb_class B on A.enLarCode=B.enLarCode and A.enMidCode=B.enMidCode where A.online=true and (" & sqlNameKeys & " or " & sqlDiscKeys & ") order by A.AddDate desc"
		'sqlprod="select A.ProdId,A.enProdName,A.ImgPrev,A.enPriceUnit,A.enPriceOrigin,A.enPriceList,A.AddtoCart from goldweb_product A left join goldweb_class B on A.enLarCode=B.enLarCode and A.enMidCode=B.enMidCode where A.online=true and (" & sqlNameKeys & " or " & sqlDiscKeys & ") order by B.LarSeq, B.MidSeq, A.enModel,  A.AddDate desc"
		End If 
		title="Search results for <font color=red>"&Keywords&"</font> "
	end If

Else 'From advanced search
		conn.close
		set conn=Nothing
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

		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Product Search-<%=enSiteName%>-<%=SiteUrl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="description" content="<%=ensitedescription%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
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
									<% if title<>"" Then response.write title %>
								</p>
                        </div>

						<%
							Set rsprod=Server.CreateObject("ADODB.RecordSet") 
							rsprod.open sqlprod,conn,1,1
							n=0
							if rsprod.bof and rsprod.eof then
								response.write "<div class='theme-popover'>We can not find the products.</div>"
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
								Do While Not rsprod.eof and pages>0
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
									response.write "<li class=""am-active""><a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page=1>&nbsp;&nbsp;&nbsp;First&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page-1&">Previous</a></li>"
								end if
								if page = allpages then
									response.write "<li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
								else
									response.write "<li class=""am-active""><a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page+1&">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&allpages&">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
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

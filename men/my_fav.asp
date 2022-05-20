<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
If request.cookies("goldweb")("userid")="" Then 
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please log in first.');"
	response.write "location.href='login.asp';"
	response.write "</script>"
	response.end
End If 

Set rs = conn.Execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'") 
if not (rs.eof and rs.bof) then fav=rs("fav")
if isnull(fav)=true then fav=""
rs.close
set rs=Nothing

' 添加、编辑 favorite products 
if request("edit")="ok" then fav=""
buyid = Split(request("prodid"), ", ")
For I=0 To UBound(buyid)
	if fav="" then
		fav = "'"&buyid(I)&"'"
	ElseIf InStr(fav, buyid(i)) <= 0 Then
		fav = fav & ", '" & buyid(i) &"'"
	End If
Next

if len(fav)>=200 then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('You choose too many favorite products.');"
	response.write "location.href='javascript:history.go(-1)';"		
	response.write "</script>"
	response.end
end If

Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
rs.open sql,conn,3,3
rs("fav")=fav
rs.update
rs.close
set rs=Nothing

' 添加 favorite products 后返回产品页面
if request("prodid")<>"" And request("edit")<>"ok" then 
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "location.href=document.referrer;"							
	response.write "</script>"	
	response.end
End If 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>Favorite Products-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

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
								<p>Favorite Products</p>
                        </div>

						<%
							Set rsprod = conn.Execute("select * from goldweb_product where ProdId in ("&fav&") order by instr ('"&Replace(fav,"'","")&"',ProdId)") 
							n=0
							if rsprod.bof and rsprod.eof then
								response.write "<div class='theme-popover'>Please choose favorite products!</div>"
							else
						%> 
						<form action="my_fav.asp" method="POST" name="check">
						<div class="product-content">
							<ul class="am-avg-sm-2 am-avg-md-3 am-avg-lg-4 boxes">
							<%
								Do While Not rsprod.eof 
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
								document.write('<input class="cartbtn" type="image" src="../mimages/incart.png"  onclick="javascript:location.href=\'shoppingcart.asp\';return false;">');
							}
							else
							{
								if('<%=rsprod("AddtoCart")%>'=='1')
								{
									if(<%=reg%>==1)
									{
										if(getCookie('userid')=='')
										{
											document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'login.asp\';return false;">');
										}
										else
										{
											document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';return false;">');
										}
									}
									else
									{
										document.write('<input class="cartbtn" type="image" src="../mimages/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';return false;">');
									}
								}
							}
						</script>
										<%
											End If 
										%>
											<input class="cartbtn" type="CheckBox" name="ProdId" value="<%=rsprod("ProdId")%>" Checked>
										</p>
									</div>
								</li>
							<%
									rsprod.movenext
								Loop
							%>
							</ul>
						</div>

						<div class="btn_box">
							<input class="btn_submit" type="submit" value="Update Select Box">
						</div>
						<input type="hidden" name="edit" value="ok">
						</form>
						<%
								rsprod.close
								set rsprod=nothing
							end if
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

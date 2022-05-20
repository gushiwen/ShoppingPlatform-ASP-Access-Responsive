<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
Page=request("page")
action=trim(request("action"))
keywords=trim(request("keywords"))

if action="topsearch" Then 'From top search
	'Response.cookies("goldweb").path="/"
	'response.cookies("goldweb")("search")=""

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
		sqlprod="select A.ProdId,A.enProdName,A.ImgPrev,A.enPriceUnit,A.enPriceOrigin,A.enPriceList,A.AddtoCart from goldweb_product A left join goldweb_class B on A.enLarCode=B.enLarCode and A.enMidCode=B.enMidCode where A.online=true and (A.ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=A.ExpiryDate and DateValue(A.AddDate)<>A.ExpiryDate)) and (" & sqlNameKeys & " or " & sqlDiscKeys & ") order by A.AddDate desc"
		End If 
		title="Search results for <font color=red>"&Keywords&"</font> "
	end If

Else 'From advanced search
		conn.close
		set conn=Nothing
		response.redirect "main.html"
		response.end
	'If request.cookies("goldweb")("search")=""  then 
		'conn.close
		'set conn=nothing
		'response.redirect "search_advanced.asp"
	'Else
		'sqlprod=request.cookies("goldweb")("search")
		'title="Advanced Search"
	'End If
End If 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Product Search-<%=enSiteName%>-<%=SiteUrl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="description" content="<%=ensitedescription%>">
<meta name="keywords" content="<%=ensitekeywords%>">

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
<DIV class="nav blueT">Current: <%=enSiteName%> 
<%
  if title<>"" then
	response.write "&gt; " & title
  end If
%>
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
<DIV class=mt6 id=homepage_center_new_new>

<%
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1
n=0
if rsprod.bof and rsprod.eof then
response.write "No products!"
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

<TABLE cellSpacing=0 cellPadding=0 width="100%" align="center" border=0>
  <DIV class=Select>
  <TBODY>
  <TR>
    <TD class="list_title">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TR>
          <TD align="center">
<!--翻页控制开始--> 
<form action="search.asp" method="post" >
Total <font color=red><%=RSprod.RecordCount%></font> products, Page <%=page%>/<%=allpages%>, &nbsp;&nbsp;
  <%
if page = 1 then
response.write "<font color=darkgray>First Previous</font>"
else
response.write "<a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page=1>First</a> <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page-1&">Previous</a>"
end if
if page = allpages then
response.write "<font color=darkgray> Next Last</font>"
else
response.write " <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page+1&">Next</a> <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&allpages&">Last</a>"
end if
%>
 &nbsp; Goto page 
<input name="page" type="text" size="2" class="inputstyle" />
    <input type="submit" name="Submit" value="GO" style="font-family: Arial; font-size: 10px"><input type="hidden" name="action" value="<%=action%>"><input type="hidden" name="Keywords" value="<%=Keywords%>">
</form>
<!--翻页控制结束-->
          </TD>
		</TR>
	  </TABLE>
	</TD>
  </TR>

  <TR>
    <TD vAlign=top align=middle>
      <DIV class="list-content grid">
      <UL class="product-list ark:list">

<%
Do While Not rsprod.eof	and pages>0
%> 
        <LI class="product">
        <H3><%=rsprod("enProdName")%></H3>
        <DIV class="info">
          <DIV class="pic">
		    <A href="product.asp?ProdId=<%=rsprod("ProdId")%>"><IMG title="<%=rsprod("enProdName")%>" src="<%=rsprod("ImgPrev")%>" width="160" height="160" onload="DrawImage(this,160,160)" border="0" /> </A>
		  </DIV>
          <DIV style="PADDING-LEFT: 10px">
			<% 
				If CSng(rsprod("enPriceList")) > 0 Or CSng(rsprod("enPriceOrigin")) >0 Then 
					response.write "<DIV class=""price"">"
					If CSng(rsprod("enPriceList")) > 0 then response.write "<span class=""red14"">"&rsprod("enPriceUnit")&comma(rsprod("enPriceList"))&"</span> "
					If CSng(rsprod("enPriceOrigin")) >0 Then response.write "<span class=""listprice gray"">"&rsprod("enPriceUnit")&comma(rsprod("enPriceOrigin"))&"</span>"
					response.write "</DIV>"
				End If 
			%>
            
            <DIV class="name">
			  <A href="product.asp?ProdId=<%=rsprod("ProdId")%>"><%=rsprod("enProdName")%></A>
			</DIV>

            <DIV class="pt5">
			  <SPAN style="MARGIN-TOP:95px;">

	<script type="text/javascript">
	if(getCookie('cart').indexOf('<%=rsprod("ProdId")%>')>-1)
	{
		document.write('<IMG style="CURSOR: pointer" src="../images/shoppingcart.png" align="bottom" border=0 onclick="javascript:location.href=\'shoppingcart.asp\';" /> <a href=\'shoppingcart.asp\'>In Shopping Cart</a>');
	}
	else
	{
		if('<%=rsprod("AddtoCart")%>'=='1')
		{
			if(<%=reg%>==1)
			{
				if(getCookie('userid')=='')
				{
					document.write('<IMG style="CURSOR: pointer" onclick="javascript:location.href=\'login.asp\';" src="images/addtocart.gif" align="bottom" border=0>');
				}
				else
				{
					document.write('<IMG style="CURSOR: pointer" onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';" src="images/addtocart.gif" align="bottom" border=0>');
				}
			}
			else
			{
				document.write('<IMG style="CURSOR: pointer" onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';" src="images/addtocart.gif" align="bottom" border=0>');
			}
		}
		else
		{
			document.write('<IMG src="images/addtocart_gray.gif" align="bottom" border=0>');
		}
	}
	</script>

				  &nbsp;&nbsp;<IMG style="CURSOR: pointer" onclick="javascript:location.href='my_fav.asp?ProdId=<%=rsprod("ProdId")%>';" src="../images/addtofavorites.png" align="bottom" border=0>
			  </SPAN>
			</DIV>
		  </DIV>
		</DIV>
        </LI>
<%
pages = pages - 1
rsprod.movenext
Loop
%>

      </UL>
	</DIV>
	</TD>
  </TR>

  <TR>
    <TD vAlign="top" align="center">

<!--翻页控制开始--> 
<form action="search.asp" method="post" >
Total <font color=red><%=RSprod.RecordCount%></font> products, Page <%=page%>/<%=allpages%>, &nbsp;&nbsp;
  <%
if page = 1 then
response.write "<font color=darkgray>First Previous</font>"
else
response.write "<a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page=1>First</a> <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page-1&">Previous</a>"
end if
if page = allpages then
response.write "<font color=darkgray> Next Last</font>"
else
response.write " <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&page+1&">Next</a> <a href=search.asp?action="&action&"&Keywords="&server.URLEncode(Keywords)&"&page="&allpages&">Last</a>"
end if
%>
 &nbsp; Goto page 
<input name="page" type="text" size="2" class="inputstyle" />
    <input type="submit" name="Submit" value="GO" style="font-family: Arial; font-size: 10px"><input type="hidden" name="action" value="<%=action%>"><input type="hidden" name="Keywords" value="<%=Keywords%>">
</form>
<!--翻页控制结束-->

    </TD>
  </TR>
  </TBODY>
  </DIV>
</TABLE>
<%
rsprod.close
set rsprod=nothing
end if
%>

</div>
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
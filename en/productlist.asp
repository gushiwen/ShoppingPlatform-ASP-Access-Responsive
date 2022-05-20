<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
Page=request("page")
If not isNumeric(page) then
response.redirect "main.html"
response.end
end if
pagecount=page

LarCode=trim(request("LarCode"))
MidCode=Trim(request("MidCode"))
ClassId=Trim(request("ClassId"))
action=trim(request("action"))
listorder=trim(request("listorder"))
'if listorder=2 Then

'if action="hot" and page<=1 then
'conn.execute ("update goldweb_product set ClickTimes=ClickTimes-1 where ClickTimes="&conn.execute ("select max(ClickTimes) as maxc from goldweb_product where online=true")("maxc"))
'end if

'sqltree="select * from goldweb_class "
sqlprod="select A.ProdId,A.enProdName,A.ImgPrev,A.enPriceUnit,A.enPriceOrigin,A.enPriceList,A.AddtoCart from goldweb_product A left join goldweb_class B on A.enLarCode=B.enLarCode and A.enMidCode=B.enMidCode where A.online=true and (A.ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=A.ExpiryDate and DateValue(A.AddDate)<>A.ExpiryDate)) "

if LarCode<>"" then 
title = LarCode & "-" & title
'sqltree = sqltree & "and enLarCode='"&LarCode&"' and midseq<>1"
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
end If

'elseif listorder=1 Then

'if action="hot" and page<=1 then
'conn.execute ("update goldweb_product set ClickTimes=ClickTimes-1 where ClickTimes="&conn.execute ("select max(ClickTimes) as maxc from goldweb_product where online=true")("maxc"))
'end If

'end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=title%><%=ensitename%>-<%=SiteUrl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=LarCode%>,<%=MidCode%>,<%=ensitename%>">
<meta name="description" content="<%=LarCode%>,<%=MidCode%>,<%=ensitename%>">

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
<DIV class="nav blueT">Current: <%=ensitename%> &gt; 
Category
<% 
If Trim(LarCode)<>"" Then response.write " &gt; "&LarCode
If Trim(MidCode)<>"" Then response.write " &gt; "&MidCode
%>
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

<!--特价商品开始-->
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  if LarCode="" then
    sql="select top 10 * from goldweb_product where online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) and tejia='1' order by AddDate desc"
  Else
    if MidCode="" then
      sql="select top 10 * from goldweb_product where enlarcode='"&LarCode&"' and online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) and tejia='1' order by AddDate desc"
	Else
      sql="select top 10 * from goldweb_product where enlarcode='"&LarCode&"' and enmidcode='"&MidCode&"' and online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) and tejia='1' order by AddDate desc"
	End if
  End If
  rs.open sql,conn,1,3

  If rs.bof And rs.eof Then
  Else
%>
<div class="border4 mt6">
  <div class="right_title2">
    <div class="font14" style="width:120px; float:left;"><strong>Promotion</strong></div>
  </div>
           <div class="product_line2">
           <ul>
<%
  do while not rs.eof
%> 
		      <li>
		        <div class="img2 left"> <a href="product.asp?ProdId=<%=rs("ProdId")%>" title="<%=rs("enProdName")%>" ><img src="<%= rs("ImgPrev")%>" width="48" height="48" border="0" onload='DrawImage(this,48,48)' /></a></div>
		        <div class="right line2">
				<%
					If rs("enModel") <> "" Then response.write"<a href='product.asp?ProdId="&rs("ProdId")&"' title='"&rs("enProdName")&"'>"&lleft(rs("enModel"),16)&"</a><br />"
					If CSng(rs("enPriceList")) > 0 then response.write "<span class=""red14"">"&rs("enPriceUnit")&comma(rs("enPriceList"))&"</span> "
				%>
		        </div>
		      </li>
<%
  rs.movenext
  loop
%>
           </ul>
          <div class="clear"></div>
          </div>
</div>
<%
  End If 

  rs.close
  set rs=Nothing
%>
<!--特价商品结束-->

</DIV>
<!--左边结束-->


<!--中间右边一起开始-->
<DIV class=mt6 id=homepage_center_new_new>

<%
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1
n=0
if rsprod.bof and rsprod.eof then
response.write "Coming soon..."
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

<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <DIV class=Select>
  <TBODY>
  <TR>
    <TD class=list_title>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TR>
          <TD>Order by: <A href="productlist.asp?action=hot&LarCode=<%=server.URLEncode(LarCode)%>&MidCode=<%=server.URLEncode(MidCode)%>&ClassId=<%=ClassId%>">Hot</A> 
           &nbsp; <A href="productlist.asp?action=tejia&LarCode=<%=server.URLEncode(LarCode)%>&MidCode=<%=server.URLEncode(MidCode)%>&ClassId=<%=ClassId%>">Special</A> 
           &nbsp;<A href="productlist.asp?action=tuijian&LarCode=<%=server.URLEncode(LarCode)%>&MidCode=<%=server.URLEncode(MidCode)%>&ClassId=<%=ClassId%>">Recommend</A> 
           &nbsp;<A href="productlist.asp?action=new&LarCode=<%=server.URLEncode(LarCode)%>&MidCode=<%=server.URLEncode(MidCode)%>&ClassId=<%=ClassId%>">New</A> 
           &nbsp;<!--<a href='<--@--search_url_orderby  orderby1="empty" orderby2="stock" ascflag="1" />'>库存</a> --></TD>
          <TD width=230>
		    Page <%=page%>/<%=allpages%>, Total <font color=red><%=RSprod.RecordCount%></font> products,
		  </TD>
          <TD width=211>		  

<!--翻页控制开始--> 
<form action="productlist.asp" method="post" >
 &nbsp; Goto Page 
<input name="page" type="text" size="2" class="inputstyle" />
    <input type="submit" name="Submit" value="GO" style="font-family: Arial; font-size: 10px"> <input type="hidden" name="LarCode" value="<%=LarCode%>"><input type="hidden" name="MidCode" value="<%=MidCode%>"><input type="hidden" name="action" value="<%=action%>">
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
		    <A href="product.asp?ProdId=<%=rsprod("ProdId")%>" title="<%=rsprod("enProdName")%>"><IMG src="<%=rsprod("ImgPrev")%>" width="160" height="160"  onload="DrawImage(this,160,160)" border="0" /> </A>
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
if rsprod.eof then exit do
Loop
%>

      </UL>
	</DIV>
	</TD>
  </TR>

  <TR>
    <TD vAlign="top" align="center">

<!--翻页控制开始--> 
<form action="productlist.asp" method="post" >
Total <font color=red><%=RSprod.RecordCount%></font> products, Page <%=page%>/<%=allpages%>, &nbsp;&nbsp;
  <%
if page = 1 then
response.write "<font color=darkgray>First Previous</font>"
else
response.write "<a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page=1>First</a> <a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&page-1&">Previous</a>"
end if
if page = allpages then
response.write "<font color=darkgray> Next Last</font>"
else
response.write " <a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&page+1&">Next</a> <a href=productlist.asp?action="&action&"&LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"&ClassId="&ClassId&"&page="&allpages&">Last</a>"
end if
%>
 &nbsp; Goto page 
<input name="page" type="text" size="2" class="inputstyle" />
    <input type="submit" name="Submit" value="GO" style="font-family: Arial; font-size: 10px"> <input type="hidden" name="LarCode" value="<%=LarCode%>"><input type="hidden" name="MidCode" value="<%=MidCode%>"><input type="hidden" name="action" value="<%=action%>">
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

<DIV class=clear></DIV>
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

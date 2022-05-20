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
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
End If 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Favorite Products-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

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
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; Favorite Products
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
Set rsprod = conn.Execute("select * from goldweb_product where ProdId in ("&fav&") order by instr ('"&Replace(fav,"'","")&"',ProdId)") 
n=0
if rsprod.bof and rsprod.eof then
	response.write "No favorite products!"
else

%> 
<form action="my_fav.asp" method="POST" name="check">
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <DIV class=Select>
  <TBODY>
  <TR>
    <TD class="list_title">&nbsp;</TD>
  </TR>

  <TR>
    <TD vAlign=top align=middle>
      <DIV class="list-content grid">
      <UL class="product-list ark:list">

<%
Do While Not rsprod.eof 
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

				  &nbsp;&nbsp;<input type="CheckBox" name="ProdId" value="<%=rsprod("ProdId")%>" Checked>
			  </SPAN>
			</DIV>
		  </DIV>
		</DIV>
        </LI>
<%
rsprod.movenext
Loop
%>

      </UL>
	</DIV>
	</TD>
  </TR>

  <TR height="38">
    <TD vAlign="top" align="center">
		<input type="submit" value="  Update Select Box  " style="font-size: 10px; padding:2px;">&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
		<input type="hidden" name="edit" value="ok">
    </TD>
  </TR>
  </TBODY>
  </DIV>
</TABLE>
</form>
<%
rsprod.close
set rsprod=nothing
end if
%>

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
set conn=Nothing
%>

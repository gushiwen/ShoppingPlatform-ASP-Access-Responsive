<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->
<!--#include file="comment_save.asp" -->

<%
'call aspsql()
id=trim(request("prodid"))
if id="" then response.redirect "main.html"
Set rsprod=conn.execute ("select * from goldweb_product where Online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) and ProdId='"&id&"'")
if (rsprod.bof and rsprod.eof) then
response.redirect "main.html"
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
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=ProdName%>-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=Prodname%>,<%=MidCode%>,<%=LarCode%>,<%=ensitename%>">
<meta name="description" content="<%=Prodname%>,<%=ensitename%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />
<link href="../style/product.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script type="text/javascript" src="../js/product.js"></script>
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
 &gt; <%=lleft(Prodname,62)%>
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

<!--相关商品开始-->
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  if LarCode="" then
    sql="select top 10 * from goldweb_product where online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) order by ClickTimes desc, AddDate desc"
  Else
    if MidCode="" then
      sql="select top 10 * from goldweb_product where enlarcode='"&LarCode&"' and prodid<>'"&ProdId&"' and online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) order by ClickTimes desc, AddDate desc"
	Else
      sql="select top 10 * from goldweb_product where enlarcode='"&LarCode&"' and enmidcode='"&MidCode&"' and prodid<>'"&ProdId&"' and online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) order by ClickTimes desc, AddDate desc"
	End if
  End if
  rs.open sql,conn,1,3

  If rs.bof And rs.eof Then
  Else
%>
<div class="border4 mt6">
  <div class="right_title2">
    <div class="font14" style="width:120px; float:left;"><strong>Related</strong></div>
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
<!--相关商品结束-->

</DIV>
<!--左边结束-->


<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>

<!-- 上半部分左边图片和右边简介 开始 -->
<DIV>

<!-- 右边产品简介 开始 -->
<DIV class=list_s style="FLOAT: right; WIDTH: 480px">

<DIV class=pb5>
<H1><SPAN id=product_name style="WIDTH: 80%"><%= ProdName%></SPAN></H1>
</DIV>
<DIV class="gray pb5">Model: <SPAN id=product_code><%= Model%></SPAN></DIV>

            <!-- 买家体验开始 -->
			<%
			  set rspl=conn.execute("select avg(jibie), count(jibie) from goldweb_pinglun where prodid='"&ProdId&"' and view='1'")
			  if Not (rspl.eof and rspl.bof) then
			%>
            <DIV class=pb5 id=detailRating>
			  <%
			    If rspl(0) Then
			      For i=1 To round(rspl(0))
			  %>
			  <IMG src="../images/star_red.gif" />
			  <%
			      next
			    Else
			  %>
			      <IMG src="../images/star_red.gif" />
			      <IMG src="../images/star_red.gif" />
			      <IMG src="../images/star_red.gif" />
			      <IMG src="../images/star_red.gif" />
			      <IMG src="../images/star_red.gif" />
			  <%
			    End if
			  %>
			  (<%= rspl(1)%> Comments)
			</DIV>
			<%
			  End if
			  rspl.close
			  set rspl=nothing
			%>
			<!-- 买家体验结束 -->

<% 
	If CSng(rsprod("enPriceList")) > 0 Or CSng(rsprod("enPriceOrigin")) >0 Then 
		response.write "<DIV class=""price"">"
		If CSng(rsprod("enPriceList")) > 0 then response.write "<span class=""red14"">"&rsprod("enPriceUnit")&comma(rsprod("enPriceList"))&"</span> "
		If CSng(rsprod("enPriceOrigin")) >0 Then response.write "<span class=""listprice gray"">"&rsprod("enPriceUnit")&comma(rsprod("enPriceOrigin"))&"</span>"
		response.write "</DIV>"
	End If 
%>

<DIV class="dot1 pt5 pb5" style="HEIGHT: 40px">
<!-- 立即购买 -->
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

<!-- 收藏商品 -->
&nbsp;&nbsp;<IMG style="CURSOR: pointer" src="../images/addtofavorites.png" align="bottom"  border=0 onclick="javascript:location.href='my_fav.asp?ProdId=<%=rsprod("ProdId")%>';" /> <A href="my_fav.asp?ProdId=<%=rsprod("ProdId")%>" >Add to Favorites</A> 
</DIV>

<DIV style="HEIGHT: 15px"></DIV>

<DIV class="gray pt5 mr5">
<IMG height=22 src="../images/icon_tj.gif" width=20 align=absMiddle><STRONG>Introduction</STRONG>
<br /><%=rsprod("enProdDisc")%>
</DIV>

</DIV>
<!-- 右边产品简介 结束 -->

<!-- 左边图片 开始 -->
<DIV style="BORDER-RIGHT: #dcddde 1px solid; FLOAT: left; WIDTH: 300px">

<TABLE cellSpacing=0 cellPadding=0 width=300 border=0>
  <TBODY>
  <TR>
    <TD align=middle height=300><!-- 局部放大 -->
	  <SPAN id=picDiv style="WIDTH: 300px; HEIGHT: 300px"><IMG class=jqzoom id=mainpicimg title="<%= ProdName%>" alt="<%= ProdName%>" src="<%= rsprod("ImgPrev")%>" width="300" height="300" border="0" onLoad="DrawImage(this,300,300)"></SPAN>
	  <SCRIPT type=text/javascript>
        document.getElementById('mainpicimg').style.cursor="default";
      </SCRIPT>
    </TD>
  </TR>


<%
Set more_pic=Server.CreateObject("ADODB.Recordset")
sqlmp="select A.ID,A.ProdId,A.FilePath,B.ProdName from more_pic A,goldweb_product B where A.ProdId=B.ProdId and A.ProdId='"&id&"' order by A.Num asc"
more_pic.Open sqlmp,conn,1,1

if more_pic.bof and more_pic.eof then
'response.write "没有图片或者显示图片出错！"
Else

page_pic_num=4
total_pic_num=more_pic.RecordCount+1 'ImgPrev一起显示
more_pic.pagesize=page_pic_num
allPages = more_pic.pageCount	'总页数
%>

<SCRIPT type=text/javascript>
  pages=<%= allPages%>; //parseInt(2.25);
  currentPage=0;
</SCRIPT>

  <TR>
    <TD align=middle height=80>
      <DIV id=pictd>

      <TABLE cellSpacing=0 cellPadding=0 width=200 border=0>
        <TBODY>
        <TR>
          <TD align=middle width=10><IMG onclick="javascript: shiftLeft();" src="../images/arrow_left.gif"></TD>

          <TD align=middle>
            <DIV id="scrollDiv" style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; spacing: 0; BACKGROUND: #ffffff; MARGIN: 0px; OVERFLOW: hidden; WIDTH: 170px; WHITE-SPACE: nowrap; HEIGHT: 40px; offsetLeft: 0;">
            <DIV style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FLOAT: left; PADDING-BOTTOM: 0px; MARGIN-LEFT: 0px; WIDTH: 200%; PADDING-TOP: 0px; spacing: 0">

<%
i=0
response.write "<DIV id=pic_page_" & i/page_pic_num & " style='FLOAT: left; WIDTH: 50%'>"
%>
			<INPUT id="pic_prev" type=hidden value="<%= rsprod("ImgPrev")%>" > 
            <INPUT id="pic800800_prev" type=hidden value="<%= rsprod("ImgPrev")%>" > 
            <DIV style="FLOAT: left; WIDTH: 24.4%" align=center>
			  <IMG id="small_prev" onmouseover="javascript: switchPic('prev');" src="<%= rsprod("ImgPrev")%>" align=center width="40" height="40" border="0" onLoad="DrawImage(this,40,40)">
			</DIV>
<%

i=1
do while not more_pic.eof
	If i/page_pic_num=Int(i/page_pic_num) Then response.write "<DIV id=pic_page_" & i/page_pic_num & " style='FLOAT: left; WIDTH: 50%'>"
%>
			<INPUT id="pic_<%= more_pic("ID")%>" type=hidden value="<%= more_pic("FilePath")%>" > 
            <INPUT id="pic800800_<%= more_pic("ID")%>" type=hidden value="<%= more_pic("FilePath")%>" > 
            <DIV style="FLOAT: left; WIDTH: 24.4%" align=center>
			  <IMG id="small_<%= more_pic("ID")%>" onmouseover="javascript: switchPic('<%= more_pic("ID")%>');" src="<%= more_pic("FilePath")%>" align=center width="40" height="40" border="0" onLoad="DrawImage(this,40,40)">
			</DIV>
<%
	If (i+1)/page_pic_num=Int((i+1)/page_pic_num) Or i+1=total_pic_num Then response.write "</DIV>"
	i=i+1
	more_pic.movenext
Loop
%>

			</DIV>
			</DIV>
		  </TD>

          <TD align=middle width=10><IMG onclick="javascript: shiftRight();" src="../images/arrow_right.gif"></TD>
		</TR>
	    </TBODY>
	  </TABLE>

	  </DIV>
    </TD>
  </TR>
<%
end if
more_pic.close
set more_pic=Nothing
%>
  </TBODY>
</TABLE>

</DIV>
<!-- 左边图片 结束 -->

</DIV>
<!-- 上半部分左边图片和右边简介 结束 -->

<DIV class=clear style="height:20px"></DIV>

<!-- 详细介绍、评论列表、发表评论、推荐给朋友 开始 -->
<DIV align=center>

<div id="tabs">
<a style="cursor:pointer; text-decoration:none" onclick="javascript: expandcontent('sc1', this);"><span>Specs</span> </a><a style="cursor:pointer; text-decoration:none" onclick="javascript: expandcontent('sc2', this);"><span>Comments</span></a><a style="cursor:pointer; text-decoration:none" onclick="javascript: expandcontent('sc3', this);"><span>Evaluate</span></a>
<div class="clear"></div>
</div>

<!-- container 开始 -->
<div id="container">

<div id="sc1" class="tabcontent">
<div class="GoodsDetailsWarp mlr15">
<%=rsprod("enMemoSpec")%>
</div>
</div>

<div id="sc2" class="tabcontent">
<div class="GoodsCommentsWrap">

<div>

<%
response.write "<script language='javascript'>"
If Request("commentpage")="" Then
	response.write "var initialtab=[1, 'sc1'];" 'product.js里面onload函数打开产品详细
Else
	response.write "var initialtab=[2, 'sc2'];"  'product.js里面onload函数打开评论分页
End If 
response.write "</script>"

sqlcomment="select * from goldweb_pinglun where prodid='"&id&"' and view='1' order by AddDate desc"
Set rscomment=Server.CreateObject("ADODB.RecordSet") 
rscomment.open sqlcomment,conn,1,1
if rscomment.eof and rscomment.bof then
	response.write "<br>No comments.<br><br>"
Else
	commentpages = 5
	rscomment.pageSize = 5 '每页记录数
	commentallPages = rscomment.pageCount	'总页数
	commentpage = Request("commentpage")	'从浏览器取得当前评论页

	'if是基本的出错处理
	If not isNumeric(commentpage) then commentpage=1

	if isEmpty(commentpage) or clng(commentpage) < 1 then
		commentpage = 1
	elseif clng(commentpage) >= commentallPages then
		commentpage = commentallPages 
	end if
	rscomment.AbsolutePage = commentpage			'转到某页头部

	response.write "<table cellSpacing=0 cellPadding=0 width=90% border=0 align=center>"
	response.write "<tr><td>"
	do while not rscomment.eof And commentpages>0
		response.write "<table cellSpacing=3 cellPadding=3 width=100% align=center bgColor=#F7FAFD border=0>"

		response.write "<tr><td width='33%'  align='left'>By:"&rscomment("name")&"</td><td widht='33%' align='center'><a href='mailto:"&rscomment("mail")&"'><img border=0 src=../images/small/e-mail.gif></a></td><td width='33%' align='right'>"&rscomment("IP")&"&nbsp;&nbsp;&nbsp;</td></tr>"
		response.write "<tr><td  align='left'>Date: "&rscomment("adddate")&"</td><td align='center'> </td><td align='right'>"
		for j=1 to rscomment("jibie")
			response.write "<IMG src=""../images/star_red.gif"" />"
		next
		response.write "&nbsp;&nbsp;&nbsp;</td></tr>"
		response.write "<tr><td colspan='3' width='100%' align='left'><font color=darkgray><span style='line-height=100%'>"&replace(server.htmlencode(rscomment("nr")),vbCRLF,"<BR>")&"</span></font></td></tr>"
		response.write "</table>"

		response.write "<table border=0 height=10><tr><td></td></tr></table>"
		commentpages = commentpages - 1
		rscomment.movenext
	Loop
	response.write "</td></tr>"
	
	response.write "<tr bgColor=#F7FAFD height=38><td align=center>"
	response.write "Total "&rscomment.RecordCount&" comments, "
	response.write "Page "&commentpage&"/"&commentallPages&", &nbsp;&nbsp;&nbsp;"
	if commentpage = 1 then
	response.write "<font color=darkgray>First Previous</font>"
	else
	response.write "<a href=product.asp?prodid="&prodid&"&commentpage=1>First</a> <a href=product.asp?prodid="&prodid&"&commentpage="&commentpage-1&">Previous</a>"
	end if
	if commentpage = commentallPages then
	response.write "<font color=darkgray> Next Last</font>"
	else
	response.write " <a href=product.asp?prodid="&prodid&"&commentpage="&commentpage+1&">Next</a> <a href=product.asp?prodid="&prodid&"&commentpage="&commentallPages&">Last</a>"
	end if
	response.write "</td></tr>"

	response.write "</table>"

end if
%>

</div>

  <div class="PagerWrap" id="CommentPager"> </div>

  <div id="goodsReviewCenter"></div>

</div>



</div>

<div id="sc3" class="tabcontent">
<%
  Randomize
  yzm=int(8999*rnd()+1000)
%>
<SCRIPT language=javascript>
function checkpinglun(){
	if(pinglunform.name.value=="") {
		alert("Please fill in your name!");
		return false;
	}
	if(pinglunform.yzm.value=="" || pinglunform.yzm.value != "<%= yzm%>") {
		alert("Please fill in correct verify code!");
		return false;
	}
	
    pinglunform.submit();
}
</SCRIPT>
<table cellSpacing="1" cellPadding="2" width="96%" align="center" bgColor="#cccccc" border="0">
	<tr bgColor="#F7FAFD">
	<td>
	<table cellSpacing="1" cellPadding="2" width="95%" align="center" border="0">
	<form name="pinglunform" action="product.asp" method="post">
		<tr>
		<td width="15%" align="left">Name: </td>
		<td width="85%" align="left"><input id="name" value="<%=request.cookies("goldweb")("userid")%>" maxLength="16" name="name" size="15" style="BORDER-RIGHT: #ece2d2 1px solid; BORDER-TOP: #ece2d2 1px solid; FONT-SIZE: 8pt; BORDER-LEFT: #ece2d2 1px solid; COLOR: #666666; BORDER-BOTTOM: #ece2d2 1px solid; FONT-FAMILY: verdana"><font color=orange> 
		<input type="radio" value="1" name="jibie"><IMG src="../images/star_red.gif" />
		<input type="radio" value="2" name="jibie"><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" />
		<input type="radio" value="3" name="jibie"><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" />
		<input type="radio" value="4" name="jibie"><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" />
		<input type="radio" checked value="5" name="jibie"><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /><IMG src="../images/star_red.gif" /></font></td>
		</tr>
		<tr>
		<td align="left">Email: </td>
		<td align="left"><input id="mail" size="30" name="mail" style="BORDER-RIGHT: #ece2d2 1px solid; BORDER-TOP: #ece2d2 1px solid; FONT-SIZE: 8pt; BORDER-LEFT: #ece2d2 1px solid; COLOR: #666666; BORDER-BOTTOM: #ece2d2 1px solid; FONT-FAMILY: verdana"></td>
		</tr>
		<tr>
		<td align="left">Verify Code: </td>
		<td align="left"><input id="yzm" size="6" maxlength=5 name="yzm" style="BORDER-RIGHT: #ece2d2 1px solid; BORDER-TOP: #ece2d2 1px solid; FONT-SIZE: 8pt; BORDER-LEFT: #ece2d2 1px solid; COLOR: #666666; BORDER-BOTTOM: #ece2d2 1px solid; FONT-FAMILY: verdana">
<%
a=int(yzm/1000)
b=int((yzm-a*1000)/100)
c=int((yzm-a*1000-b*100)/10)
d=int(yzm-a*1000-b*100-c*10)
response.write "<img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&a&".gif><img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&b&".gif><img  align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&c&".gif><img align=top height=15 border=0 src=../images/yzm/"&yzm_skin&"/"&d&".gif>"
%></td>
 </tr>
		<tr>
		<td vAlign="top" align="left">Comment: <br><font color=red>≤100 Characters</font></td>
		<td align="left"><textarea id="nr" name="nr" rows="4" cols="70" style="BORDER-RIGHT: #ece2d2 1px solid; BORDER-TOP: #ece2d2 1px solid; FONT-SIZE: 8pt; BORDER-LEFT: #ece2d2 1px solid; COLOR: #666666; BORDER-BOTTOM: #ece2d2 1px solid; FONT-FAMILY: verdana" style="overflow:auto;"></textarea>
		</td></tr>
		<tr>
		<td></td>
		<td align="left">
		<input type=hidden name='prodid' value='<%=prodid%>'>
		<input type=hidden name='save' value='ok'>
		<input type=hidden name='yzmcheck' value='<%=yzm%>'>
		<input type="button" value=" Save " name="submit1" onClick="javascript:checkpinglun();"> &nbsp; 
		<input type="reset" value=" Clear " name="Submit2">
		</td></tr>
        </form>
	</table>
        </td>
        </tr>
</table>
</div>

</div>
<!-- container 结束 -->

</div>
<!-- 详细介绍、评论列表、发表评论、推荐给朋友 结束 -->

<DIV class=clear style="height:20px"></DIV>

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

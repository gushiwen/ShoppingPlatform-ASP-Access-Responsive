<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
Dim newstitle(5)
Set rsn = conn.Execute("select ennewstitle1,ennewstitle2,ennewstitle3,ennewstitle4,ennewstitle5 from shopsetup") 
newstitle(1)=rsn("ennewstitle1")
newstitle(2)=rsn("ennewstitle2")
newstitle(3)=rsn("ennewstitle3")
newstitle(4)=rsn("ennewstitle4")
newstitle(5)=rsn("ennewstitle5")
rsn.close
set rsn=nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>News Center-<%=ensitename%>-<%=SiteUrl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="News Center,Discount,Promotion,<%=ensitename%>">
<meta name="description" content="News Center,Discount,Promotion,<%=ensitename%>">

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
&gt; News Center
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_newsclasstree.asp"-->
<!--商品分类结束-->

</DIV>
<!--左边结束-->


<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>

<DIV STYLE="MARGIN:5px; ">

<table cellspacing=0 cellpadding=0 width=100% border=0 align="center">

<%
set rs=conn.execute ("select * from News where online=true and enNewsTitle<>''")
if rs.eof and rs.bof then
response.write "<tr><td height=50 align=center>No information</td></tr>"
else

for i=1 to 5
set rs=conn.execute ("select top 10 * from News where online=true and NewsClass='"&i&"' and enNewsTitle<>'' order by uup desc,Pubdate desc")

if newstitle(i)<>"" then
if not (rs.eof and rs.bof) then
%>
<tr><td width=98%  valign=top>

<table border=0>
<tr><td width=165 height=28 background=../images/small/newsbar.gif align=center><a href='<%=i&"-"&Replace(newstitle(i)," ","-")%>.html' alt="View more about <%=newstitle(i)%>"><%=newstitle(i)%></a></td></tr>
</table>

<table   cellpadding="2"  cellspacing="0" border=0  width=95% align=center>
<%
	Do while Not rs.eof
	%>
        <tr>
			<td width="55%">&nbsp;&nbsp;&nbsp;&nbsp;<a href="news.asp?NewsId=<%=Cstr(rs("NewsId"))%>"  title="<%=rs("enNewsTitle")%>" ><%=lleft(rs("enNewsTitle"),50)%></a></td>
			<td  width="25%"><%=rs("PubDate")%></td>
			<td  width="20%">Clicked: <%=rs("cktimes")%></td>
		</tr>
	<%
	rs.movenext
	loop
set rs=nothing
%>
</table>
</td></tr>

<tr><td align="right"><a href='<%=i&"-"&Replace(newstitle(i)," ","-")%>.html' alt="View more about <%=newstitle(i)%>">&gt;&gt;View More...</a>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
<%
end if
end if
next
end if
%>

</tbody> 
</table>

</DIV>

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

<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
action=trim(request("class"))

Dim newstitle(5)
Set rsn = conn.Execute("select ennewstitle1,ennewstitle2,ennewstitle3,ennewstitle4,ennewstitle5 from shopsetup") 
newstitle(1)=rsn("ennewstitle1")
newstitle(2)=rsn("ennewstitle2")
newstitle(3)=rsn("ennewstitle3")
newstitle(4)=rsn("ennewstitle4")
newstitle(5)=rsn("ennewstitle5")
rsn.close
set rsn=nothing

title=newstitle(action)
if title="" then 
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.redirect "news_home.html"
	response.end
End If 
rs.close
set rs=Nothing

pages=30
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=Title%>-<%=ensitename%>-<%=SiteUrl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=Title%>,News Center,Discount,Promotion,<%=ensitename%>">
<meta name="description" content="<%=Title%>,News Center,Discount,Promotion,<%=ensitename%>">

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
&gt; <%=title%>
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

<table border=0  width=165>
<tr><td width=165 height=28 background=../images/small/newsbar.gif align=center><%=title%></td></tr>
</table>
<table cellpadding=2 cellspacing=0 border=0  width=95% align=center bgcolor=#FFFFFF>
<%
set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from News where online=true and NewsClass='"&action&"' and enNewsTitle<>'' order by uup desc,Pubdate desc"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<tr><td height=50 align=center>No information</td></tr>"
else
rs.pageSize = pages '每页记录数
allPages = rs.pageCount	'总页数
page = Request("page")	'从浏览器取得当前页

'if是基本的出错处理

If not isNumeric(page) then page=1

if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page			'转到某页头部
do while not rs.eof and pages>0
%>
<tr>
	<td width="55%">&nbsp;&nbsp;&nbsp;&nbsp;<a href="news.asp?NewsId=<%=rs("NewsId")%>" title="<%=rs("enNewsTitle")%>"><%=lleft(rs("enNewsTitle"),50)%></a></td>
	<td width="25%"><%=rs("PubDate")%></td>
	<td width="20%">Clicked: <%=rs("cktimes")%></td>
</tr>
<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop

end if
%>
</table>
</td></tr>

<%
call listpage()
%>

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
sub listpage()
'if allpages<=1 then exit sub
response.write "<tr><td height=30> </td></tr><tr><td colspan=3 height=1 background=../images/small/bgline.gif></td></tr>"
response.write "<tr><td align=center>"
response.write "<br>Total "&RS.RecordCount&" records &nbsp;"
if page = 1 then
response.write "<font color=darkgray>First Previous</font>"
else
response.write "<a href=news_home_more.asp?class="&action&"&page=1>First</a> <a href=news_home_more.asp?class="&action&"&page="&page-1&">Previous</a>"
end if
if page = allpages then
response.write "<font color=darkgray> Next Last</font>"
else
response.write " <a href=news_home_more.asp?class="&action&"&page="&page+1&">Next</a> <a href=news_home_more.asp?class="&action&"&page="&allpages&">Last</a>"
end if
response.write " &nbsp; Page "&page&"/"&allpages&" </td></tr>"
end sub

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 新闻首页
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
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=FullSiteUrl%>"+"/en/news_home.html");</script>
		<link href="<%=FullSiteUrl%>/en/news_home.html" rel="canonical" />

		<title>News Center-<%=ensitename%>-<%=SiteUrl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="News Center,Discount,Promotion,<%=ensitename%>">
		<meta name="description" content="News Center,Discount,Promotion,<%=ensitename%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/newshomestyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/common.js" type="text/javascript"></script>
	</head>

	<body>
		<!--#include file="goldweb_top.asp"-->

		<div class="top-box">
		</div>

<%
For i=1 To 5
	If newstitle(i)<>"" then
%>
		<div class="list-box">
			<div class="list_title"><a href="news_home_more.asp?class=<%=i%>"><%=newstitle(i)%></a></div> 
			<%
				set rs=conn.execute ("select top 10 * from News where online=true and NewsClass='"&i&"' and enNewsTitle<>'' order by uup desc,Pubdate desc")
				Do while Not rs.eof 
			%>
			<div class="list_row">
				<a class="btn_news" href="news.asp?NewsId=<%=Cstr(rs("NewsId"))%>" ><%=rs("enNewsTitle")%></a>
			</div> 
			<%
					rs.movenext
				Loop
				
				rs.close
				set rs=nothing
			%>
			<div class="list_row">
				<input class="btn_submit" type="button" value="View more about <%=newstitle(i)%>" onclick="document.location.href='news_home_more.asp?class=<%=i%>';">
			</div>
		</div>
<%
	end if
next
%>

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>

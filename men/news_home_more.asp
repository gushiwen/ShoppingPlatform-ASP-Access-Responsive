<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
' 新闻列表
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
	response.redirect "news_home.asp"
	response.end
End If 
rs.close
set rs=Nothing

pages=30
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<%
			If InStr(Request.ServerVariables("QUERY_STRING"),"page=") > 0 Then
				PCLocationURL = GetPCLocationURL()
			Else
				PCLocationURL = FullSiteUrl & "/en/"
				If action<>"" Then PCLocationURL = PCLocationURL & action
				If newstitle(action)<>"" Then PCLocationURL = PCLocationURL & "-" & Replace(newstitle(action), " ", "-")
				PCLocationURL = PCLocationURL & ".html"
			End If 
		%>
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=PCLocationURL%>");</script>
		<link href="<%=PCLocationURL%>" rel="canonical" />

		<title><%=Title%>-<%=ensitename%>-<%=SiteUrl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=Title%>,News Center,Discount,Promotion,<%=ensitename%>">
		<meta name="description" content="<%=Title%>,News Center,Discount,Promotion,<%=ensitename%>">

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

		<div class="list-box">
			<div class="list_title"><%=Title%></div> 

<%
set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from News where online=true and NewsClass='"&action&"' and enNewsTitle<>'' order by uup desc,Pubdate desc"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
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

			<div class="list_row">
				<a class="btn_news" href="news.asp?NewsId=<%=rs("NewsId")%>" ><%=rs("enNewsTitle")%></a>
			</div> 

<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>
							<!--分页 -->
							<ul class="am-pagination am-pagination-centered">
							<%
								if page = 1 then
									response.write "<li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;First&nbsp;&nbsp;&nbsp;</a></li><li class=""am-disabled""><a href=""#"">Previous</a></li>"
								else
									response.write "<li class=""am-active""><a href=news_home_more.asp?class="&action&"&page=1>&nbsp;&nbsp;&nbsp;First&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=news_home_more.asp?class="&action&"&page="&page-1&">Previous</a></li>"
								end if
								if page = allpages then
									response.write "<li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-disabled""><a href=""#"">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
								else
									response.write "<li class=""am-active""><a href=news_home_more.asp?class="&action&"&page="&page+1&">&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;</a></li><li class=""am-active""><a href=news_home_more.asp?class="&action&"&page="&allpages&">&nbsp;&nbsp;&nbsp;&nbsp;Last&nbsp;&nbsp;&nbsp;&nbsp;</a></li>"
								end if
							%>
								<p>Total <%=RS.RecordCount%> News, Page <%=page%>/<%=allpages%></p>
							</ul>

<%
end If

rs.close
set rs=nothing
%>
		</div>




		<div class="list-box">
			<div class="list_row">
				<input class="btn_submit" type="button" value="Return to News Home" onclick="document.location.href='news_home.asp';">
			</div> 
		</div> 

		<!--#include file="goldweb_down.asp"-->
	</body>
</html>

<%
conn.close
set conn=nothing
%>
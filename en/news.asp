<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
NewsId=Request("NewsId")
If not isNumeric(NewsId) Then
	conn.close
	set conn=nothing
	response.redirect "news_home.html"
	response.end
end If

set rs=Server.CreateObject("ADODB.RecordSet")
sql= "select * from News where online=true and enNewsTitle<>'' and NewsId="&request("NewsId")
rs.open sql,conn,1,1
if rs.bof and rs.eof Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.redirect "news_home.html"
	response.end
Else
	Title= rs("enNewsTitle")
	NewsContain= rs("enNewsContain")
	NewsClass= rs("enNewsClass")
	NewsSource= rs("enSource")
	PubDate= rs("PubDate")
	Cktimes= rs("cktimes")
	conn.execute "UPDATE news SET ckTimes ="&Cktimes+1&" WHERE NewsId="&NewsId
	rs.close
	set rs=Nothing
end If

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

<title><%=Title%>-<%=ensitename%>-<%=SiteUrl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=Title%>,<%=ensitename%>">
<meta name="description" content="<%=Title%>,<%=ensitename%>">

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
&gt; <%=newstitle(NewsClass)%>
&gt; <%=Title%>
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
<p><h1><%= Title%> </h1></p>
<p>Source: <%= NewsSource%> &nbsp; Published: <%= PubDate%> &nbsp; Clicked: <%=cktimes%> </p>
<p><%= NewsContain%> </p>
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

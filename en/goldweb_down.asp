<div class="bottom blueT"  id="yihaodianBottom">

<%
Set rsp = conn.Execute("select * from page") 
bottomGuidTitle = "About; Service; Corporate; Member; Help"
'body99=rsp("body99")
%>

<div class="bottom1">
<div style="width:950px">
<%
for lie=1 to 5
response.Write "<div class=""list_b left"" style=""width:190px"">"
response.write "<ul class=""gray2"">" & Split(bottomGuidTitle, ";")(lie-1)
for hang=1 to help_hang
ID=lie*10+hang
ID=cstr(id)
title="entitle"&id
url="url"&id
'body="enbody"&id
response.write "<li>"
if rsp(url)<>"" then
response.write "<a href='"&rsp(url)&"'>&nbsp;"&rsp(title)&"&nbsp;</a>"
else
response.write "<a href='page.asp?id="&id&"'>&nbsp;"&rsp(title)&"&nbsp;</a>"
end if
response.write "</li>"
next
response.write "</ul>"
response.write "</div>"
next

%>
</div>
<div class="clear"></div>
</div>

<div class="mt6">Friend Link：
<%
  Set rs = conn.Execute("select top 10 * from links where online='1' order by txt desc,num asc")
  do while not rs.eof
%>
<a href="<%=rs("url")%>" target="_blank"><%=rs("site")%></a><span class="gray2"> | </span>
<%
  rs.movenext
  loop
  rs.close
  set rs=Nothing
%>
</div>

<div class="clear" style="height:10px"></div>

<div class="mt6">
<%= enadm_comp%>&nbsp;&nbsp;Tel: <%= adm_tel%>&nbsp;&nbsp; Email: <%= adm_mail%>&nbsp;&nbsp;<br />
Website: <%=siteurl%>, Address: <%= enadm_address%>&nbsp;&nbsp; 

<script type="text/javascript">
// JS to check admin login or not
if(getCookie("adminid")=="")
{
  document.write('<A href="../admin.asp" target=_blank>Admin Login</A> ');
}
else
{
  document.write('<A href="../admin/adminindex.asp" target=_blank>Admin '+getCookie("adminid")+'</A> ');
}
</script>

<a href="http://www.ol.sg" target=_blank><img border="0" src="../images/supplier.gif" alt="RichmanNetwork" title="RichmanNetwork"></a>
</div>

<%
rsp.close
set rsp=Nothing
%>

</div>

<!--统计访问量开始-->
<script type="text/javascript">
document.write('<scr' + 'ipt type="text/javascript" src="../include/goldweb_tj.asp?where=' + URLencode(document.referrer) + '"><' + '/script>');
</script>
<!--统计访问量结束-->

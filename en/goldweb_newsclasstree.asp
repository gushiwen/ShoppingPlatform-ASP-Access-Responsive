<DIV class="border4 mt6">
<DIV class="left_title2"><STRONG>News Classify</STRONG></DIV>
<DIV class="list1">
<UL>
<% ' 新闻分类
	For i=1 To 5
		If newstitle(i)<>"" then
%>
	<LI>
		<a href="<%=i&"-"&Replace(newstitle(i)," ","-")%>.html" alt="View <%=newstitle(i)%>"><%=newstitle(i)%></a>
	</LI>
<%
		End if
	next
%>
</UL>
</DIV>
</DIV>
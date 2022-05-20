<DIV class="border4 mt6">
<DIV class="left_title2"><STRONG>Category</STRONG></DIV>
<DIV class="category-01">
<UL>
<% ' 商品分类
    if LarCode="" then 
		sql="select * from goldweb_class where MidSeq=1 order by larseq"
    Else
		sql="select * from goldweb_class where enLarCode='"&LarCode&"' and MidSeq=1"
    End If 
    Set rs=Server.CreateObject("ADODB.Recordset")
    rs.open sql,conn,1,3
    do while not rs.eof
%>
<LI class="limb-01">
<h2><a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>" ><%=rs("enLarCode")%></a></h2>
<UL>
<%
  Set rsm=Server.CreateObject("ADODB.Recordset")
  sqlm="select * from goldweb_class where enlarcode='"&rs("enLarCode")&"' and midseq<>1 order by midseq"
  rsm.open sqlm,conn,1,3
  do while not rsm.eof
%>
    <li <% if LarCode="" Then response.write "class=""lifo-01"""%>><a href="productlist.asp?LarCode=<%=Server.URLEncode(rsm("enLarCode"))%>&MidCode=<%=Server.URLEncode(rsm("enMidCode"))%>&ClassId=<%=rsm("ClassId")%>" title="<%=rsm("enMidCode")%>"><%=rsm("enMidCode")%></a></LI>
<%
  rsm.movenext
  loop
  rsm.close
  set rsm=Nothing
%>    
</UL>
</LI>
<%
    rs.movenext
    loop
    rs.close
    set rs=Nothing
%>
</UL>
</DIV>
</DIV>
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from goldweb_class where MidSeq=1 order by larseq"
  rs.open sql,conn,1,3
  do while not rs.eof
  larcode=rs("enLarCode")
  If rs("LarSeq") Mod 5 <> 0 Then
	classseqid = rs("LarSeq") Mod 5
  Else
	classseqid = 5
  End If 
%>

<!--推荐产品大分类开始-->
<div style="margin-top:15px; margin-bottom:5px; ">
  <div class="class<%=classseqid%>_bg">
    <div class="class<%=classseqid%> left class<%=classseqid%>_title" style="width:150px; text-align:left">
      <h1><a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>"><%=rs("enLarCode")%></a></h1>
    </div>
    <div class="right h27"><a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>" class="white">&gt;&gt;More</a>&nbsp;&nbsp;</div>
  </div>
   <div class="left">
   <div class="product_line3">
    <ul>
<%
  Set rsm=Server.CreateObject("ADODB.Recordset")
  sqlm="select top 4 * from goldweb_product where enlarcode='"&rs("enLarCode")&"' and online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) and Remark='1' order by TJDate desc"
  rsm.open sqlm,conn,1,3
  do while not rsm.eof
%>
	              <li>
			        <div class="img11"><a href="product.asp?ProdId=<%=rsm("ProdId")%>" title="<%=rsm("enProdName")%>"><img src="<%=rsm("ImgPrev")%>" title="<%=rsm("enProdName")%>" width="115" height="115" border="0" onLoad="DrawImage(this,115,115)"/></a><span class="Edge"></span></div>
			        <div class="center">
					<% 
						If CSng(rsm("enPriceList")) > 0 then response.write "<span class=""red14"">"&rsm("enPriceUnit")&comma(rsm("enPriceList"))&"</span> "
						If CSng(rsm("enPriceOrigin")) >0 Then response.write "<span class=""listprice gray"">"&rsm("enPriceUnit")&comma(rsm("enPriceOrigin"))&"</span>"
					%>
						<br />
						<h3><a href="product.asp?ProdId=<%=rsm("ProdId")%>" title="<%=rsm("enProdName")%>"><%=rsm("enProdName")%></a></h3>
			        </div>
			      </li>

<%
  rsm.movenext
  loop
  rsm.close
  set rsm=Nothing
%>
  </ul>
  </div>
  </div>

  <div class="clear"></div>
</div>
<!--推荐产品大分类结束-->
<%
  rs.movenext
  loop
  rs.close
  set rs=Nothing
%>
<div class="mt6" style="height:28px;">
<div class="tag_bg3" id="leftTopTab1" style="cursor:pointer" onmouseover="changeLeftTopTab(1)">Category</div>
<div class="tag_bg3_" id="leftTopTab2" style="cursor:pointer" onmouseover="changeLeftTopTab(2)">Popular</div>
</div>

<!--商品分类开始-->
<div class="border3" id="leftTopContent_1" >
<div class="category" >
<ul class="trunk">
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from goldweb_class where MidSeq=1 order by larseq"
  rs.open sql,conn,1,3
  do while not rs.eof
  larcode=rs("enLarCode")
%>
<li class="limb-01">
<h2><a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>" ><%=rs("enLarCode")%></a></h2>
<ul>
<%
  Set rsm=Server.CreateObject("ADODB.Recordset")
  sqlm="select * from goldweb_class where enlarcode='"&larcode&"' and midseq<>1 order by midseq"
  rsm.open sqlm,conn,1,3
  do while not rsm.eof
%>

    <li><a href="productlist.asp?LarCode=<%=Server.URLEncode(rsm("enLarCode"))%>&MidCode=<%=Server.URLEncode(rsm("enMidCode"))%>&ClassId=<%=rsm("ClassId")%>" title="<%=rsm("enMidCode")%>"><%=rsm("enMidCode")%></a></li>

<%
  rsm.movenext
  loop
  rsm.close
  set rsm=Nothing
%>    
</ul>
</li>
<%
  rs.movenext
  loop
  rs.close
  set rs=Nothing
%>

</ul>
</div>
<div class="clear"></div>
</div>
<!--商品分类结束-->

<!--热门商品开始-->
<div class="border3 clear" id="leftTopContent_2" style="display:none">
<div class="product_line_main" id="newproduct_content_0"  >
           <ul>
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  sql="select top " & renmen_num & " * from goldweb_product where online=true and (ExpiryDate is null or ("&DateValue(DateAdd("h", TimeOffset, now()))&"<=ExpiryDate and DateValue(AddDate)<>ExpiryDate)) order by ClickTimes desc, AddDate desc"
  rs.open sql,conn,1,3
  do while not rs.eof
%> 
		      <li>
		        <div class="img_main left"> <a href="product.asp?ProdId=<%=rs("ProdId")%>" title="<%=rs("enProdName")%>" ><img src="<%= rs("ImgPrev")%>" width="48" height="48" border="0" onload='DrawImage(this,48,48)' /></a></div>
		        <div class="right line_main">
				<%
					If rs("enModel") <> "" Then response.write"<a href='product.asp?ProdId="&rs("ProdId")&"' title='"&rs("enProdName")&"'>"&lleft(rs("enModel"),20)&"</a><br />"
					If CSng(rs("enPriceList")) > 0 then response.write "<span class=""red14"">"&rs("enPriceUnit")&comma(rs("enPriceList"))&"</span> "
				%>
		        </div>
		      </li>
<%
  rs.movenext
  loop
  rs.close
  set rs=Nothing
%>
           </ul>
<div class="clear"></div>
</div>
</div>
<!--热门商品结束-->

<!--首页左侧图片链接开始-->
<%
If iadv1<>"" Then
response.write "<a href='" & iadvurl1 & "'><img src='" & iadv1 & "' width='205' height='100' border='0' style='cursor:pointer' /></a>"
End If 
If iadv5<>"" Then
response.write "<a href='" & iadvurl5 & "'><img src='" & iadv5 & "' width='205' height='100' border='0' style='cursor:pointer' /></a>"
End If 
%>
<!--首页左侧图片链接结束-->

<!--最新资讯开始-->
<%
if tree_view<>0 then
%>
<div class="border4 mt6">
  <div class="right_title2">
    <div class="font14" style="width:120px; float:left;"><strong>Latest News</strong></div>
  </div>

  		  <div class="list1 blue" >
            <ul>
<%
  set rs=conn.execute ("select top " & tree_num & " * from News where online=true and enNewsTitle<>'' order by Pubdate desc")
  do while not rs.eof
%> 
	<li>
      <a href="news.asp?NewsId=<%=Cstr(rs("NewsId"))%>" title="<%= rs("enNewsTitle")%>"><%= lleft(rs("enNewsTitle"),26)%></a>
	</li>
<%
  rs.movenext
  loop
  rs.close
  set rs=Nothing
%>
            </ul>
          <div class="clear"></div>
          </div>

</div>
<%
End If 
%>
<!--最新资讯结束-->

<!--热门调查开始-->
<%
Set rsVote= conn.execute("select top 1 * from Vote where enview='1'")
if (rsVote.bof and rsVote.eof) then 

Else
%>
<div class="border4 mt6" style="height:200px;">
<form name='VoteForm' method='post' action='vote.asp' target='_blank'>
  <div class="right_title2">
    <div class="font14" style="width:80px; float:left;"><strong>Survey</strong></div>
  </div>

  <div class="list_b">
<%
response.Write "<ul>"
response.write "<li>"&rsVote("enTitle") & "</li>"

if rsVote("VoteType")="dan" then
for i=1 to 8
if trim(rsVote("enSelect" & i) & "")="" then exit for
response.Write "<li><input type='radio' name='VoteOption' class='inputno' value='" & i & "'>" & " " & rsVote("enSelect" & i) & "</li>"
next
else
for i=1 to 8
if trim(rsVote("Select" & i) & "")="" then exit for
response.Write "<li><input type='checkbox' name='VoteOption' class='inputno' value='" & i & "'>" & " " & rsVote("enSelect" & i) & "</li>"
next
end If

response.Write "</ul>"
%>
      <input name='VoteType' type='hidden' value='<%=rsVote("VoteType")%>'>
      <input name='Action' type='hidden' value='vote'>
    <div style="PADDING-LEFT: 60px; PADDING-BOTTOM: 10px;PADDING-TOP: 10px; ">
      <input type="button" class="button4" value=" Vote " style="cursor:pointer" onclick="javascript: VoteForm.submit();"/>
      <input type="button" class="button4" value=" View " style="cursor:pointer" onclick="javascript:location.href='Vote.asp?Action=view';"/>
    </div>
  </div>

</form>
</div>
<%
end if
rsVote.close
set rsVote=nothing
%>
<!--热门调查结束-->
<!--滚动开始-->
<div class="slides">
	<div class="camera_wrap camera_emboss pattern_1" id="camera_wrap_4">
		<div data-src="<%=pic1%>">
		</div> 
		<div data-src="<%=pic2%>">
		</div>
        <div data-src="<%=pic3%>">
        </div> 
        <div data-src="<%=pic4%>">
        </div> 
      </div><!-- #camera_wrap_3 -->
</div>
<!--滚动结束-->

<!--促销信息开始-->
<div class="three mt6">

<%
dim newstitle(5)
Set rs = conn.Execute("select * from shopsetup") 
newstitle(1)=rs("ennewstitle1")
newstitle(2)=rs("ennewstitle2")
newstitle(3)=rs("ennewstitle3")
newstitle(4)=rs("ennewstitle4")
newstitle(5)=rs("ennewstitle5")
set rs=Nothing

dim iadvcenter(3)
iadvcenter(1)=iadv6
iadvcenter(2)=iadv7
iadvcenter(3)=iadv8

dim iadvurlcenter(3)
iadvurlcenter(1)=iadvurl6
iadvurlcenter(2)=iadvurl7
iadvurlcenter(3)=iadvurl8
%>


<%
set rs=conn.execute ("select * from News where online=true")
if rs.eof and rs.bof then
'response.write "No infomation"
else

for i=1 to 3

' 开始1个图片和资讯显示
If iadvcenter(i)="" And newstitle(i)="" Then 
' 图片和资讯都为空则不显示
else
response.write "<div class='left border4 mr5' style='width:179px;'>"

' 显示中间图片
If iadvcenter(i)<>"" Then
response.write "<div class='clear' style='margin-top:5px; margin-left:4px;'><img src='" & iadvcenter(i) & "' width='172' height='134'/></div>"
End If 

'开始资讯分类显示
if newstitle(i)<>"" Then
%>
  <div class="three_title">
    <a href="news_home_more.asp?class=<%=i%>" title="<%= newstitle(i)%>"><strong><font color="#F87501"><%= newstitle(i)%></font></strong></a>
  </div>

  <div class="list1 blue">
    <ul>
<%
set rs=conn.execute ("select top 3 * from News where online=true and NewsClass='"&i&"' and enNewsTitle<>'' order by uup desc,Pubdate desc")

if not (rs.eof and rs.bof) then
	Do while Not rs.eof
%>
      <li>
		<a href="news.asp?NewsId=<%=Cstr(rs("NewsId"))%>" title="<%= rs("enNewsTitle")%>"><%= lleft(rs("enNewsTitle"),22)%></a>
	  </li>
<%
	rs.movenext
	loop
end If
%>
    </ul>
  </div>
<%
end If
' 结束资讯分类显示

response.write "</div>"
End If 
' 结束1个图片和资讯显示

Next

end If

set rs=nothing
%>

</div>
<!--促销信息结束-->
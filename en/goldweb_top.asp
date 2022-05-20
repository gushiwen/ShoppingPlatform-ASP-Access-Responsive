<%
if kaiguan=0 then
response.write guanbi
response.end 
end if

Set ggrs = conn.Execute("select * from adv") 
pic1= ggrs("pic1")
pic1_lnk= ggrs("pic1_lnk")
tit1= ggrs("tit1")
pic2= ggrs("pic2")
pic2_lnk= ggrs("pic2_lnk")
tit2= ggrs("tit2")
pic3= ggrs("pic3")
pic3_lnk= ggrs("pic3_lnk")
tit3= ggrs("tit3")
pic4= ggrs("pic4")
pic4_lnk=ggrs("pic4_lnk")
tit4= ggrs("tit4")
flash=ggrs("flash")
flashwidth=ggrs("flashwidth")
flashheight=ggrs("flashheight")
flashurl=ggrs("flashurl")
flash2=ggrs("flash2")
flash2width=ggrs("flash2width")
flash2height=ggrs("flash2height")
flash2url=ggrs("flash2url")
bannerxiaoguo=ggrs("bannerxiaoguo")
bannerxiaoguo_top=ggrs("bannerxiaoguo_top")
logo=ggrs("logo")
hfpic= ggrs("hfpic")
hfurl= ggrs("hfurl")
hftit= ggrs("hftit")
hf2pic= ggrs("hf2pic")
hf2url= ggrs("hf2url")
hf2tit= ggrs("hf2tit")
piaofu=ggrs("piaofu")
piaofuurl=ggrs("piaofuurl")
piaofupic=ggrs("piaofupic")
piaofutit=ggrs("piaofutit")
tanchu=ggrs("tanchu")
tanchu_time=ggrs("tanchu_time")
tanurl=ggrs("tanurl")
tanheight=ggrs("tanheight")
tanwidth=ggrs("tanwidth")
tantop=ggrs("tantop")
tanleft=ggrs("tanleft")
cebian=ggrs("cebian")
lefturl=ggrs("lefturl")
righturl=ggrs("righturl")
leftpic=ggrs("leftpic")
rightpic=ggrs("rightpic")

iadv1=ggrs("iadv1")
iadv2=ggrs("iadv2")
iadv3=ggrs("iadv3")
iadv4=ggrs("iadv4")
iadv5=ggrs("iadv5")
iadv6=ggrs("iadv6")
iadv7=ggrs("iadv7")
iadv8=ggrs("iadv8")
iadvurl1=ggrs("iadvurl1")
iadvurl2=ggrs("iadvurl2")
iadvurl3=ggrs("iadvurl3")
iadvurl4=ggrs("iadvurl4")
iadvurl5=ggrs("iadvurl5")
iadvurl6=ggrs("iadvurl6")
iadvurl7=ggrs("iadvurl7")
iadvurl8=ggrs("iadvurl8")
ggrs.close
set ggrs=nothing
%>

<!-- TOP START header-skin -->
<div class="header-skin">

  <!-- TOP START header-top -->
  <div class="header-top" style=" Z-INDEX: 100;">
    
    <!-- TOP START LOGO -->
	<div class="logo" title="<%=ensitename%>" id="logo_top_id"><h1><a href="#" style=" BACKGROUND: url(<%=logo%>) no-repeat 0px 0px; "><%=ensitename%> Shopping Online </a></h1></div>
	<!-- TOP END LOGO -->

    <div class="slogan"><%=ensitename%></div>


<div class="quick-menu">
  <img src="../images/shoppingcart.png" align="bottom" border="0" /> 
  <a href="shoppingcart.asp">Shopping Cart</a> <span class="gray2">|</span> <img src="../images/favorites.png" align="bottom" border="0" /> <a href="my_fav.asp">My Favorites</a> <span class="gray2">|</span> 
<script type="text/javascript">
// JS checking login and JS write page
if(getCookie("userid")=="")
{
  document.write('<a href="login.asp">Log In</a> <span class="gray2">|</span> <a href="reg_member.asp">Register</a> <span class="gray2">|</span> ');
}
else
{
  document.write('<a href="user_center.asp">My Account</a> <span class="gray2">|</span> <a href="logout.asp">Log Out</a> <span class="gray2">|</span> ');
}
</script>
 <a href="../ch/main.html">中文</a>
</div>

    <!-- header-main START -->
    <div class="header-main">

      <!-- header-main-skin START -->
      <div class="header-main-skin">

            <!--频道菜单开始-->
            <div class="channel-menu">
                <div class="channel-lv1">
                    <ul>
                        <li class="li-index">
                            <ul>
                                <li <%If InStr(request.servervariables("PATH_INFO"),"main.asp")>0 Or InStr(request.servervariables("PATH_INFO"),"AreaTemplate.asp")>0 Then
								      response.write " class=""selected""" 
									  else
								      response.write " class=""gray""" 
									  End if%>><a href="main.html"><strong>Home</strong></a></li>
                            </ul>
                        </li>

                        <%
						  Dim LarCodeStr
						  If InStr(request.servervariables("PATH_INFO"),"productlist.asp")>0 Then
						    LarCodeStr = request.servervariables("QUERY_STRING")
						  ElseIf InStr(request.servervariables("PATH_INFO"),"product.asp")>0 Then
						    LarCodeStr = Server.URLEncode(conn.execute("Select enLarCode from goldweb_product where ProdId='"&Trim(Split(request.servervariables("QUERY_STRING"),"=")(1))&"'")("enLarCode"))
						  End If
						  
                          Set rsc=Server.CreateObject("ADODB.Recordset")
                          sqlc="select ClassId,enLarCode from goldweb_class where MidSeq=1 order by larseq"
                          rsc.open sqlc,conn,1,3
                          do while not rsc.eof
                        %>
                        <li class="li-index">
                            <ul>
                                <li <%If InStr(LarCodeStr,Server.URLEncode(rsc("enLarCode")))>0 Then '
								      response.write " class=""selected""" 
									  else
								      response.write " class=""gray""" 
									  End if%>><a href="productlist.asp?LarCode=<%=Server.URLEncode(rsc("enLarCode"))%>&ClassId=<%=rsc("ClassId")%>" ><strong><%=rsc("enLarCode")%></strong></a></li>
                            </ul>
                        </li>
                        <%
                          rsc.movenext
                          loop
                          rsc.close
                          set rsc=Nothing		 
                        %>

                    </ul>
                </div>
            </div>
            <!--频道菜单结束-->

            <!-- 搜索开始 -->

            <!-- Search Left Start -->
            <div class="left" style="padding:38px 20px 0 50px;">
              <a href="news_home.html" class="white"><b>Newsletter</b></a>
            </div>
            <!-- Search Left End -->

            <!-- search-box start -->
            <div class="search-box">
              <form name="FORM_TPL_SEARCH" action="productsearch.asp" method="post">
                <input name="action" type="hidden" value="topsearch">
                <input name="keywords" type="text" style="padding-left:2px; width:380px; height:18px; BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666;" value="<%=keywords%>" autocomplete="off" />
            	<a href="#" onclick="javascript:FORM_TPL_SEARCH.submit();"><img src="images/icon_search.gif" width="53" height="20" align="absmiddle" border="0"/></a> 
              </form>
            </div>
            <!-- search-box end -->

            <!-- hot product start -->
            <div class="hotkewords">
            <% set search_rs=Server.CreateObject("ADODB.recordset")
                    search_sql="select top 3 * from goldweb_key order by num desc, keydate desc"
                    search_rs.open search_sql,conn,1,3
                   do while not search_rs.eof
			          response.write "&nbsp;&nbsp;"&search_rs("keywords")
                      search_rs.movenext
                   loop
                   search_rs.close
				   Set search_rs=nothing
			%>
            </div>
            <!-- hot product end -->
            <!-- 搜索结束 -->

      </div>
      <!-- header-main-skin END -->

    </div>
    <!-- header-main END -->

  </div>
  <!-- TOP END header-top -->

</div>
<!-- TOP END header-skin -->
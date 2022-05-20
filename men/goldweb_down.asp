<div class="footer ">
	<div class="footer-hd ">
	</div>
</div>
<!--底部导航 -->
<div class="navCir">
	<li <%If InStr(request.servervariables("PATH_INFO"),"main")>0 Then response.write " class=""active""" %>><a href="main.asp"><i class="am-icon-home "></i>Home</a></li>
	<li <%If InStr(request.servervariables("PATH_INFO"),"category")>0 Then response.write " class=""active""" %>><a href="category.asp"><i class="am-icon-list"></i>Category</a></li>
	<li <%If InStr(request.servervariables("PATH_INFO"),"cart")>0 Then response.write " class=""active""" %>><a href="shoppingcart.asp"><i class="am-icon-shopping-basket"></i>Cart</a></li>	
	<li <%If InStr(request.servervariables("PATH_INFO"),"login")>0 Or InStr(request.servervariables("PATH_INFO"),"user_center")>0 Then response.write " class=""active""" %>>
		<script type="text/javascript">
			// JS checking login and JS write page
			if(getCookie("userid")=="")
			{
				document.write('<a href="login.asp"><i class="am-icon-user"></i>Login</a>');
			}
			else
			{
				document.write('<a href="user_center.asp"><i class="am-icon-user"></i>Account</a>');
			}
		</script>
	</li>
</div>

<!--统计访问量-->
<script type="text/javascript" src="../include/goldweb_tj.asp?where=<%=server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>"></script>
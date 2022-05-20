<%
'if kaiguan=0 then
'response.write guanbi
'response.end 
'end if
%>
<!--悬浮搜索框-->
<div class="nav white">
	<div class="search-bar pr">
		<form name="FORM_TPL_SEARCH" action="productsearch.asp" method="post">
			<input name="action" type="hidden" value="topsearch">
			<input id="searchInput" name="keywords" value="<%=keywords%>" type="text" placeholder="Search Products" autocomplete="off">
			<input id="ai-topsearch" class="submit am-btn" value="Search" index="1" type="submit">
		</form>
	</div>
	<input class="btn_post" type="button" value="Post Ad" onclick="document.location.href='my_adv.asp';">
	<div class="clear"></div>
</div>

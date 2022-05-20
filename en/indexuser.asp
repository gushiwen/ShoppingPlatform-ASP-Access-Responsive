<!--小右边开始-->
<script type="text/javascript">
var rdm = Math.random();
var yzm = Math.floor(8999*rdm+1000);
var yzm_skin=3;

var yzm_a = Math.floor(yzm/1000);
var yzm_b = Math.floor((yzm-yzm_a*1000)/100);
var yzm_c = Math.floor((yzm-yzm_a*1000-yzm_b*100)/10);
var yzm_d = Math.floor(yzm-yzm_a*1000-yzm_b*100-yzm_c*10);

// 个人用户登陆
function  CheckPLoginForm() 
{
    if(document.ploginform.userid.value=="")
    {
		alert("Please input your account name!");
		return false;
    }
    if(document.ploginform.password.value=="") 
    {
		alert("Please input password!");
		return false; 
    }	

	if(document.ploginform.verifycode.value=="" || document.ploginform.verifycode.value != yzm)
	{
		alert("Please input correct verify code!");
		return false;
	}

    document.ploginform.submit();
}
// 订单查询
function  CheckOrderForm() 
{
    if(document.orderform.OrderNum.value=="")
    {
		alert("Please input order number!");
		return false;
    }
    if(document.orderform.MobilePhone.value=="" && getCookie("userid")=="") 
    {
		alert("Please input mobile number!");
		return false; 
    }	
    document.orderform.submit();
}
</script>

<div class="mt6" style="height:28px;">
<div class="tag_bg3" id="rightTopTab_1" onmouseover="changeRightTop(1)" style="cursor:pointer"><img src="../images/icon_login.gif" width="17" height="19" align="absmiddle"/> Member</div>
<div class="tag_bg3_" id="rightTopTab_2" onmouseover="changeRightTop(2)" style="cursor:pointer"> Hotline</div>
</div>
  <div class="border3" id="rightTopContent_1">
                        <div style="padding-top:15px;padding-bottom:5px;padding-left:10px;padding-right:25px;height:120px" >

<script type="text/javascript">

// JS to check user login or not
if(getCookie("userid")=="")
{
  document.write('<form name="ploginform" action="indexlogin.asp" method="post">');
  document.write('  <div style="padding-bottom:5px;text-align:right;">Account <input name="userid" type="text" class="login_input" style="width:106px" /></div>');
  document.write('  <div style="padding-bottom:5px;text-align:right;">Password <input name="password" type="password" class="login_input" style="width:106px" onkeydown="doEnter(event);"/></div>');
  document.write('  <div style="padding-bottom:5px;text-align:right;">Verify <input name="verifycode" type="text" class="login_input" style="width:55px" /> <img align=top height=15 border=0 src=../images/yzm/'+yzm_skin+'/'+yzm_a+'.gif><img align=top height=15 border=0 src=../images/yzm/'+yzm_skin+'/'+yzm_b+'.gif><img  align=top height=15 border=0 src=../images/yzm/'+yzm_skin+'/'+yzm_c+'.gif><img align=top height=15 border=0 src=../images/yzm/'+yzm_skin+'/'+yzm_d+'.gif></div>');
  document.write('  <div style="padding-top:6px;text-align:right;"><input type="button" value=" Log In " onclick="javascript: CheckPLoginForm();" style="font-size:10px; padding:2px;" />&nbsp;&nbsp;<input type="button" value=" Register " onclick="document.location.href=\'reg_member.asp\';"  style="font-size:10px; padding:2px;"/></div>');
  document.write('  <input type="hidden" name="login" value="ok"><input type="hidden" name="cook" value="0">');
  document.write('</form>');
  document.write('  <div style="padding-top:6px;text-align:right;"><span class="gray"><a href="getpsw.asp">Forgot Password</a></span> </div>');
}
else
{
  document.write('Hi '+getCookie("userid")+', ');
  document.write('<br/>Welcome to our online shop,');
  if(getCookie("userkou")!=10){
    document.write('<br/>you have -'+(100-getCookie("userkou")*10)+'% discount.');
  }
  else
  {
  document.write('<br/>you can order now, and receive products quickly.');
  }
  document.write('<br/><br/><a href="shoppingcart.asp">Shopping Cart</a>');
  document.write('<br/><a href="my_fav.asp">Favourite Products</a>');
  document.write('<br/><a href="user_center.asp">My Account</a>');
}

</script>
						  
                        </div>
  </div>
  <div class="border3" id="rightTopContent_2"  style="display:none">
                        <div style="padding:10px;height:120px;color:<%=kf_color%>;" >
						  <%=enadm_kf%>
                        </div>
                  
  </div>

  <div class="border4 mt6">
    <div class="right_title2">
      <div class="font14" style="width:100px; float:left;"><strong>Order Enquiry</strong></div>
    </div>
    <div style="margin:3px;">
      <div style="padding-top:10px;padding-bottom:5px;padding-left:10px;padding-right:15px;" >
	  <form name="orderform" action="my_order_view.asp" method="post">
	  <div style="padding-bottom:5px;text-align:right;">Order No. <input type="text" name="OrderNum" value="" class="login_input" style="width:106px"></div>

<script type="text/javascript">
// JS to check user login or not
if(getCookie("userid")=="")
{
  document.write('<div style="padding-bottom:5px;text-align:right;">Mobile No. <input type="text" name="MobilePhone" value="" class="login_input" style="width:106px;"></div>');
}
else
{
  document.write('<div style="padding-bottom:5px;text-align:right;display:none;">Mobile No. <input type="text" name="MobilePhone" value="" class="login_input" style="width:106px;"></div>');
}
</script>

	  <div style="padding-bottom:5px;text-align:right;"><input type="button" value="  Submit  " onclick="javascript: CheckOrderForm();" style="font-size:10px; padding:2px;"></div>
	  </form>
    </div>
	</div>
  </div>

  <%If iadv2<>"" Or iadv3<>"" Then %>
  <div class="border4 mt6">
    <div class="right_title2">
      <div class="font14" style="width:100px; float:left;"><strong>Recommend</strong></div>
    </div>
    <%If iadv2<>"" Then %>
    <div style="margin:3px;">
	  <a href="<%= iadvurl2%>"><img src="<%= iadv2%>" width="197" height="88" border="0" style="cursor:pointer" /></a>
	</div>
	<%End If %>
    <%If iadv3<>"" Then %>
    <div style="margin:3px;">
	  <a href="<%= iadvurl3%>"><img src="<%= iadv3%>" width="197" height="88" border="0" style="cursor:pointer" /></a>
	</div>
	<%End If %>
  </div>
  <%End If %>

  <div class="border4 mt6">
    <div class="right_title2">
      <div class="font14" style="width:100px; float:left;"><strong>Introduction</strong></div>
    </div>
   
    <div style="margin:3px;"><%= conn.Execute("select enbody99 from page")("enbody99")%></div>
  </div>
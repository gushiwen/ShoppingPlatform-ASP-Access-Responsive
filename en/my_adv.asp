<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
call checklogin()
If request("DelProdId")<>"" Then
	Call delete()
End If 
If request("FriendEmail")<>"" Then
	Call recommend()
End If 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>My Advertisements-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script type="text/javascript">
//判断表单输入正误
function Recommend()
{
	if (document.getElementById("FriendEmail").value.length <10 || document.getElementById("FriendEmail").value.length >50) {
		alert("Please input a valid email address.");
		document.getElementById("FriendEmail").focus();
		return false;
	}
	if (document.getElementById("FriendEmail").value.length > 0 && !document.getElementById("FriendEmail").value.match( /^.+@.+$/ ) ) {
	    alert("Please input a valid email address.");
		document.getElementById("FriendEmail").focus();
		return false;
	}
	document.getElementById("RecommendFriend").disabled=true;
	location.href="my_adv.asp?FriendEmail="+document.getElementById("FriendEmail").value;
}
</script>
</head>
<body>
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; My Advertisements
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_usertree.asp" -->
<!--商品分类结束-->

</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class=mt6 id=homepage_center_new_new>

<TABLE cellSpacing=0 cellPadding=0 width="96%" align=center border=0>
	<TR height="40px">
		<TD align="left">
			<%
				LeftAdvDays = conn.execute("select AdvDays from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("AdvDays")
				response.write "<b>You can add or extend product advertising display for " & LeftAdvDays & " days.</b>"
			%>
		</TD>
	</TR>

	<TR height="30px">
		<TD align="left" valign="middle">
			Recommend our website to one friend to get 30 days more for product advertising display. <br>
			Email Address <input name="FriendEmail" id="FriendEmail" value="" type="text" size="28" maxlength="100" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> <input name="RecommendFriend" id="RecommendFriend" type="button" value=" Recommend " style="font-size:10px;" onclick="Recommend();">
		</TD>
	</TR>
	<TR height="10px">
		<TD align="left">
		</TD>
	</TR>

	<TR height="30px">
		<TD align="left">
		<%
			' ShowAdvBtn = False 
			AdvDaysArray = Split(AdvDaysSetting,",")
			For Each AdvDaysValue in AdvDaysArray
				response.write "<input type=""button"" value=""  Buy "&AdvDaysValue&" days  "" onclick=""document.location.href='productorder.asp?ProdId=00003&Quantity="&AdvDaysValue&"';"" style=""font-size:10px; padding:2px;"">&nbsp;&nbsp;&nbsp;&nbsp;"
				' If CInt(AdvDaysValue) <= CInt(LeftAdvDays) Then ShowAdvBtn = True 
			Next
		%>
		</TD>
	</TR>
	<TR height="20px">
		<TD align="left">
		</TD>
	</TR>
    <TR>
      <TD width="100%" height="30px" align="center">
		<%
			response.write "<input type=""button"" value=""Add Product Advertising Display on the Web Platform"" onclick=""document.location.href='my_prod.asp';"" style=""width:100%; font-weight:bold; font-size:10px; padding:6px 20px;"">"
		%>
	  </TD>
    </TR>

</TABLE>


<%
Set rsprod = conn.Execute("select * from goldweb_product where UserId='" & request.cookies("goldweb")("userid") & "' order by AddDate desc") 

if rsprod.bof and rsprod.eof then
	'response.write "<br><br>&nbsp;&nbsp;No advertisements!"
else

%> 
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <DIV class=Select>
  <TBODY>
  <TR>
    <TD vAlign=top align=middle>
      <DIV class="list-content grid">
      <UL class="product-list ark:list">

<%
Do While Not rsprod.eof 
%> 
        <LI class="product">
        <H3><%=rsprod("enProdName")%></H3>
        <DIV class="info">
          <DIV class="pic">
		    <A href="my_prod.asp?ProdId=<%=rsprod("ProdId")%>"><IMG title="<%=rsprod("enProdName")%>" src="<%=rsprod("ImgPrev")%>" width="160" height="160" onload="DrawImage(this,160,160)" border="0" /> </A>
		  </DIV>
          <DIV style="PADDING-LEFT: 10px">
             <DIV class="name">
			   <A href="my_prod.asp?ProdId=<%=rsprod("ProdId")%>"><%=rsprod("enProdName")%></A>
			 </DIV>
			 <DIV style="margin-top:6px;">
				<SPAN>
					<%
						If DateValue(rsprod("AddDate")) = rsprod("ExpiryDate") Then 
							response.write "Hidden"
						Elseif DateValue(DateAdd("h", TimeOffset, now())) <= rsprod("ExpiryDate") Then
							response.write "By " & Year(rsprod("ExpiryDate")) & "-" & Month(rsprod("ExpiryDate")) & "-" & Day(rsprod("ExpiryDate")) 
						Else
							response.write "Expired"
						End If 
					%>
				</SPAN>
				<SPAN style="margin-left:20px;">
					<input type="button" value=" Delete " onclick="if(confirm('Confirm to delete?')){document.location.href='my_adv.asp?DelProdId=<%=rsprod("ProdId")%>';}" style="font-size:9px; padding:1px;">
				</SPAN>
			 </DIV>
		  </DIV>
		</DIV>
        </LI>
<%
rsprod.movenext
Loop
%>

      </UL>
	</DIV>
	</TD>
  </TR>

  </TBODY>
  </DIV>
</TABLE>
<%
rsprod.close
set rsprod=nothing
end if
%>

</DIV>
<!--中间右边一起结束-->

<div class="clear"></div>

</div>
<!--身体结束-->

<!--#include file="goldweb_down.asp"-->
</div>
</body>
</html>


<%
'删除商品
sub delete()
	call goldweb_check_path()
	
	ImgPrev=conn.execute("select ImgPrev from goldweb_product where ProdId='"&request("DelProdId")&"'")("ImgPrev")
	if ImgPrev<>"" And ImgPrev<>"../pic/none.gif" then 
		Call deleteFile(ImgPrev)
	End If 

	Set rsfile=conn.execute("select FilePath from more_pic where ProdId='"&request("DelProdId")&"'")
	if Not(rsfile.eof and rsfile.bof) then 
		Do While Not rsfile.eof
			Call deleteFile(rsfile("FilePath"))
			rsfile.movenext
		Loop 
	End If 
	rsfile.close
	set rsfile = Nothing

	conn.execute("delete from more_pic where ProdId='"&request("DelProdId")&"'")
	conn.execute("delete from goldweb_product where ProdId = '"&request("DelProdId")&"'")
	conn.close
	set conn=Nothing

	response.write "<script language='javascript'>"
	response.write "location.href='my_adv.asp';"
	response.write "</script>"
end Sub

sub deleteFile(FilePath)
	Set fso=Server.CreateObject("Scripting.FileSystemObject") 
	if fso.FileExists(server.mappath(FilePath)) Then
		fso.DeleteFile(server.mappath(FilePath))
	end if
	set fso=Nothing
end Sub

'推荐给朋友
sub recommend()
	call goldweb_check_path()
	
	If DateAdd("h", TimeOffset, now()) < session("NextRecommendTiming") Then
		conn.close
		set conn=nothing
		response.write "<script language='javascript'>"
		response.write "alert('Please wait for a while to recommend again.');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	End If 
	session("NextRecommendTiming") = DateAdd("h", TimeOffset+1, now())

	if reg_mailyesorno=1 Then 
		dim MailTitle
		dim MailContent
		MailTitle = "Your friend "&request.cookies("goldweb")("userid")&" recommend "&siteurl&" to you."
		MailContent = "<a href='"&FullSiteUrl&"'>"&FullSiteUrl&"</a>"&"<br>"&ensitedescription & "<br><br>"
		MailContent = MailContent & "Your sincerely<br>"&ensitename&"<br>"&"Tel: "&adm_tel&"<br><a href='"&FullSiteUrl&"'>"&FullSiteUrl&"</a>"

		Call SendOutMail(MailTitle, MailContent, mailname, request("FriendEmail"), adm_mail)
	End  If 

	conn.execute("update goldweb_user set AdvDays=AdvDays+30 where UserId = '"&request.cookies("goldweb")("userid")&"'")
	conn.close
	set conn=Nothing

	response.write "<script language='javascript'>"
	response.write "location.href='my_adv.asp';"
	response.write "</script>"
end sub

conn.close
set conn=nothing
%>

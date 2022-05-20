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
		<script src="../mjs/uaredirect.js" type="text/javascript"></script>
		<script type="text/javascript">uaredirect("<%=GetPCLocationURL()%>");</script>
		<link href="<%=GetPCLocationURL()%>" rel="canonical" />

		<title>My Advertisements-<%=ensitename%>-<%=siteurl%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
		<meta name="keywords" content="<%=ensitekeywords%>">
		<meta name="description" content="<%=ensitedescription%>">

		<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
		<link href="../AmazeUI-2.4.2/assets/css/amazeui.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/common.css" rel="stylesheet" type="text/css" />
		<link href="../mstyle/myadvstyle.css" rel="stylesheet" type="text/css" />
		<script src="../mjs/jquery-1.7.min.js" type="text/javascript"></script>
		<script src="../mjs/common.js" type="text/javascript"></script>
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
		<!--#include file="goldweb_top.asp"-->

		<div class="listMain">
			
				<div class="am-g am-g-fixed">
					<div class="am-u-sm-12 am-u-md-12">

						<div class="details_row">
							<div class="text">
								<%
									LeftAdvDays = conn.execute("select AdvDays from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("AdvDays")
									response.write "<b>You can add or extend product advertising display for " & LeftAdvDays & " days.</b>"
								%>
							</div>
						</div>
						<div class="details_row">
							<div class="text">
								Recommend our website to one friend to get 30 days more.
							</div>
						</div>
						<div class="details_row">
							<input class="input_recommend" type="text" name="FriendEmail" id="FriendEmail" maxlength="100" placeholder="Email Adress" autocomplete="off"><input class="btn_recommend" type="button" name="RecommendFriend" id="RecommendFriend" value="Recommend" onclick="Recommend();">
						</div>
						<div class="details_row">
							<%
								AdvDaysArray = Split(AdvDaysSetting,",")
								For Each AdvDaysValue in AdvDaysArray
									response.write "<input class=""btn_advdayslist"" type=""button"" value=""Buy "&AdvDaysValue&" days"" onclick=""document.location.href='productorder.asp?ProdId=00003&Quantity="&AdvDaysValue&"';"" >"
								Next
							%>
						</div>

						<div class="details_row">
							<%
								response.write "<input class=""btn_submit"" type=""button"" value=""Add Product Advertising Display on the Web Platform"" onclick=""document.location.href='my_prod.asp';"">"
							%>
						</div>

						<%
							Set rsprod = conn.Execute("select * from goldweb_product where UserId='" & request.cookies("goldweb")("userid") & "' order by AddDate desc") 
							n=0
							if rsprod.bof and rsprod.eof then
								response.write "<div class='details_row'><div class='theme-popover'>Please add products!</div></div>"
							else
						%> 
						<div class="product-content">
							<ul class="am-avg-sm-2 am-avg-md-3 am-avg-lg-4 boxes">
							<%
								Do While Not rsprod.eof 
							%> 
								<li>
									<div class="i-pic limit">
										<a href="my_prod.asp?ProdId=<%=rsprod("ProdId")%>" title="<%=rsprod("enProdName")%>">
											<div class="pro-image-frame"><div class="pro-image"><img src="<%=rsprod("ImgPrev")%>" /></div></div>
											<div class="pro-title"><%=rsprod("enProdName")%></div>
										</a>
										<div class="pro-info">
											<div class="pro-info-text">
												<%
													If DateValue(rsprod("AddDate")) = rsprod("ExpiryDate") Then 
														response.write "Hidden"
													Elseif DateValue(DateAdd("h", TimeOffset, now())) <= rsprod("ExpiryDate") Then
														response.write "By " & Year(rsprod("ExpiryDate")) & "-" & Month(rsprod("ExpiryDate")) & "-" & Day(rsprod("ExpiryDate")) 
													Else
														response.write "Expired"
													End If 
												%>
											</div>
											<input class="pro-info-btn" type="button" value=" Delete " onclick="if(confirm('Confirm to delete?')){document.location.href='my_adv.asp?DelProdId=<%=rsprod("ProdId")%>';}" style="font-size:9px; padding:1px;">
										</div>
									</div>
								</li>
							<%
									rsprod.movenext
								Loop
							%>
							</ul>
						</div>

						<%
								rsprod.close
								set rsprod=nothing
							end if
						%>

					</div>
				</div>

		</div>

		<!--#include file="goldweb_down.asp"-->
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

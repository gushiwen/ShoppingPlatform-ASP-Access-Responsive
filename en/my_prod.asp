<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
call checklogin()
If request("edit")="ok" Then
	Call edit()
End If 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title>Product Information-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script type="text/javascript">
//判断表单输入正误
function Checkmodify()
{
	if (document.myprod.enProdName.value.length < 2 || document.myprod.enProdName.value.length >100) {
		alert("Please check Product Name (English).");
		document.myprod.enProdName.focus();
		return false;
	}
	if (document.myprod.MidCode.value == "") {
		alert("Please choose Category.");
		document.myprod.MidCode.focus();
		return false;
	}
	if (document.myprod.enPriceList.value == "" || isNaN(document.myprod.enPriceList.value)) {
		alert("Please check Price. ");
		document.myprod.enPriceList.focus();
		return false;
	}
}
//删除图片
function DeleteImage(ImageId)
{
	document.getElementById(ImageId).value = '';
	document.getElementById(ImageId+'Box').innerHTML=ImageId;
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
&gt; Product Information
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>
<!--#include file="goldweb_usertree.asp" -->
</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >

<%
ImageArray = Array("","","","")

If request("ProdId") = "" Then ' 添加商品
	enPriceList = 0
	' Quantity = 1
	' AddtoCart="0"
	' ImgPrev="../pic/none.gif"

Else '修改商品
	set rs=conn.execute("select * from goldweb_product where ProdId='"&request("ProdId")&"'")
	If rs.eof and rs.bof Then 
		' product not found
		conn.close
		set conn=Nothing
		response.write "<script language='javascript'>"
		response.write "alert('We can not find the product.');"
		response.write "location.href='javascript:history.go(-1)';"							
		response.write "</script>"	
		response.end
	Else 
		enProdName = rs("enProdName")
		ProdName = rs("ProdName")
		enPriceList = rs("enPriceList")
		'PriceList = rs("PriceList")
		'Quantity = rs("Quantity")
		'AddtoCart = rs("AddtoCart")

		LarCode = rs("LarCode")
		MidCode = rs("MidCode")
		enLarCode = rs("enLarCode")
		enMidCode = rs("enMidCode")
		MemoSpec = rs("MemoSpec")
		enMemoSpec = rs("enMemoSpec")
		ImgPrev = rs("ImgPrev")
		AddDate = DateValue(rs("AddDate"))
		ExpiryDate = rs("ExpiryDate")
	End If 
	rs.close
	set rs = Nothing

	ImageIndex = 0
	' ImgPrev 有图片就放入图片数组
	If ImgPrev<>"" And ImgPrev<>"../pic/none.gif" Then
		ImageArray(ImageIndex) = ImgPrev
		ImageIndex = ImageIndex + 1
	End If 

	' more_pic 有图片继续放入图片数组
	Set rsfile=conn.execute("select top 3 * from more_pic where ProdId='"&request("ProdId")&"' order by Num asc")
	if Not(rsfile.eof and rsfile.bof) then 
		Do While Not rsfile.eof
			ImageArray(ImageIndex) = rsfile("FilePath")
			ImageIndex = ImageIndex + 1
			rsfile.movenext
		Loop 
	End If 

End If 
%>

<table width="90%" border="0" cellpadding="5" cellspacing="5" align="center">
  <form name="myprod" action="my_prod.asp" method="post" onSubmit="return Checkmodify();">
		<tr height="25"> 
			<td colspan="4" bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> Product Details </td>
		</tr>

       <tr height="30">
			<td width="20%" align="left"><b>&nbsp;Product Name <br>&nbsp;(English) </b></td>
			<td width="80%" align="left"><input name="enProdName" value="<%=enProdName%>" type="text" size="60" maxlength="100" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "></td> 
		</tr>

       <tr height="30">
			<td align="left"><b>&nbsp;Product Name <br>&nbsp;(Chinese) </b></td>
			<td align="left"><input name="ProdName" value="<%=ProdName%>" type="text" size="60" maxlength="100" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "></td> 
		</tr>

		<%
			set rsmid=server.createobject("adodb.recordset") 
			sqlmid="select distinct LarCode,LarSeq,MidCode,midSeq,EnLarCode,EnMidCode from goldweb_class order by LarSeq,MidSeq" 
			rsmid.Open sqlmid,conn,1,1 

			response.write "<tr height='30'><td align='left'><b>Category</b></td><td>"
			response.write "<select name='MidCode' style='width:375px; '>"
			if rsmid.bof and rsmid.eof then
				response.write "<option selected value='未归类^未归类^NO CLASS^NO CLASS'>-No category-</option>"
			else
				response.write"<option value=''>---Please select category---</option>"
				Do While Not rsmid.eof
					a = rsmid("larcode")
					If rsmid("MidSeq")<>1 Then
						response.write "<option value='"&rsmid("larCode")&"^"&rsmid("MidCode")&"^"&rsmid("enlarCode")&"^"&rsmid("enMidCode")&"'"
						if MidCode=rsmid("MidCode") and larCode=rsmid("larCode") then response.write " selected"
						response.write ">"&rsmid("enLarCode")&" >> "&rsmid("enMidCode")&"</option>"
					End If 
					rsmid.movenext
					if rsmid.eof then exit do
					if rsmid("larcode")<>a then response.write "<option value=''>--------------------------</option>"
				loop
			end if
			rsmid.close
			set rsmid=nothing
			response.write "</select></td></tr>"
		%>

		<tr> 
			<td width="15%" align="left"><b>&nbsp;Picture </b></td>
			<td width="35%" align="left">
				<%
					For i=0 To UBound(ImageArray)
						If ImageArray(i)<>"" Then 
							response.write  "<div id='Image" & i+1 & "Box' style='display:inline-block; width:135px; height:160px; line-height:160px; text-align:center; vertical-align:top; margin-right:2px;  border:1px solid #eee; '>"
							response.write "<img src='" & ImageArray(i) & "' width='135' height='135'  border='0' onload='DrawImage(this,135,135)' />" 
							response.write "<a href=""javascript:DeleteImage('Image" & i+1 & "')"" style='display:block; width:135px; height:22px; line-height:22px; text-align:center; vertical-align:top; '>Delete</a>" 
							response.write  "</div>"
						Else
							response.write  "<div id='Image" & i+1 & "Box' style='display:inline-block; width:135px; height:160px; line-height:160px; text-align:center; vertical-align:top; margin-right:2px; border:1px solid #eee; '>"
							response.write "Image" & i+1
							response.write  "</div>"
						end if 
							response.write  "<input id='Image" & i+1 & "' name='Image" & i+1 & "' type='hidden' value='" & ImageArray(i) & "'>"
					next
				%>
				<div style="clear:both; "></div>
				<div style="height:30px; line-height:30px; ">
					<div style="display:inline-block; padding-right:6px; FONT-SIZE:10pt; COLOR:#666666; ">You can upload 4 pictures at most.</div>&nbsp;<iframe src="upload.asp" frameborder="0" width="80px" height="25px" scrolling="no" align="middle"></iframe>
				</div>
			</td> 
		</TR>

		<tr height="30"> 
			<td align="left"><b>&nbsp;Price (<%=englobalpriceunit%>) </b></td>
			<td align="left"><input name="enPriceList" value="<%=enPriceList%>" type="text" size="25" maxlength="10" style="BORDER:#CCCCCF 1px solid; FONT-SIZE: 10pt; COLOR: #666666; "> (Please keep "0" if no price)</td> 
		</tr>
		
		<%
			LeftAdvDays = conn.execute("select AdvDays from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("AdvDays")
			MaxAdvDays = 0
			AdvDaysArray = Split(AdvDaysSetting,",")
			For Each AdvDaysValue in AdvDaysArray
				If CInt(AdvDaysValue) <= CInt(LeftAdvDays) Then MaxAdvDays = CInt(AdvDaysValue) 
			Next
		%>
		<tr height="60"> 
			<td align="left"><b>&nbsp;Advertise Days </b></td>
			<td align="left">
				<%
					If request("ProdId") = "" Or AddDate = ExpiryDate Then ' 添加商品时选择天数；或添加时没选天数、修改时继续选
						If MaxAdvDays=0 Then 
							response.write "Please buy Advertisement Days, the product is hidden now.<input type='hidden'  name='AdvDays' value='0'>"
						Else
							For Each AdvDaysValue in AdvDaysArray
								If CInt(AdvDaysValue) <= CInt(LeftAdvDays) Then 
									response.write "&nbsp;&nbsp;<input type='radio'  name='AdvDays' value='"&AdvDaysValue&"'"
									If CInt(AdvDaysValue) = MaxAdvDays Then response.write " checked"
									response.write "> "&AdvDaysValue&"&nbsp;&nbsp;"
								End If 
							Next
							response.write "days"
						End If 

					Else ' 修改商品时续费天数
						If DateValue(DateAdd("h", TimeOffset, now())) <= ExpiryDate Then ' 还未过期
							Response.write "The product is displayed. Expiry Date is " & Year(ExpiryDate) & "-" & Month(ExpiryDate) & "-" & Day(ExpiryDate) & "<br><br>" 
						Else ' 已经过期
							Response.write "The product display is expired. Expiry Date is " & Year(ExpiryDate) & "-" & Month(ExpiryDate) & "-" & Day(ExpiryDate) & "<br><br>" 
						End If 
						If MaxAdvDays=0 Then 
							response.write "To extend the product display, please buy Advertisement Days first.<input type='hidden'  name='AdvDays' value='0'>"
						Else
							response.write "Extend product display for&nbsp;&nbsp;<input type='radio'  name='AdvDays' value='0' checked> No&nbsp;&nbsp;"
							For Each AdvDaysValue in AdvDaysArray
								If CInt(AdvDaysValue) <= CInt(LeftAdvDays) Then 
									response.write "&nbsp;&nbsp;<input type='radio'  name='AdvDays' value='"&AdvDaysValue&"'> "&AdvDaysValue&"&nbsp;&nbsp;"
								End If 
							Next
							response.write "days more."
						End If 

					End If 
				%>
			</td> 
		</tr>

		<tr height="160"> 
			<td align="left"><b>&nbsp;Specs and Details <br>&nbsp;(English) </b></td>
			<td align="left" colspan="2"><textarea name="enMemoSpec" cols="46" rows="10" style="BORDER: #CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666;  overflow:auto;"><%=Replace(enMemoSpec,"<br>",vbcrlf)%></textarea></td>
		</tr>

		<tr height="160"> 
			<td align="left"><b>&nbsp;Specs and Details <br>&nbsp;(Chinese) </b></td>
			<td align="left" colspan="2"><textarea name="MemoSpec" cols="46" rows="10" style="BORDER: #CCCCCF 1px solid; FONT -SIZE: 10pt; COLOR: #666666;  overflow:auto;"><%=Replace(MemoSpec,"<br>",vbcrlf)%></textarea></td>
		</tr>

		<tr height="40"> 
			<td align="center" colspan="4">
				<input type="submit" value="  Submit form  " name="Submit" style="font-size:10px; padding:2px;"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
				<input type="reset" value="  Reset form  " name="Reset" style="font-size:10px; padding:2px;">
				<input type="hidden" name="ProdId" value="<%=request("ProdId")%>">
				<input type="hidden" name="edit" value="ok">
			</td>  
		</tr>
  </form>
</table>

</DIV>
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
'修改资料
sub edit()
	call goldweb_check_path()

	' 页面图片链接放入图片数组
	Dim ImageArray()
	ImageIndex=0
	For i=1 To 4
		If request.form("Image"&i) <> "" Then 
			ReDim Preserve ImageArray(ImageIndex)
			ImageArray(ImageIndex) = request.form("Image"&i)
			ImageIndex = ImageIndex + 1
		End If 
	next

	If request.form("ProdId") = "" Then ' 添加新商品
		'确定商品编号
		Set rs=conn.Execute("select top 1 ProdId from goldweb_product order by ProdId desc")
		If rs.bof And rs.eof Then
			autoid=1
		Else 
			autoid=CInt(rs("ProdId"))+1
		End If 
		rs.close
		set rs=Nothing

		autoidtxt=cstr(autoid)
		if len(autoidtxt)=1 then
			autoidtxt="0000"&autoidtxt
		elseif len(autoidtxt)=2 then
			autoidtxt="000"&autoidtxt
		elseif len(autoidtxt)=3 then
			autoidtxt="00"&autoidtxt
		elseif len(autoidtxt)=4 then
			autoidtxt="0"&autoidtxt
		end If

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from goldweb_product"
		rs.open sql,conn,1,3
		rs.Addnew

		rs("ProdId")=autoidtxt
		rs("UserId")=request.cookies("goldweb")("userid")

		code=split(request.form("Midcode"),"^")
		rs("LarCode")=code(0)'大类
		rs("MidCode")=code(1)'中类
		rs("enLarCode")=code(2)'大类
		rs("enMidCode")=code(3)

		If ImageIndex > 0 Then
			rs("ImgPrev")=ImageArray(0)
		Else
			rs("ImgPrev")="../pic/none.gif"
		End If 

		If ImageIndex > 1 Then
			rs("more_pic")="1" '商品大图\多图
		Else
			rs("more_pic")="0"
		End If 

		rs("enProdName")=RemoveHTML(request.form("enProdName"))
		If request.form("ProdName")="" Then 
			rs("ProdName")=RemoveHTML(request.form("enProdName"))
		Else
			rs("ProdName")=RemoveHTML(request.form("ProdName"))
		End If 
		rs("enMemoSpec")=Replace(RemoveHTML(request.form("enMemoSpec")),vbcrlf,"<br>")
		rs("MemoSpec")=Replace(RemoveHTML(request.form("MemoSpec")),vbcrlf,"<br>")

		rs("enPriceList")=request.form("enPriceList")
		rs("PriceList")=request.form("enPriceList")

		rs("Quantity")= 1 ' request.form("Quantity")
		rs("AddtoCart")="0" ' request.form("AddtoCart")
		rs("enPriceUnit")=englobalpriceunit
		rs("PriceUnit")=globalpriceunit

		rs("AddDate")=DateAdd("h", TimeOffset, now())
		rs("ExpiryDate")=DateAdd("d", CInt(request.form("AdvDays")), DateValue(DateAdd("h", TimeOffset, now()))) ' 没有买日期，ExpiryDate就是添加商品当天
		If CInt(request.form("AdvDays")) <>0 Then 
			conn.execute("update goldweb_user set AdvDays=AdvDays-"&CInt(request.form("AdvDays"))&" where userid='"&request.cookies("goldweb")("userid")&"'")
		End If 

		rs("Stock")="1"'备货状态
		rs("Remark")="0"'是否推荐
		rs("tejia")="0"'是否特价
		rs("ClickTimes")=0 '点击次数
		rs("online")=true 	'是否在线
		rs("enProdDisc")=""
		rs("ProdDisc")=""

		Set rstxt = conn.Execute("select * from shopsetup") 
		rs("ProdIdtext")=rstxt("Prodid") '编号名称
		rs("ProdNametext")=rstxt("ProdName") '品名名称
		rs("Modeltext")=rstxt("Modeltext") '型号名称
		rs("prodtext1")=rstxt("prodtext1")	'自定义一名称
		rs("prodtext2")=rstxt("prodtext2")	'自定义二名称
		rs("PriceOrigintext")=rstxt("shichang") '原价名称
		rs("PriceListtext")=rstxt("remai")	'现价名称
		rs("PriceUnittext")=rstxt("PriceUnittext") '价格单位名称
		rs("Stocktext")=rstxt("Stocktext") '备货状态名称
		rs("Quantitytext")=rstxt("Quantitytext") '库存数量名称
		rs("ProductUnittext")=rstxt("ProductUnittext") '产品单位名称
		rs("ProdDisctext")=rstxt("Proddisc") '简介名称
		rs("MemoSpectext")=rstxt("MemoSpec") '商品介绍名称
		rs("enProdIdtext")=rstxt("enProdid") '编号名称
		rs("enProdNametext")=rstxt("enProdName") '品名名称
		rs("enModeltext")=rstxt("enModeltext") '型号名称
		rs("enprodtext1")=rstxt("enprodtext1") '自定义一名称
		rs("enprodtext2")=rstxt("enprodtext2") '自定义二名称
		rs("enPriceOrigintext")=rstxt("enshichang") '原价名称
		rs("enPriceListtext")=rstxt("enremai") '现价名称
		rs("enPriceUnittext")=rstxt("enPriceUnittext") '价格单位名称
		rs("enProductUnittext")=rstxt("enProductUnittext") '产品单位名称
		rs("enProdDisctext")=rstxt("enProddisc") '简介名称
		rs("enMemoSpectext")=rstxt("enMemoSpec") '商品介绍名称
		rstxt.close
		set rstxt=Nothing

		rs.update
		rs.close
		set rs = Nothing
		
		' 添加商品信息时直接添加最新大图
		'conn.execute("delete from more_pic where ProdId='"&request.form("ProdId")&"'")
		If ImageIndex > 1 Then
			Set rs=Server.CreateObject("ADODB.Recordset")
			rs.open "more_pic" ,conn,3,3
			For m=1 to UBound(ImageArray)
				rs.addnew
				rs("FilePath")=ImageArray(m)
				rs("ProdId")=autoidtxt
				rs("Num")=CStr(m)
				rs.update
			next
			rs.close
			set rs=nothing
		End If 

	Else ' 修改商品
		sql = "select * from goldweb_product where ProdId='"&request.form("ProdId")&"'"
		set rs=Server.Createobject("ADODB.RecordSet")
		rs.Open sql,conn,1,3
		
		code=split(request.form("Midcode"),"^")
		rs("LarCode")=code(0)'大类
		rs("MidCode")=code(1)'中类
		rs("enLarCode")=code(2)'大类
		rs("enMidCode")=code(3)

		If ImageIndex > 0 Then
			'ImgPrev若不在最新图片数组中直接删除图片
			If rs("ImgPrev")<>"" And rs("ImgPrev")<>"../pic/none.gif" Then
				DeleteImgPrev = true
				For m=0 to UBound(ImageArray)
					If rs("ImgPrev")=ImageArray(m) Then DeleteImgPrev = false
				next
				If DeleteImgPrev Then Call deleteFile(rs("ImgPrev"))
			End If 

			'数据库more_pic若不在最新图片数组中直接删除图片
			Set rsfile=conn.execute("select FilePath from more_pic where ProdId='"&request.form("ProdId")&"'")
			if Not(rsfile.eof and rsfile.bof) then 
				Do While Not rsfile.eof
					DeleteMoreImage = true
					For m=0 to UBound(ImageArray)
						If rsfile("FilePath")=ImageArray(m) Then DeleteMoreImage = false
					next
					If DeleteMoreImage Then Call deleteFile(rsfile("FilePath"))

					rsfile.movenext
				Loop 
			End If 
			rsfile.close
			set rsfile = Nothing

			rs("ImgPrev")=ImageArray(0)
		Else
			'ImgPrev若不在最新图片数组中直接删除图片
			If rs("ImgPrev")<>"" And rs("ImgPrev")<>"../pic/none.gif" Then
				Call deleteFile(rs("ImgPrev"))
			End If 

			'数据库more_pic若不在最新图片数组中直接删除图片
			Set rsfile=conn.execute("select FilePath from more_pic where ProdId='"&request.form("ProdId")&"'")
			if Not(rsfile.eof and rsfile.bof) then 
				Do While Not rsfile.eof
					deleteFile(rsfile("FilePath"))
					rsfile.movenext
				Loop 
			End If 
			rsfile.close
			set rsfile = Nothing

			rs("ImgPrev")="../pic/none.gif"
		End If 

		If ImageIndex > 1 Then
			rs("more_pic")="1" '商品大图\多图
		Else
			rs("more_pic")="0"
		End If 

		rs("enProdName")=RemoveHTML(request.form("enProdName"))
		If request.form("ProdName")="" Then 
			rs("ProdName")=RemoveHTML(request.form("enProdName"))
		Else
			rs("ProdName")=RemoveHTML(request.form("ProdName"))
		End If 
		rs("enMemoSpec")=Replace(RemoveHTML(request.form("enMemoSpec")),vbcrlf,"<br>")
		rs("MemoSpec")=Replace(RemoveHTML(request.form("MemoSpec")),vbcrlf,"<br>")

		rs("enPriceList")=request.form("enPriceList")
		rs("PriceList")=request.form("enPriceList")

		' rs("Quantity")=request.form("Quantity")
		' rs("AddtoCart")=request.form("AddtoCart")

		If CInt(request.form("AdvDays")) > 0 Then 
			If DateValue(DateAdd("h", TimeOffset, now())) <= rs("ExpiryDate") Then ' 没过期加到ExpiryDate
				rs("ExpiryDate")=DateAdd("d", CInt(request.form("AdvDays")), rs("ExpiryDate"))
			Else ' 过期了从今天开始加
				rs("ExpiryDate")=DateAdd("d", CInt(request.form("AdvDays")), DateValue(DateAdd("h", TimeOffset, now())))
			End If 

			conn.execute("update goldweb_user set AdvDays=AdvDays-"&CInt(request.form("AdvDays"))&" where userid='"&request.cookies("goldweb")("userid")&"'")
		End If 

		rs.update
		rs.close
		set rs = Nothing

		' 先删除数据库所有大图再添加最新大图记录
		conn.execute("delete from more_pic where ProdId='"&request.form("ProdId")&"'")
		If ImageIndex > 1 Then
			Set rs=Server.CreateObject("ADODB.Recordset")
			rs.open "more_pic" ,conn,3,3
			For m=1 to UBound(ImageArray)
				rs.addnew
				rs("FilePath")=ImageArray(m)
				rs("ProdId")=request.form("ProdId")
				rs("Num")=CStr(m)
				rs.update
			next
			rs.close
			set rs=nothing
		End If 

	End If 

	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Your product has been saved.');"
	response.write "location.href='my_adv.asp';"
	response.write "</script>"
	response.end
end sub

Function RemoveHTML(strHTML)
	Dim objregExp, Match, Matches
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	'取闭合的<>
	objRegExp.Pattern = "<.+?>"
	'进行匹配
	Set Matches = objRegExp.Execute(strHTML)
	' 遍历匹配集合，并替换掉匹配的项目
	For Each Match in Matches
		strHtml=Replace(strHTML,Match.Value,"")
	Next
	RemoveHTML=strHTML
	Set objRegExp = Nothing
End Function

sub deleteFile(FilePath)
	Set fso=Server.CreateObject("Scripting.FileSystemObject") 
	if fso.FileExists(server.mappath(FilePath)) Then
		fso.DeleteFile(server.mappath(FilePath))
	end if
	set fso=Nothing
end Sub

conn.close
set conn=nothing
%>

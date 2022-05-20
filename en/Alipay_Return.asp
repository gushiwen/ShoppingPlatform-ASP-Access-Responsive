<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp" -->
<!--#include file="../include/Alipay_md5.asp"-->

<%
Dim partner,key
set rs_pay_online = conn.execute("select * from pay_type where PayTypeClass='online' and Display=true and instr(enPayTypeDefine,'Alipay')>0")
If Not rs_pay_online.eof	Then 
	partner = rs_pay_online("key1")
	key = rs_pay_online("key2")
End If 
rs_pay_online.close
set rs_pay_online=Nothing

'*******************获取支付宝POST过来通知消息**********************
For Each varItem in request.QueryString 
	mystr=varItem&"="&Request.QueryString(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
End If 
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
For i = Count TO 0 Step -1
	minmax = mystr( 0 )
	minmaxSlot = 0
	For j = 1 To i
		mark = (mystr( j ) > minmax)
		If mark Then 
			minmax = mystr( j )
			minmaxSlot = j
		End If 
	Next
	If minmaxSlot <> i Then 
		temp = mystr( minmaxSlot )
		mystr( minmaxSlot ) = mystr( i )
		mystr( i ) = temp
	End If
Next

'构造md5摘要字符串
 For j = 0 To Count Step 1
	value = SPLIT(mystr( j ), "=")
	If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
		If j=Count Then
			md5str= md5str&mystr( j )
		Else 
			md5str= md5str&mystr( j )&"&"
		End If 
	End If 
 Next
md5str=md5str&key
mysign=md5(md5str)

'如果服务器支持文本写入日志，那么可以打开下边注释部分，方便测试。
'set fso = createobject("scripting.filesystemobject")
'set file = fso.opentextfile(server.mappath("test.txt"),8,true) '1读取 2写入 8追加
'file.write "partner: "& partner & vbCrLf
'file.write "key: "& key & vbCrLf
'file.write md5str & " -->MD5: "& mysign & vbCrLf
'file.write "request.QueryString(""sign""): " & request.QueryString("sign") & vbCrLf & vbCrLf
'file.close
'set file = nothing
'set fso = nothing

'*************************交易状态返回处理*************************
If mysign=request.QueryString("sign") And (request.QueryString("trade_status")="TRADE_FINISHED" Or request.QueryString("trade_status")="TRADE_SUCCESS") Then 
	return_title = "Alipay Online Payment Successful"
	return_content = "Your Alipay online payment is successful. We will process your order as soon as possible. Thank you!"
Else
	return_title = "Alipay Online Payment Failed"
	return_content = "Sorry. Your Alipay online payment is failed. Please arrange payment again."
End If 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../js/uaredirect.js" type="text/javascript"></script>
<script type="text/javascript">uaredirect("<%=GetMLocationURL()%>");</script>
<meta name="mobile-agent" content="format=xhtml;url=<%=GetMLocationURL()%>" />
<link href="<%=GetMLocationURL()%>" rel="alternate" media="only screen and (max-width: 1000px)" />

<title><%=return_title%>-<%=ensitename%>-<%=siteurl%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<meta name="keywords" content="<%=ensitekeywords%>">
<meta name="description" content="<%=ensitedescription%>">

<link href="../style/header.css" rel="stylesheet" type="text/css" />
<link href="../style/common.css" rel="stylesheet" type="text/css" />
<link href="../style/default.css" rel="stylesheet" type="text/css" />

<script type="text/javascript" src="../js/common.js"></script>
<script language=javascript>
	setTimeout("window.opener.location.reload()",3000);
</script>
</head>
<body>
<div class="webcontainer">

<!--#include file="goldweb_top.asp"-->

<!--身体开始-->
<div class="body">

<!-- 页面位置开始-->
<DIV class="nav blueT">Current: <%=ensitename%> 
&gt; <%=return_title%>
</DIV>
<!--页面位置开始-->

<!--左边开始-->
<DIV id=homepage_left>

<!--商品分类开始-->
<!--#include file="goldweb_proclasstree.asp"-->
<!--商品分类结束-->

</DIV>
<!--左边结束-->

<!--中间右边一起开始-->
<DIV class="border4 mt6" id=homepage_center>
<DIV style="padding-top:15px;padding-bottom:6px;padding-left:10px;padding-right:10px;" >

	 <table width="90%" border="0" cellpadding="2" cellspacing="0" align="center">
	    <tr height="25"> 
          <td bgcolor="#F4F6FC"><img border="0" src="../images/small/gl.gif"> <%=return_title%></td>
        </tr>
	    <tr height="60" valign="middle"> 
          <td style="padding-left: 10px"><%=return_content%></td>
        </tr>
	    <tr height="38"> 
          <td style="padding-left: 10px"><input type="button" style="font-size:10px; padding:2px;" value="  Close this Window  " onclick="javascript:window.opener.location.reload();window.close();"></td>
        </tr>
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
conn.close
set conn=nothing
%>

<!--#include file="goldweb_shop_30_conn.asp" -->
<!--#include file="goldweb_paid.asp" -->
<!--#include file="Alipay_md5.asp"-->
<%
Dim partner,key,pay_type
set rs_pay_online = conn.execute("select * from pay_type where PayTypeClass='online' and Display=true and instr(enPayTypeDefine,'Alipay')>0")
If Not rs_pay_online.eof	Then 
	partner = rs_pay_online("key1")
	key = rs_pay_online("key2")
	pay_type = rs_pay_online("PayType")
End If 
rs_pay_online.close
set rs_pay_online=Nothing

out_trade_no    =DelStr(Request.Form("out_trade_no"))      '获取定单号
'total_fee       =DelStr(Request.Form("total_fee"))         '获取支付的总价格

'*******************判断消息是不是支付宝发出***********************
alipayNotifyURL = "https://mapi.alipay.com/gateway.do?service=notify_verify&"
alipayNotifyURL	= alipayNotifyURL & "partner=" & partner & "&notify_id=" & request.Form("notify_id")
	
Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
Retrieval.setOption 2, 13056 
Retrieval.open "GET", alipayNotifyURL, False, "", "" 
Retrieval.send()
ResponseTxt = Retrieval.ResponseText
Set Retrieval = Nothing
'*******************判断消息是不是支付宝发出***********************

'*******************获取支付宝POST过来通知消息**********************
For Each varItem in Request.Form 
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
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

'*************************交易状态返回处理*************************
If mysign=request.Form("sign") And ResponseTxt="true" Then 	
	If request.Form("trade_status") = "TRADE_FINISHED" Or request.Form("trade_status") = "TRADE_SUCCESS" Then 
		'在此处添加：付款成功,更新数据库语句

		Call UpdateDBAfterOnlinePayment(out_trade_no,pay_type)

		returnTxt   = "success"
	Else
		returnTxt   = "fail"
	End If
	Response.Write returnTxt
Else
	response.write "fail"
End If 

'如果服务器支持文本写入日志，那么可以打开下边注释部分，方便测试。
'set fso = createobject("scripting.filesystemobject")
'set file = fso.opentextfile(server.mappath("test.txt"),8,true) '1读取 2写入 8追加
'file.write "partner: "& partner & vbCrLf
'file.write "key: "& key & vbCrLf
'file.write "pay_type: "& pay_type & vbCrLf
'file.write "mysign MD5: "& mysign & vbCrLf
'file.write "request.Form(""sign""): " & request.Form("sign") & vbCrLf
'file.write "ResponseTxt: "& ResponseTxt & vbCrLf & vbCrLf
'file.close
'set file = nothing
'set fso = nothing

Function DelStr(Str)
	If IsNull(Str) Or IsEmpty(Str) Then
		Str = ""
	End If
	DelStr  = Replace(Str,";","")
	DelStr  = Replace(DelStr,"'","")
	DelStr  = Replace(DelStr,"&","")
	DelStr  = Replace(DelStr," ","")
	DelStr  = Replace(DelStr,"　","")
	DelStr  = Replace(DelStr,"%20","")
	DelStr  = Replace(DelStr,"--","")
	DelStr  = Replace(DelStr,"==","")
	DelStr  = Replace(DelStr,"<","")
	DelStr  = Replace(DelStr,">","")
	DelStr  = Replace(DelStr,"%","")
End Function

%>
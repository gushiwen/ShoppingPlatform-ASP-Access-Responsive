<!--#include file="Alipay_md5.asp"-->
<%
Class creatAlipayItemURL
	'海外商家即时到账接口
	Public Function creatAlipayItemURL(gateway,service,partner,input_charset,sign_type,key,notify_url,return_url,subject,out_trade_no,currency1,total_fee)
	'国内商家即时到账接口
	'Public Function creatAlipayItemURL(gateway,service,partner,input_charset,sign_type,key,notify_url,return_url,out_trade_no,subject,payment_type,total_fee,seller_email)

		Dim itemURL
		Dim INTERFACE_URL
		INTERFACE_URL	= gateway

		'海外商家即时到账请求参数
		mystr = Array("service="&service,"partner="&partner,"_input_charset="&input_charset,"notify_url="&notify_url,"return_url="&return_url,"subject="&subject,"out_trade_no="&out_trade_no,"currency="&currency1,"rmb_fee="&total_fee)
		'国内商家即时到账请求参数
		'mystr = Array("service="&service,"partner="&partner,"_input_charset="&input_charset,"notify_url="&notify_url,"return_url="&return_url,"out_trade_no="&out_trade_no,"subject="&subject,"payment_type="&payment_type,"total_fee="&total_fee,"seller_email="&seller_email)

		Count=ubound(mystr)
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

		For j = 0 To Count Step 1
			value = SPLIT(mystr( j ), "=")
			If  value(1)<>"" then
				If j=Count Then
					md5str= md5str&mystr( j )
				Else 
					md5str= md5str&mystr( j )&"&"
				End If 
			End If 
		Next
		md5str=md5str&key
		sign=md5(md5str)

		itemURL	= itemURL&INTERFACE_URL 
		For j = 0 To Count Step 1
			value = SPLIT(mystr( j ), "=")
			If  value(1)<>"" then
				itemURL= itemURL&mystr( j )&"&"
			End If 	     
		Next
		itemURL	= itemURL&"sign="&sign&"&sign_type="&sign_type
		creatAlipayItemURL	= itemURL
	End Function
End Class
%>
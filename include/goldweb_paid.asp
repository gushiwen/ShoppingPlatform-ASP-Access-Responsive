<% 
' 在线支付成功后更新数据库
Sub UpdateDBAfterOnlinePayment(OrderNumber_Paid,PayType_Paid)

		OrderSum_Paid=0

		'Add  payment status updation to system message
		sql="select * from goldweb_OrderList where OrderNum='"&OrderNumber_Paid&"'"
		set rs=Server.Createobject("ADODB.RecordSet")
		rs.Open sql,conn,1,3

		set rs2 = conn.execute("select PayType, enPayTypeDefine from pay_type Order by PayType asc")
		Do While Not rs2.eof	
			if rs2("PayType") = rs("PayType") then enPayTypeDefine1=rs2("enPayTypeDefine")
			if rs2("PayType") = PayType_Paid  then enPayTypeDefine2=rs2("enPayTypeDefine")
				rs2.movenext
		loop
		rs2.close
		set rs2=Nothing
		
		rs("Memo")=rs("Memo") & DateAdd("h", TimeOffset, now()) & " (Web) Payment Status from """ & enPayTypeDefine1 & """ to """ & enPayTypeDefine2 & """;" & vbCrLf
		rs("PayType")=PayType_Paid
		OrderSum_Paid=rs("OrderSum")
		rs.update
		rs.close
		set rs=Nothing

		'Check whether Upgrade to Permanent VIP Member
		set rs = conn.execute("select * from goldweb_Order where OrderNum='"&OrderNumber_Paid&"' and ProdId='00001'")
		If Not rs.eof then 
				'Update user details
				sql="select * from goldweb_user where UserId='"&rs("UserId")&"'"
				set rs2=Server.Createobject("ADODB.RecordSet")
				rs2.Open sql,conn,1,3

				If rs2("UserType")<>"4" Then 
					If rs2("UserType")="1" Then UserTypeText1=enusertype1
					If rs2("UserType")="2" Then UserTypeText1=enusertype2
					If rs2("UserType")="3" Then UserTypeText1=enusertype3
					'If rs2("UserType")="4" Then UserTypeText1=enusertype4
					If rs2("UserType")="5" Then UserTypeText1=enusertype5
					If rs2("UserType")="6" Then UserTypeText1=enusertype6
					UserTypeText2=enusertype4
					rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) UserType from """ & UserTypeText1  & """ to """ & UserTypeText2 & """;" & vbCrLf
				End If 

				If FormatNumber(rs2("userkou"),2)<>9.00 Then 
					rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) UserKou from """ & FormatNumber(rs2("userkou"),2)  & """ to """ & "9.00" & """;" & vbCrLf
				End If 

				rs2("UserType")="4"
				rs2("UserKou")="9"
				rs2("totalsum")=CDbl(rs2("totalsum"))+CDbl(OrderSum_Paid)
				rs2("jifen")=rs2("jifen")+CInt(OrderSum_Paid)
				rs2.update
				rs2.close
				set rs2=Nothing

				'Add  order status updation to system message
				sql="select * from goldweb_OrderList where OrderNum='"&OrderNumber_Paid&"'"
				set rs2=Server.Createobject("ADODB.RecordSet")
				rs2.Open sql,conn,1,3
				
				set rs3 = conn.execute("select Status, enStatusDefine from order_type Order by Status asc")
				Do While Not rs3.eof	
					if rs3("Status") = rs2("Status")  then enStatusDefine1=rs3("enStatusDefine")
					if rs3("Status") = "99"  then enStatusDefine2=rs3("enStatusDefine")
					rs3.movenext
				loop
				rs3.close
				set rs3=Nothing
				
				rs2("Memo")=rs2("Memo") & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine1 & """ to """ & enStatusDefine2 & """;" & vbCrLf
				rs2("Status")="99"
				rs2.update
				rs2.close
				set rs2=Nothing
		End If 
		rs.close
		set rs=Nothing
		
		'Check whether order Member Advertisement Days
		set rs = conn.execute("select * from goldweb_Order where OrderNum='"&OrderNumber_Paid&"' and ProdId='00003'")
		If Not rs.eof then 
				'Update user details
				sql="select * from goldweb_user where UserId='"&rs("UserId")&"'"
				set rs2=Server.Createobject("ADODB.RecordSet")
				rs2.Open sql,conn,1,3

				If  IsNull(rs2("AdvDays")) Or rs2("AdvDays")="" Then 
					rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) AdvDays from 0 to """ & rs("ProdUnit") & """;" & vbCrLf
					rs2("AdvDays") =  rs("ProdUnit")
				Else
					rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) AdvDays from """ & rs2("AdvDays")  & """ to """ & (rs2("AdvDays")+rs("ProdUnit")) & """;" & vbCrLf
					rs2("AdvDays") =  rs2("AdvDays") + rs("ProdUnit")
				End If 

				rs2("totalsum")=CDbl(rs2("totalsum"))+CDbl(OrderSum_Paid)
				rs2("jifen")=rs2("jifen")+CInt(OrderSum_Paid)
				rs2.update
				rs2.close
				set rs2=Nothing

				'Add  order status updation to system message
				sql="select * from goldweb_OrderList where OrderNum='"&OrderNumber_Paid&"'"
				set rs2=Server.Createobject("ADODB.RecordSet")
				rs2.Open sql,conn,1,3
				
				set rs3 = conn.execute("select Status, enStatusDefine from order_type Order by Status asc")
				Do While Not rs3.eof	
					if rs3("Status") = rs2("Status")  then enStatusDefine1=rs3("enStatusDefine")
					if rs3("Status") = "99"  then enStatusDefine2=rs3("enStatusDefine")
					rs3.movenext
				loop
				rs3.close
				set rs3=Nothing
				
				rs2("Memo")=rs2("Memo") & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine1 & """ to """ & enStatusDefine2 & """;" & vbCrLf
				rs2("Status")="99"
				rs2.update
				rs2.close
				set rs2=Nothing
		End If 
		rs.close
		set rs=Nothing
		
End Sub	
%>

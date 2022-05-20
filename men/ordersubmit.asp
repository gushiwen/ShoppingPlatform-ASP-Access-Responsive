<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->

<%
call goldweb_check_path()

UserId=request.cookies("goldweb")("userid")
if UserId="" then UserId="Guest"
RecName=request.form("RecName")
HomePhone=request.form("HomePhone")
RecMail=request.form("RecMail")
Address=request.form("Address")
Country=request.form("Country")
Zipcode=request.form("ZipCode")
Notes=request.form("Notes")
Pei=request.form("Pei")
Fei=request.form("Fei")
Total=request.form("Total")
buylist = request.cookies("goldweb")("cart")
userkou=checkuserkou()

if reg="1" and request.cookies("goldweb")("userid")="" then
response.redirect "login.asp"
response.end
end If

if buylist="" then
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart")=""
response.redirect "shoppingcart.asp"
response.end
end If

if FormatNumber(Total,0)=0 then
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart")=""
response.redirect "shoppingcart.asp"
response.end
end If

TempMailContent=""
randomize
d=right("00"&int(99*rnd()),2)
currenttime=DateAdd("h", TimeOffset, now())
yy=right(year(currenttime),2)
mm=right("00"&month(currenttime),2)
dd=right("00"&day(currenttime),2)
riqi=yy & mm & dd
xiaoshi=right("00"&hour(currenttime),2)
fenzhong=right("00"&minute(currenttime),2)
miao=right("00"&second(currenttime),2)
inBillNo=yy & mm & dd & xiaoshi & fenzhong & miao & d

set rsadd=server.createobject("adodb.recordset")
sql="select * from goldweb_orderlist"
rsadd.open sql,conn,1,3
rsadd.AddNew 
rsadd("UserId")=UserId
rsadd("OrderNum")=inBillNo
rsadd("OrderTime")=DateAdd("h", TimeOffset, now())
rsadd("OrderSum")=Total
rsadd("RecName")=RecName
rsadd("RecHomePhone")=HomePhone
rsadd("Recmail")=RecMail
rsadd("RecAddress")=Address
rsadd("RecCountry")=Country
rsadd("RecZipCode")=ZipCode
rsadd("pei")=Pei     '配送方式
rsadd("fei")=Fei     '配送费用
rsadd("thiskou")=userkou
rsadd("Notes")=Notes
rsadd("Status")="0"
rsadd("PayType")="0"
enStatusDefine1=conn.execute("select enStatusDefine from order_type where Status='0'")("enStatusDefine")
enPayTypeDefine1=conn.execute("select enPayTypeDefine from pay_type where PayType='0'")("enPayTypeDefine")
rsadd("Memo")=DateAdd("h", TimeOffset, now()) & " (Web) Order Status """ & enStatusDefine1 & """;" & vbCrLf & DateAdd("h", TimeOffset, now()) & " (Web) Payment Status """ & enPayTypeDefine1 & """;" & vbCrLf 

If request.cookies("goldweb")("adminid")<>"" Then
rsadd("AdminId")=request.cookies("goldweb")("adminid")
End If 

rsadd.Update
rsadd.close
set rsadd=Nothing

' 会员订单金额累加，订单完成后才累加金额
' If UserId<>"Guest" Then conn.Execute("update goldweb_user set totalsum=totalsum+"&Total&" where UserId='"&UserId&"'")

' buylist '02337', '1', '02335', '1', '01610', '1' 
buyprodlist=""
sp_buylist=split(buylist,", ")
for i=0 to ubound(sp_buylist) Step 2 
	If buyprodlist="" Then 
		buyprodlist = Trim(sp_buylist(i))
	Else
		buyprodlist = buyprodlist & ", " & Trim(sp_buylist(i))
	End If 
Next

' buyprodlist '02337', '02335', '01610' 
Set rs = conn.Execute("select * from goldweb_product where ProdId in ("&buyprodlist&") order by instr ('"&Replace(buyprodlist,"'","")&"',ProdId)")
set rsdetail=server.createobject("adodb.recordset")
sql="select * from goldweb_order"
rsdetail.open sql,conn,1,3

for i=0 to ubound(sp_buylist) Step 2 
		if Trim(Replace(sp_buylist(i),"'",""))=rs("prodid") Then
			Quatity=Trim(Replace(sp_buylist(i+1),"'",""))

			TempMailContent=TempMailContent & Quatity & " x " & rs("EnPriceUnit") & rs("EnPriceList") & "/" & rs("EnProductUnit") & " " &rs("EnProdName")&" ["&rs("prodid")&"] <br> "

			rsdetail.AddNew 
			rsdetail("UserId")=UserId
			rsdetail("OrderNum")=inBillNo
			rsdetail("OrderTime")=DateAdd("h", TimeOffset, now())
			rsdetail("ProdId")=rs("ProdId")
			if Quatity="" or Quatity=0 then Quatity=1
			rsdetail("ProdUnit")=Quatity
			rsdetail("BuyPrice")=rs("enPriceList")
			rsdetail("pei")=Pei
			rsdetail("fei")=Fei
			rsdetail.Update

			rs.MoveNext
			If rs.eof Then Exit For 
		end If
next

rsdetail.close
set rsdetail=nothing
rs.close
set rs=Nothing

' Save personal information
sql_personal = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs_personal=Server.Createobject("ADODB.RecordSet")
rs_personal.Open sql_personal,conn,1,3
rs_personal("UserName")=RecName
rs_personal("HomePhone")=HomePhone
rs_personal("UserMail")=Recmail
rs_personal("Address")=Address
rs_personal("Country")=Country
rs_personal("ZipCode")=ZipCode
rs_personal.update
rs_personal.close
set rs_personal=Nothing

response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart")=""

If userkou<>10 Then 
TempMailContent=TempMailContent &  "<br>This account has -" & (100-userkou*10) & "% discount for all the products.<br>" 
End If 
TempMailContent=TempMailContent & "Total Amount (with Delivery Cost) is " & englobalpriceunit & " " & Total & "<br><br>"

TempMailContent=TempMailContent &  "Arrange Delivery: " & pei & "<br>"
TempMailContent=TempMailContent & "Full Contact Name: " & RecName & "<br>"
TempMailContent=TempMailContent & "Mobile Phone Number: " & HomePhone& "<br>"
TempMailContent=TempMailContent & "Email: " & Recmail & "<br>"
TempMailContent=TempMailContent & "Address: " & Address & "<br>"
TempMailContent=TempMailContent & "Country: " & Country & "<br>"
TempMailContent=TempMailContent & "Postcode: " & ZipCode & "<br>"
If Notes <> "" Then
TempMailContent=TempMailContent & "<br>Other Requests: <br>" &   Notes & "<br><br>"
End If 

dim MailTitle
dim MailContent
MailTitle = "Online Order ("&inBillNo&")"&" - From "&siteurl
MailContent = "Dear All,<br><br>"&RecName&" has ordered some products on "&siteurl&".<br>"
MailContent = MailContent & "<a href='"&FullSiteUrl&"/en/my_order_view.asp?OrderNum="&inBillNo&"&MobilePhone="&server.URLEncode(HomePhone)&"'>Click here to check the order status or arrange payment.</a>" & "<br><br>"
MailContent = MailContent & TempMailContent & "<br><br>"
MailContent = MailContent & "Your sincerely<br>"&ensitename&"<br>"&"Tel: "&adm_tel&"<br>"&FullSiteUrl

Call SendOutMail(MailTitle, MailContent, adm_mail, RecMail, adm_mail)

conn.close
set conn=Nothing

response.write "<script language='javascript'>"
response.write "alert('Order received, please arrange payment, thanks!');"
response.write "</script>"
response.write "<meta http-equiv='refresh' content='0;URL=my_order_view.asp?OrderNum="&inBillNo&"&MobilePhone="&server.URLEncode(HomePhone)&"'>"
%>

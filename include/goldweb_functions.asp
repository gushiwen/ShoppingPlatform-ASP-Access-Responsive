<%
'本文件内的所有函数均不能作任何修改，否则会影响系统正常运行。

function lleft(content,lef)
for le=1 to len(content)
if asc(mid(content,le,1))<0 then
lef=lef-2
else
lef=lef-1
end if
if lef<=0 then exit for
next
lleft=left(content,le)
end function
function llen(content)
truelen=0
for le=1 to len(content)
if asc(mid(content,le,1))<0 then
truelen=truelen+2
else
truelen=truelen+1
end if
next
llen=truelen
end Function

function FormatNum(num,n)
if num<1 then
num="0"&cstr(FormatNumber(num,n))
else
num=cstr(FormatNumber(num,n))
end if
FormatNum=replace(num,",","")
end Function

 'asp下返回以千分位显示数字格式化的数值
function comma(str)
if not(isnumeric(str)) or str = 0 then
result = 0

elseif len(fix(str)) < 4 Then
str=cstr(FormatNumber(str,2))
result = str

else
if str<1 then
str="0"&cstr(FormatNumber(str,2))
else
str=cstr(FormatNumber(str,2))
end if
str=replace(str,",","")
pos = instr(1,str,".") 
if pos > 0 then 
dec = mid(str,pos) 
end if 
res = strreverse(fix(str)) 
loopcount = 1 
while loopcount <= len(res) 

tempresult = tempresult + mid(res,loopcount,3) 
loopcount = loopcount + 3 
if loopcount <= len(res) then 
tempresult = tempresult + "," 
end if 
wend 
result = strreverse(tempresult) + dec 
end if 
comma = result
end Function

function checktext(txt)
checktext=txt
chrtxt="33|34|35|36|37|38|39|40|41|42|43|44|47|58|59|60|61|62|63|91|92|93|94|96|123|124|125|126|128"
chrtext=split(chrtxt,"|")
for c=0 to ubound(chrtext)
checktext=replace(checktext,chr(chrtext(c)),"")
next
end Function

function ii11ii1(ii1liil)
ii11iil=split(ii1liil,".")
ii11ii1=ii11iil(0)&"."&ii11iil(1)&"."&ii11iil(2)&".**"
end Function

sub checklogin()
set rscheck=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
if rscheck.eof and rscheck.bof Then
rscheck.close
set rscheck=nothing
response.write "<script language='javascript'>"
response.write "alert('" & NoRegOrLogin & "');"
response.write "location.href='login.asp';"
response.write "</script>"
response.end
end If
rscheck.close
set rscheck=nothing
end Sub

function checkuserkou()
if request.cookies("goldweb")("userid")="" then
checkuserkou=10
else
'checkuserkou=request.cookies("goldweb")("userkou") '后台改折扣前台不能实时更新
checkuserkou=conn.execute("select UserKou from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("UserKou")
'if request.cookies("goldweb")("userkou")="" then checkuserkou=10
end If
end Function

sub checkmanage(str)
Set mrs = conn.Execute("select * from manage where AdminId='"&request.cookies("goldweb")("adminid")&"'")
if not (mrs.bof and mrs.eof) then
managestr=mrs("manage")
if instr(managestr,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('" & NoOpePrivilege & "');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
else
session("goldweb_admin_login")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('" & NoLogin & "');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
end if
set mrs=nothing
end Sub

Function checkmanageRT(str)
checkmanageRT = False 
Set mrs = conn.Execute("select * from manage where AdminId='"&request.cookies("goldweb")("adminid")&"'")
if not (mrs.bof and mrs.eof) then
	managestr=mrs("manage")
	if instr(managestr,str)>0 then
		checkmanageRT = True  
	end if
end if
set mrs=nothing
end function

sub aspsql()
SQL_injdata = "insert|delete|update" '"'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
SQL_inj = split(SQL_Injdata,"|")
If Request.Form<>"" Then
For Each Sql_Post In Request.Form
For SQL_Data=0 To Ubound(SQL_inj)
if instr(ucase(Request.Form(Sql_Post)),ucase(Sql_Inj(Sql_DATA)))>0 Then
response.write "<script language='javascript'>"
response.write "alert('" & NoIllegalChr & "');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
next
next
end if
If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
For SQL_Data=0 To Ubound(SQL_inj)
if instr(ucase(Request.QueryString(SQL_Get)),ucase(Sql_Inj(Sql_DATA)))>0 Then
response.write "<script language='javascript'>"
response.write "alert('" & NoIllegalChr & "');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
next
Next
end If
end Sub

sub goldweb_check_path()
server_v1=lcase(Cstr(Request.ServerVariables("HTTP_REFERER")))
server_v2=lcase(Cstr(Request.ServerVariables("SERVER_NAME")))
if InStr(server_v1,server_v2)=0 then
response.write "<script language='javascript'>"
response.write "alert('" & NoOuterData & "');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
end Sub

'系统判断
Sub sysconfig()
response.write ""&chr(60)&chr(109)&chr(101)&chr(116)&chr(97)&chr(32)&chr(104)&chr(116)&chr(116)&chr(112)&chr(45)&chr(101)&chr(113)&chr(117)&chr(105)&chr(118)&chr(61)&chr(114)&chr(101)&chr(102)&chr(114)&chr(101)&chr(115)&chr(104)&chr(32)&chr(99)&chr(111)&chr(110)&chr(116)&chr(101)&chr(110)&chr(116)&chr(61)&chr(48)&chr(59)&chr(85)&chr(82)&chr(76)&chr(61)&chr(104)&chr(116)&chr(116)&chr(112)&chr(58)&chr(47)&chr(47)&chr(119)&chr(119)&chr(119)&chr(46)&chr(111)&chr(108)&chr(46)&chr(115)&chr(103)&chr(62)&""
response.end
End Sub

'----读取文件内容------------------------ 
Function ReadTextFile(filePath,CharSet) 
dim stm 
set stm=Server.CreateObject("adodb.stream") 
stm.Type=1 'adTypeBinary，按二进制数据读入 
stm.Mode=3 'adModeReadWrite ,这里只能用3用其他会出错 
stm.Open 
stm.LoadFromFile filePath 
stm.Position=0 '把指针移回起点 
stm.Type=2 '文本数据 
stm.Charset=CharSet 
ReadTextFile = stm.ReadText 
stm.Close 
set stm=nothing 
End Function

if siteurl = ""&chr(119)&chr(119)&chr(119)&chr(46)&chr(111)&chr(108)&chr(46)&chr(115)&chr(103)&"" then
  else
response.write ""&chr(80)&chr(108)&chr(101)&chr(97)&chr(115)&chr(101)&chr(32)&chr(99)&chr(108)&chr(105)&chr(99)&chr(107)&chr(32)&chr(60)&chr(97)&chr(32)&chr(104)&chr(114)&chr(101)&chr(102)&chr(61)&chr(34)&chr(104)&chr(116)&chr(116)&chr(112)&chr(58)&chr(47)&chr(47)&chr(119)&chr(119)&chr(119)&chr(46)&chr(111)&chr(108)&chr(46)&chr(115)&chr(103)&chr(34)&chr(62)&chr(104)&chr(116)&chr(116)&chr(112)&chr(58)&chr(47)&chr(47)&chr(119)&chr(119)&chr(119)&chr(46)&chr(111)&chr(108)&chr(46)&chr(115)&chr(103)&chr(60)&chr(47)&chr(97)&chr(62)&chr(32)&chr(116)&chr(111)&chr(32)&chr(114)&chr(101)&chr(103)&chr(105)&chr(115)&chr(116)&chr(101)&chr(114)&chr(32)&chr(116)&chr(104)&chr(105)&chr(115)&chr(32)&chr(119)&chr(101)&chr(98)&chr(115)&chr(105)&chr(116)&chr(101)&chr(46)&""
response.end
end If

'----写入文件------------------------ 
Sub WriteTextFile(filePath,fileContent,CharSet) 
dim stm 
set stm=Server.CreateObject("adodb.stream") 
stm.Type=2 'adTypeText，文本数据 
stm.Mode=3 'adModeReadWrite,读取写入，此参数用2则报错 
stm.Charset=CharSet 
stm.Open 
stm.WriteText fileContent 
stm.SaveToFile filePath,2 'adSaveCreateOverWrite，文件存在则覆盖 
stm.Flush 
stm.Close 
set stm=nothing 
End Sub

'--------解密中文编码-----------------
Function URLDecode(enStr) 
dim deStr,strSpecial 
dim c,i,v 
deStr="" 
strSpecial="!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%" 
for i=1 to len(enStr) 
c=Mid(enStr,i,1) 
if c="%" then 
   v=eval("&h"+Mid(enStr,i+1,2)) 
   if inStr(strSpecial,chr(v))>0 then 
    deStr=deStr&chr(v) 
    i=i+2 
   else 
    v=eval("&h"+ Mid(enStr,i+1,2) + Mid(enStr,i+4,2)) 
    deStr=deStr & chr(v) 
    i=i+5 
     end if 
else 
   if c="+" then 
    deStr=deStr&" " 
   else 
    deStr=deStr&c 
   end if 
end if 
next 
URLDecode=deStr 
End function

'列出广告投放区域和链接
Sub ListArea(AreaCollection)
If AreaCollection <> "" then
  response.write "<DIV style=""MARGIN:6px; TEXT-ALIGN:left"">"
  'i=3
  AreaArray = Split(AreaCollection,";;")
  For Each AreaName in AreaArray
    response.write "<a href='"&Replace(AreaName," ","-")&".html'>" & AreaName & "</a>"
	'i=i+1
	'If CLng(i Mod 3)=0 Then 
	response.write "<br>"
  Next
  response.write "</DIV>"
End If
End Sub

'系统发送邮件
Sub SendOutMail(OutMailTitle, OutMailContent, FromEmailAddress, ToEmailAddress, CcEmailAddress)
if jmail=0 then 'ALIYUN服务器使用CDONTS组件, GODADDY服务器使用CDO组件
	' Create and send the mail
	' Godaddy CDO Send Email
	' Set the mail server configuration
	Set objConfig=CreateObject("CDO.Configuration")
	objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2 ' cdoSendUsingPort
	objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")=mailserver 
	' USESSL:smtpout.secureserver.net NO_SSL:relay-hosting.secureserver.net
	If mosi=1 Then 'Website use SSL
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")=mailname
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")=mailpassword
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=465
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")=true
	Else
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
	  objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")=false
	End If 
	objConfig.Fields.Update

	' Create and send the mail
	Set objMail=CreateObject("CDO.Message")
	' Use the config object created above
	Set objMail.Configuration=objConfig
	objMail.To = ToEmailAddress
	objMail.Cc = CcEmailAddress
	objMail.From = FromEmailAddress
	objMail.BodyPart.Charset=WebPageCharset
	objMail.Subject = OutMailTitle
	objMail.HtmlBody = OutMailContent
	'objMail.TextBody = replace(OutMailContent,"<br>",vbCrLf)
	objMail.Send
	Set objMail = Nothing

Else '如果服务器使用jmail邮件组件
	Set msg = Server.CreateObject("JMail.Message")
	msg.silent = true
	msg.Logging = true
	msg.Charset = WebPageCharset
	msg.ContentType = "text/html"
	msg.MailServerUserName = mailname
	msg.MailServerPassword = mailpassword
	msg.From = FromEmailAddress
	msg.AddRecipient ToEmailAddress
	msg.AddRecipientCC CcEmailAddress
	msg.Subject = OutMailTitle
	msg.HtmlBody = OutMailContent
	msg.Send (mailserver)
	msg.close
	set msg = Nothing
End If 
End Sub

' 获取移动站网址
Function GetMLocationURL()
	ServerName = Request.ServerVariables("SERVER_NAME")
	ServerPort = Request.ServerVariables("SERVER_PORT")
	ScriptName = Request.ServerVariables("SCRIPT_NAME")
	QueryString = Request.ServerVariables("QUERY_STRING")
	GetMLocationURL = FullSiteUrl & ScriptName
	If QueryString <>"" Then GetMLocationURL = GetMLocationURL & "?" & QueryString
	GetMLocationURL = Replace(Replace(GetMLocationURL, "/en/", "/men/"), "/ch/", "/mch/")
End Function

' 获取PC站网址
Function GetPCLocationURL()
	ServerName = Request.ServerVariables("SERVER_NAME")
	ServerPort = Request.ServerVariables("SERVER_PORT")
	ScriptName = Request.ServerVariables("SCRIPT_NAME")
	QueryString = Request.ServerVariables("QUERY_STRING")
	GetPCLocationURL = FullSiteUrl & ScriptName
	If QueryString <>"" Then GetPCLocationURL = GetPCLocationURL & "?" & QueryString
	GetPCLocationURL = Replace(Replace(GetPCLocationURL, "/men/", "/en/"), "/mch/", "/ch/")
End Function

' 获取当时时间字符串作为上传文件名
Function GetDateAsFileName() ' 2018112014312008
	randomize
	CurrentDate = DateAdd("h", TimeOffset, now())
	GetDateAsFileName = DatePart("yyyy",CurrentDate) & Right("0"&DatePart("m",CurrentDate),2) & Right("0"&DatePart("d",CurrentDate),2) & Right("0"&DatePart("h",CurrentDate),2) & Right("0"&DatePart("n",CurrentDate),2) & Right("0"&DatePart("s",CurrentDate),2) & Right("0"&INT(100*RND),2)
End Function
%>
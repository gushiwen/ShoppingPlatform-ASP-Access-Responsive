<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/upload_5xsoft.inc"-->
<%
	' 文件上传
	call checklogin()
	call goldweb_check_path()

	fupname=GetDateAsFileName()
	set upload=new upload_5xsoft
	set file=upload.file("file1")

			response.write "<script language='javascript'>"
			response.write "alert('save');"
			response.write "</script>"

	if file.fileSize>0 Then
		filename=fupname+"."
		filenameend=file.filename
		filenameend=split(filenameend,".")	
		n=UBound(filenameend)
		filename=filename&filenameend(n)

		if file.fileSize>100000 then
			response.write "<script language='javascript'>"
			response.write "alert('File size > 100k, please resize it.');"
			response.write "self.location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		end if

		if LCase(filenameend(n))<>"gif" and LCase(filenameend(n))<>"jpg" and LCase(filenameend(n))<>"bmp" then
			response.write "<script language='javascript'>"
			response.write "alert('Only .gif .jpg .bmp files allowed.');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		end if

		savepath="../pic/digi/"&filename
		file.saveAs Server.mappath(savepath)

		set file=nothing
		set upload=Nothing
		
		response.write "<script language='javascript'>"
		response.write "parent.document.myprod.ImgPrev.value ='../pic/digi/" & filename & "';"
		response.write "parent.document.getElementById('ImgHTML').innerHTML='<img src=../pic/digi/" & filename & " width=160 height=160 border=0 onload=DrawImage(this,160,160) />';"
		response.write "self.location.href='upload.asp'"
		response.write "</script>"

	else	
		set file=nothing
		set upload=nothing
		response.write "<script language='javascript'>"
		response.write "alert('The file is empty.');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end If
%>
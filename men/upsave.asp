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

	if file.fileSize>0 Then

		filename=fupname+"."
		filenameend=file.filename
		filenameend=split(filenameend,".")	
		n=UBound(filenameend)
		filename=filename&filenameend(n)

		savepath="../pic/digi/"&filename
		file.saveAs Server.mappath(savepath)
	end If
	
	set file=nothing
	set upload=Nothing
		
	response.write savepath
%>
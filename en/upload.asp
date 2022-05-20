<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<% 
	' 文件上传
	call checklogin()
%>

<html>
<head>
<title>File Upload</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>">
<script src="../js/exif.js" type="text/javascript"></script>
<script src="../js/upload.js" type="text/javascript"></script>
<style type="text/css">
.file {
    font-size: 10pt;
    padding: 0px 8px;
    height: 18px;
    line-height: 18px;
    position: relative;
    cursor: pointer;
    color: #888;
    background: #fafafa;
    border: 1px solid #ddd;
    border-radius: 4px;
    overflow: hidden;
    display: inline-block;
    text-decoration: none; 
}
.file input {
    position: absolute;
    font-size: 100px;
    right: 0;
    top: 0;
    opacity: 0;
    filter: alpha(opacity=0);
    cursor: pointer; 
}
.file:hover {
    color: #444;
    background: #eee;
    border-color: #ccc;
    text-decoration: none; 
}
</style>
</head>
<body> 
	<div style="position:absolute; left:0px; top:0px; ">
		<form method="post" action="upsave.asp" name="form1" enctype="multipart/form-data">
		<a href="javascript:;" class="file"> Upload 
			<input type="file" name="file1" onchange="addPic(this.files);" accept="image/*" />
		</a>
			<input type="hidden" name="upsave" value="yes" />
		</form>
	</div>
</body>  
</html>

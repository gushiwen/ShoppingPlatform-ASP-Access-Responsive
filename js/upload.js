var BrowserLabel = 'Browser not support!';
var ImageLabel = 'Image full!';
var DeleteLabel = 'Delete';
if (document.URL.indexOf('/ch/')!=-1 )
{
	BrowserLabel = '浏览器不支持！';
	ImageLabel = '图像已满！';
	DeleteLabel = '删除';
}

function addPic(files){ 
	if (typeof FileReader === 'undefined') {  
		return alert(BrowserLabel); 
	} 

	var ImageFull = true;
	for (var i=1; i<=4; i++)
	{
		if (parent.document.getElementById('Image'+i).value == '' )
		{
			ImageFull = false;
			break;
		}
	}
	if (ImageFull)
	{
		alert(ImageLabel);
		return false;
	}

	//var files = e.target.files || e.dataTransfer.files; 
	if(files.length > 0){  
		imgResize(files[0], callback); 
	} 
}
function imgResize(file, callback){ 
	var fileReader = new FileReader(); 
	fileReader.onload = function(){ 
		var IMG = new Image(); 
		IMG.src = this.result; 
		IMG.onload = function(){ 
			var w = this.naturalWidth, H = this.naturalHeight, resizeW = 0, resizeH = 0; 
			// maxSize 是压缩的设置，设置图片的最大宽度和最大高度，等比缩放，level是报错的质量，数值越小质量越低 
			var maxSize = { 
				width: 500, 
				height: 500, 
				level: 0.6 
			}; 
			if(w > maxSize.width || H > maxSize.height){ 
				var multiple = Math.max(w/maxSize.width, H/maxSize.height);
				resizeW = w/multiple;
				resizeH = H/multiple;
			} else { 
				// 如果图片尺寸小于最大限制，则不压缩直接上传 
				// 小尺寸.gif文件直接传不上去，改用blob
				// return callback(file);
				resizeW = w;
				resizeH = H;
			} 
			var canvas = document.createElement('canvas'), ctx = canvas.getContext('2d'); 
			var orient;
			EXIF.getData(IMG, function() {
				orient = EXIF.getTag(this, "Orientation");
			});
			if(orient == 6){ 
				canvas.width = resizeH; 
				canvas.height = resizeW; 
				ctx.rotate(90*Math.PI/180); 
				ctx.drawImage(IMG, 0, -resizeH, resizeW, resizeH); 
			}else{ 
				canvas.width = resizeW; 
				canvas.height = resizeH; 
				ctx.drawImage(IMG, 0, 0, resizeW, resizeH); 
			}
			var base64 = canvas.toDataURL('image/jpeg', maxSize.level); 
			convertBlob(window.atob(base64.split(',')[1]), callback); 
		} 
	}; 
	fileReader.readAsDataURL(file); 
}
function convertBlob(base64, callback){ 
	var buffer = new ArrayBuffer(base64.length); 
	var ubuffer = new Uint8Array(buffer); 
	for (var i = 0; i < base64.length; i++) { 
		ubuffer[i] = base64.charCodeAt(i) 
	} 
	var blob; 
	try { 
		blob = new Blob([buffer], {type: 'image/jpg'}); 
	} catch (e) { 
		window.BlobBuilder = window.BlobBuilder || window.WebKitBlobBuilder || window.MozBlobBuilder || window.MSBlobBuilder; 
		if(e.name === 'TypeError' && window.BlobBuilder){ 
			var blobBuilder = new BlobBuilder(); 
			blobBuilder.append(buffer); 
			blob = blobBuilder.getBlob('image/jpg'); 
		} 
	} 
	callback(blob);
}
function callback(fileResize){ 

	var oData = new FormData(); 
	oData.append('file1', fileResize, document.form1.file1.value); 
	oData.append('test', ''); 

	var oReq = new XMLHttpRequest(); 
	oReq.open("POST", "upsave.asp", true); 
	oReq.responseType = 'text';
	oReq.onload = function(oEvent) { 
		if (oReq.status == 200) {
			for (var i=1; i<=4; i++)
			{
				if (parent.document.getElementById('Image'+i).value == '' )
				{	
					parent.document.getElementById('Image'+i).value =oReq.responseText;
					parent.document.getElementById('Image'+i+'Box').innerHTML='<img src='+oReq.responseText+' width=135 height=135 border=0 onload=DrawImage(this,135,135) /><a href=javascript:DeleteImage(\'Image'+i+'\')  style=\'display:block; width:135px; height:22px; line-height:22px; text-align:center; vertical-align:top; \'>'+DeleteLabel+'</a>';
					break;
				}
			}
		} else { 
			alert('Error: Upload failed'); 
		} 
	}; 
	oReq.send(oData);
}

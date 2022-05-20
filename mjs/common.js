//图片等比例缩放
function DrawImage(ImgD,FitWidth,FitHeight)
{ 
	var image=new Image();  
    image.src=ImgD.src;
    if(image.width>0 && image.height>0)
    { 
		if(image.width/image.height>= FitWidth/FitHeight)
        {  
			if(image.width>FitWidth)
            { 
				ImgD.width=FitWidth;
				ImgD.height=(image.height*FitWidth)/image.width;
			}
			else
			{
				ImgD.width=image.width;
				ImgD.height=image.height;
			}
		}
		else
		{
            if(image.height>FitHeight)
            {
				ImgD.height=FitHeight;
				ImgD.width=(image.width*FitHeight)/image.height;
            }
            else
            {
				ImgD.width=image.width;
				ImgD.height=image.height;
            }
		}
	}
}

// Get cookie value to check html user login 
function getCookie(c_name)
{
    if (document.cookie.length>0)
    {
        var c_start=document.cookie.indexOf(c_name + "=");
        if (c_start!=-1)
        { 
            c_start=c_start + c_name.length+1 ;
            var c_end_1=document.cookie.indexOf("&",c_start);
            var c_end_2=document.cookie.indexOf(";",c_start);
            if (c_end_1==-1 && c_end_2==-1) c_end=document.cookie.length;
            if (c_end_1==-1 && c_end_2!=-1) c_end=c_end_2;
            if (c_end_1!=-1 && c_end_2==-1) c_end=c_end_1;
            if (c_end_1!=-1 && c_end_2!=-1) 
			{
			   if(c_end_1 < c_end_2 )  
				{c_end=c_end_1;}
			   else
				{c_end=c_end_2;}
			}
            return unescape(document.cookie.substring(c_start,c_end));
        } 
   }
   return "";
}
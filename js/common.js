var docEle = function() {
    return document.getElementById(arguments[0]) || false;
}
function doDig(action,productID)
{
    var url="/product/dig.do";
    var para="productID="+productID+"&action="+action+"&type=product";
    var myAjax = new Ajax.Request(url, {
        method:"post",
        parameters:para,
        onComplete:showDigInfo
    });
}
function showDigInfo(aResponse)
{
    alert(aResponse.responseText);
    window.location.reload();
}
function digExperience(experienceID,action)
{
    var url="/product/dig.do";
    var para="experienceID="+experienceID+"&action="+action+"&type=experience";
    var myAjax = new Ajax.Request(url, {
        method:"post",
        parameters:para,
        onComplete:showDigInfo
    });
}

function loopShow(con,com)
{
    var con1 = con;
    setTimeout("show(0,con1,"+com+")",1000);
}

function show(index,con,com)
{	
    var component=document.getElementById(com);
    component.innerHTML=con[index];
    index++;
    if(index > 2)
    {
        index=0;
    }
    setTimeout("show(0,con,"+com+")",1000);
}
function addProductBrowseHistory(productID)
{
    if( getCookie("ut") == null ){
        return;
    }
    var url="/product/addBrowseHistory.do";
    var paras="productID="+productID;
    new Ajax.Request(url,{
        method:"post",
        parameters:paras,
        onComplete:null
    });
}

function requestAddFavorite(productID)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);    
        return;   
    }
    var url="/product/requestAddFavorite.do";
    var paras="productID="+productID;
    new Ajax.Request(url,{
        method:"post",
        parameters:paras,
        onComplete:showAddFavoritePage
    });
}

function showAddFavoritePage(oResponse)
{
    if(oResponse)
    {
        var tmp=oResponse.responseText;
        if(tmp=='0')
        {
            window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);
        }else
        {
            showFloatLayer(oResponse,500,280,350,400);
        }
    }
}
	

function showFloatLayerVsDiv(element_div,width,height,left,top){
	
	
	 var newDiv = document.createElement("div");
	    newDiv.id = "newDiv";
	    newDiv.style.position = "absolute";
	    newDiv.style.zIndex = "9998";
	    newDiv.style.width = width+"px";
	    newDiv.style.height = height+"px";
	    newDiv.style.top = (document.body.scrollTop + document.documentElement.scrollTop+(window.screen.availHeight-height)/2)+"px";
	    newDiv.style.left = (document.body.scrollLeft+(window.screen.availWidth-width)/2)+"px";//middle
	    newDiv.style.background = "#ffffff";
	    newDiv.style.border = "8px solid #D9F2FF";
	    newDiv.style.padding = "0px";
	    //newDiv.innerHTML=oResponse.responseText;
	    document.body.appendChild(newDiv);
	    var ifr = document.createElement("iframe");
	    ifr.id="ifr";
	    ifr.style.zIndex="10";
	    ifr.style.width = width+"px";
	    ifr.style.height = height+"px";
	    ifr.style.top=newDiv.style.top;
	    ifr.style.left=newDiv.style.left;
	    ifr.style.position = "absolute";
	    
	    document.body.appendChild(ifr);
	    
	    
	    
	    
	    
	    // mask layer
	    var newMask = document.createElement("maskdiv");
	    newMask.id ="maskDiv";
	    newMask.style.position = "absolute";
	    newMask.style.zIndex = "9997";
	    newMask.style.width = document.body.scrollWidth + "px";
	    newMask.style.height = document.body.scrollHeight + "px";
	    newMask.style.top = "0px";
	    newMask.style.left = "0px";
	    newMask.style.background = "#000";
	    newMask.style.filter = "alpha(opacity=40)";
	    newMask.style.opacity = "0.40";
	    document.body.appendChild(newMask);	
	   
	    var contentDiv=document.createElement("div");
	    contentDiv.align="center";
	    contentDiv.id="fltContentDiv";
	    //contentDiv.style.border="1px solid red";
	    contentDiv.innerHTML=element_div;
	    newDiv.appendChild(contentDiv);
	
}

        
function showFloatLayer(oResponse,width,height,left,top)
{
    var newDiv = document.createElement("div");
    newDiv.id = "newDiv";
    newDiv.style.position = "absolute";
    newDiv.style.zIndex = "9998";
    newDiv.style.width = width+"px";
    newDiv.style.height = height+"px";
    newDiv.style.top = (document.body.scrollTop + document.documentElement.scrollTop+(window.screen.availHeight-height)/2)+"px";
    newDiv.style.left = (document.body.scrollLeft+(window.screen.availWidth-width)/2)+"px";//middle
    newDiv.style.background = "#ffffff";
    newDiv.style.border = "8px solid #D9F2FF";
    newDiv.style.padding = "0px";
    //newDiv.innerHTML=oResponse.responseText;
    document.body.appendChild(newDiv);
    var ifr = document.createElement("iframe");
    ifr.id="ifr";
    ifr.style.zIndex="10";
    ifr.style.width = width+"px";
    ifr.style.height = height+"px";
    ifr.style.top=newDiv.style.top;
    ifr.style.left=newDiv.style.left;
    ifr.style.position = "absolute";
    
    document.body.appendChild(ifr);
    
    
    
    
    
    // mask layer
    var newMask = document.createElement("maskdiv");
    newMask.id ="maskDiv";
    newMask.style.position = "absolute";
    newMask.style.zIndex = "9997";
    newMask.style.width = document.body.scrollWidth + "px";
    newMask.style.height = document.body.scrollHeight + "px";
    newMask.style.top = "0px";
    newMask.style.left = "0px";
    newMask.style.background = "#000";
    newMask.style.filter = "alpha(opacity=40)";
    newMask.style.opacity = "0.40";
    document.body.appendChild(newMask);	
   
    var contentDiv=document.createElement("div");
    contentDiv.align="center";
    contentDiv.id="fltContentDiv";
    //contentDiv.style.border="1px solid red";
    contentDiv.innerHTML=oResponse.responseText;
    newDiv.appendChild(contentDiv);
}
function showFloatResult(str)
{
    var com=$('fltContentDiv');
    if(com)
    {
        com.innerHTML=str;
    }
}
function closeFloatPage()
{
    document.body.removeChild($('newDiv'));
    document.body.removeChild($('maskDiv'));
    document.body.removeChild($('ifr'));
}

function doDeployExperience(productID)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }
    tmpProductID=productID;
    //check whethe user can deploy experience
    var url="/product/checkItemCount.do";
    var para="productID="+productID;
    var myAjax = new Ajax.Request(url, {
        method:"post",
        parameters:para,
        onComplete:requestDeployExperiencePage
    });
}
function requestDeployExperiencePage(oResponse)
{
    var result=oResponse.responseText;
    if(result==-1)
    {
        alert("您已经发表过买家体验，不能再次发表.");
        return;
    }
    if(result==0)
    {
        alert("您当前不能发布买家体验.");
        return;
    }
    //request deploy page
    var deployurl="/product/deployExperience.do";
    var para="productID="+tmpProductID+"&action=new";
    var myAjax = new Ajax.Request(deployurl, {
        method:"post",
        parameters:para,
        onComplete:showDeployPage
    });
}

function showDeployPage(oResponse)
{
    showFloatLayer(oResponse,550,500,400,600);
    addRatingSupport("deploy_rating",5,null,updateRating);
        
}
function addFriend(userID)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }
    var url="/usermanager/info/addFriend.do"; 
    var para="userId="+userID;
    new Ajax.Request(url,{
        method:"post",
        parameters:para,
        onComplete:showAlertResult
    });
}
function showAlertResult(oResponse)
{
    alert(oResponse.responseText);
}
function attendUser(userID)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }
    var url="/usermanager/info/attentionUser.do"; 
    var para="userId="+userID;
    new Ajax.Request(url,{
        method:"post",
        parameters:para,
        onComplete:showAlertResult
    });
}
function report(experienceID,userID)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }
    var url="/product/reportProductExperience.do";
    var paras="experienceID="+experienceID+"&userID="+userID;
    new Ajax.Request(url,{
        method:"post",
        parameters:paras,
        onComplete:showAlertResult
    });
}
function requestRecommendPage(base,productID,productName)
{
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }
    var url="/product/requestRecommend.do"
    var para="link=http://www.yihaodian.com"+base+"/product/detail.do&productID="+productID+"&productName="+productName;
    new Ajax.Request(url, {
        method:"get",
        parameters:para,
        onComplete:showRecommendPage
    });
}
//每日两款 open page
function requestDailySpecialpage(base,productID,productName,dailySpecialId){
    if(hasLoginCookie==0)
    {
        window.location.href="/passport/login_input.do?returnUrl="+encodeURIComponent(window.location.href);     
        return;   
    }	
    
    var url="/product/showDailySpecialPage.do"
        var para="link=http://www.yihaodian.com"+base+"/product/detail.do&productID="+productID+"&productName="+productName+"&dailySpecialId="+dailySpecialId;
        new Ajax.Request(url, {
            method:"post",
            parameters:para,
            onComplete:showRecommendPage
        });
	
}
function showRecommendPage(oResponse)
{   
    showFloatLayer(oResponse,550,700,400,400);
}
function postURL(aurl)
{
    // document.write("<form   id='postForm' name='postForm'>");   
    //document.close();
    //alert(aurl);
    var   myForm=$('postForm');   
    myForm.action=aurl;   
    myForm.method='POST';
    myForm.submit();   

}
function submitFavorite()
{
    var hidden=0;
    if($('hidden').checked)
    {
        hidden=1;
    }
    var tagName=$("tagName").value;
    if(tagName.trim()!='')
    {
        //alert(tagName);
        var tags=tagName.trim().split(" +");
        if(tags.length>3)
        {
            alert("每次最多输入3个标签");
            return;
        }else
        {
            for(var i=0;i<tags.length;i++)
            {
                if(tags[i].length>10)
                {
                    alert("标签名字过长，每个标签不能超过10个汉字或字符");
                    return;     
                }
                
            }
        }
        if(tagName.length>33)
        {
            alert("标签名字过长，每个标签不能超过10个汉字或字符");
            return;
        }
    }
    var url="/product/favorite.do";
    var paras="productID="+$('fproductID').value+"&tagName="+encodeURIComponent($("tagName").value)+"&hidden="+hidden;
    var myAjax = new Ajax.Request(url, {
        method:"post",
        parameters:paras,
        onComplete:showFavoriteResult
    });
}
function showFavoriteResult(oResponse)
{
    var res=oResponse.responseText;
    var s=res.split('=');
    var flag=s[0];
    var promt=s[1];
    if(flag!=1)
    {
        alert(promt);
    }
    closeFloatPage();
//window.location.reload();
}

function focusOtherFriend()
{
    var otherFriend=$('otherFriend').value;
    //alert(otherFriend);
    if(otherFriend=="您也可以自己输入您朋友的email,多个请用逗号隔开")
    {
        $('otherFriend').value="";
    }
}

function findFriend()
{
    var userName=$('user').value;
    var password=$('password').value;
    var siteName=$('siteName').value;
    
    if(siteName=='')
    {
        alert("请选择信箱类型");
        return;
    }

    if(userName==""||password=="")
    {
        alert("用户名或密码填写不完整");
        return;
    }
    
    userName=userName+"@"+siteName;
    var siteType=$('siteType').value;
    //$('friendList').width="10px";
    $('friendList').style.display="";
    $('friendList').innerHTML="<img src='/images/loading.gif' height='13' width='150'>";
    var url="/product/requestContactBook.do";
    var paras="userName="+userName+"&siteType="+siteType+"&password="+password;
    new Ajax.Request(url,{
        method:"post",
        parameters:paras,
        onComplete:updateFriendList
    });
}
function updateFriendList(oResponse)
{
    // $('friendList').style.height="100px";
    //  $('friendList').style.display="";
    $('friendList').innerHTML=oResponse.responseText;
}
function selectAllFriendEmail(com)
{
    var checked=com.checked;
    var friendemails=document.getElementsByName("friendEmail");
    for(var i=0;i<friendemails.length;i++)
    {
        friendemails[i].checked=checked;
    }
}

function selectAllFriendEmail1(com)
{
    var checked=com.checked;
    var friendemails=document.getElementsByName("friendEmail1");
    for(var i=0;i<friendemails.length;i++)
    {
        friendemails[i].checked=checked;
    }
}

function selectAllUserEmail(com)
{
    var checked=com.checked;
    var useremails=document.getElementsByName("userEmail");
    for(var i=0;i<useremails.length;i++)
    {
        useremails[i].checked=checked;
    }
}

function submitRecommend()
{
    var email="";
    var emailComs=null;
    var sender="";
    var contentID="content2";
    var senderID="sender2";
    //alert(currentTabPage);
    if(currentTabPage==1)
    {
        //recommend to user of yihaodian
        emailComs=document.getElementsByName("userEmail");
        contentID="content1";
    }else
    {
        if(currentTabPage==2)
        {
            emailComs=document.getElementsByName("friendEmail");
            contentID="content2";
            senderID="sender2";
        }else
        {
            contentID="content3";
            senderID="sender3";
            var otherFriend=$('otherFriend').value;
            //alert(otherFriend);
            if(otherFriend!='')
            {
                if(email!="")       
                {
                    email+=",";    
                }
                email+=otherFriend;
            }
        }
        var senderCom=document.getElementById(senderID);
        if(senderCom)
        {
            sender=senderCom.value;
        }
    }
    if(emailComs)
    {
        for(var i=0;i<emailComs.length;i++)
        {
            if(emailComs[i].checked)
            {
                if(email!="")
                {
                    email+=",";
                }
                email+=emailComs[i].value;
            }
        }
    }
    
   
    if(email=="")
    {
        alert("未选择用户好友,不能推荐");
        return;
    }
    if(!checkEmail(email))
    {
        alert("邮件格式不正确,邮件之间请用半角逗号隔开");
        return;
    }
    var link=$('link').value;
    // alert(link)
    var productID=$("rproductID").value;
    var productName=$("productName").value;
    var content=$(contentID).value;
    var url="/product/recommend.do";
    var paras="link="+link+"&productID="+productID+"&productName="+productName+"&email="+email+"&userWall.content="+encodeURIComponent(content)+"&sender="+encodeURIComponent(sender);
    //alert(paras);
    //window.location.href="/product/recommend.do?link="+link+"&productID="+productID+"&productName="+productName+"&email="+email+"&userWall.content="+encodeURIComponent(content);
    new Ajax.Request(url,{
        method:"post",
        parameters:paras,
        onComplete:showRecommendResult
    });
}

//每日两款
function submitDailySpecialRecommend(){
	 var email="";
	    var emailComs=null;
	    var sender="";
	    var contentID="content2";
	    var senderID="sender2";
	    var dailySpecialId=document.getElementById('dailySpecialId').value;
	    //alert(currentTabPage);
	    if(currentTabPage==1)
	    {
	        //recommend to user of yihaodian
	        emailComs=document.getElementsByName("userEmail");
	        contentID="content1";
	    }else
	    {
	        if(currentTabPage==2)
	        {
	            emailComs=document.getElementsByName("friendEmail");
	            contentID="content2";
	            senderID="sender2";
	        }else
	        {
	            contentID="content3";
	            senderID="sender3";
	            var otherFriend=$('otherFriend').value;
	            //alert(otherFriend);
	            if(otherFriend!='')
	            {
	                if(email!="")       
	                {
	                    email+=",";    
	                }
	                email+=otherFriend;
	            }
	        }
	        var senderCom=document.getElementById(senderID);
	        if(senderCom)
	        {
	            sender=senderCom.value;
	        }
	    }
	    if(emailComs)
	    {
	        for(var i=0;i<emailComs.length;i++)
	        {
	            if(emailComs[i].checked)
	            {
	                if(email!="")
	                {
	                    email+=",";
	                }
	                email+=emailComs[i].value;
	            }
	        }
	    }
	    
	   
	    if(email=="")
	    {
	        alert("未选择用户好友,不能推荐");
	        return;
	    }
	    if(!checkEmail(email))
	    {
	        alert("邮件格式不正确,邮件之间请用半角逗号隔开");
	        return;
	    }
	    var link=$('link').value;
	    // alert(link)
	    var productID=$("rproductID").value;
	    var productName=$("productName").value;
	    var content=$(contentID).value;
	    var url="/product/dailySpecialSendEmail.do";
	    var paras="link="+link+"&productID="+productID+"&productName="+productName+"&email="+email+"&userWall.content="+encodeURIComponent(content)+"&sender="+encodeURIComponent(sender)+"&dailySpecialId="+dailySpecialId;
	    //alert(paras);
	    //window.location.href="/product/recommend.do?link="+link+"&productID="+productID+"&productName="+productName+"&email="+email+"&userWall.content="+encodeURIComponent(content);
	    new Ajax.Request(url,{
	        method:"post",
	        parameters:paras,
	        onComplete:showRecommendResult
	    });
}

function showRecommendResult(oResponse)
{
    var res=oResponse.responseText;
    var s=res.split("=");
    var flag=s[0];
    var promt=s[1];
    if(flag!=1)
    {
        var descs=promt.split("|");
        var desc=descs[0];
        var message=descs[1];
        alert(desc+message);
    }else
    {
        closeFloatPage();
    }
}

function checkEmail(str)
{
    var result=true;
    var reg= /^[0-9a-zA-Z][_.0-9a-z-A-Z]{0,31}@([0-9a-zA-Z][0-9a-z-A-Z_]{0,30}[0-9a-zA-Z]\.){1,4}[a-zA-Z]{2,4}(,[0-9a-zA-Z][_.0-9a-z-A-Z]{0,31}@([0-9a-zA-Z][0-9a-z-A-Z_]{0,30}[0-9a-zA-Z]\.){1,4}[a-zA-Z]{2,4})*$/g;
    if(!str.match(reg))
    {
        result=false;
    }

    return result;
}


var currentTabPage=2;
function showTab(n)
{
    var i=0;
    currentTabPage=n;
    for(i=1;i<4;i++)
    {
        var tobj=document.getElementById('trecommend'+i);
        if(tobj != null)
        {
            if(i==n)
            { 
                tobj.style.display='';
            }
            else
            {
                tobj.style.display='none';
            }
        }
    }
    
    
    for(i=1;i<4;i++)
    {
        if(document.getElementById('celrecommend'+i) != null)
        {
            document.getElementById('celrecommend'+i).className="comment_bg2";
        }

    }
    document.getElementById('celrecommend'+n).className="comment_bg";
}



function showRecommandTab(n)
{
    var i=0;
    currentTabPage=n;
    for(i=1;i<4;i++)
    {
        var tobj=document.getElementById('trecommend'+i);
        if(tobj != null)
        {
            if(i==n)
            { 
                tobj.style.display='';
            }
            else
            {
                tobj.style.display='none';
            }
        }
    }
    
    
    for(i=1;i<4;i++)
    {
        if(document.getElementById('celrecommend'+i) != null)
        {
            document.getElementById('celrecommend'+i).className="comment_bg2";
        }

    }
    document.getElementById('celrecommend'+n).className="comment_bg";
}




function toBreakWord(intLen, id){
    var obj=document.getElementById(id);
    var strContent=obj.innerHTML;
    var strTemp="";
    while(strContent.length>intLen){
        strTemp+=strContent.substr(0,intLen)+"<br/>";
        strContent=strContent.substr(intLen,strContent.length);
    }
    strTemp+= strContent;
    obj.innerHTML=strTemp;
}
    
function getQueryStringRegExp(name)
{
    var reg = new RegExp("(^|\\?|&)"+ name +"=([^&]*)(\\s|&|$)", "i");
    if (reg.test(location.href)) return unescape(RegExp.$2.replace(/\+/g, " ")); return "";
}

function skipAddFriend()
{
    var allDiv=document.getElementById("allEmailDiv");
    if(allDiv)
    {
        allDiv.style.display='';
        var memberDiv=document.getElementById("memberDiv");
        if(memberDiv)
        {
            memberDiv.style.display="none";
        }
    }
}

function byteLength(sStr){
    var aMatch=sStr.match(/[^\x00-\x80]/g);
    return(sStr.length+(!aMatch?0:aMatch.length));
}  


function getIEVersion()
{
    if(navigator.appName == "Microsoft Internet Explorer")
    {
        if(navigator.appVersion.match(/7./i)=='7.')
        {
            return "7";
        }
        if(navigator.appVersion.match(/6./i)=='6.')
        {
            return "6";
        }
    }
    return "0";       
} 

function isIE6()
{
    return getIEVersion()=="6";
}

function findDimensions() //函数：获取尺寸

{
    //获取窗口宽度
    if (window.innerWidth)
    {
        winWidth = window.innerWidth;
    }
    else if ((document.body) && (document.body.clientWidth))
    {
        winWidth = document.body.clientWidth;
    }
    //获取窗口高度

    if (window.innerHeight)
    {
        winHeight = window.innerHeight;
    }
    else if ((document.body) && (document.body.clientHeight))
    {
        winHeight = document.body.clientHeight;
    }

    //通过深入Document内部对body进行检测，获取窗口大小

    if (document.documentElement  && document.documentElement.clientHeight &&

        document.documentElement.clientWidth)

        {
        winHeight = document.documentElement.clientHeight;
        winWidth = document.documentElement.clientWidth;
    }

    return ("width:"+winWidth,"height:"+winHeight); 

}


function submitPriceFeedback()
{
    var branchName=$('priceFeedback.marketBranchName').value;
    var marketDate=$('marketDate').value;
    var marketPrice=$('priceFeedback.marketPrice').value;
    if(branchName.trim()=='')
    {
        alert("请输入门店名字");
        return;
    }
    if(branchName.length>255)
    {
        alert("门店名字长度不能超过255字符");
        return;
    }
    if(marketPrice.length>8)
    {
        alert("价格不能超过99999999");
        return;
    }
    var priceReg=/^\d+(.\d{1,2})?$/g;
    if(!marketPrice.match(priceReg))
    {
        alert("价格不正确");
        return;
    }
    var dateReg=/^[2-9]\d{3}-((0?[1-9])|(1[0-2]))-((0?[1-9])|([1-2]\d)|(3[0-1]))$/g;
    if(marketDate==''||!marketDate.match(dateReg))
    {
        alert("日期格式不正确");
        return;
    }
    var currentDate=new Date();
    var dstr=marketDate.replace(/-/g, "/");
    var s=new Date(dstr);
    if(s.getTime()>currentDate.getTime())
    {
        alert("日期不正确");
        return;
    }
    
    
    var url="/feedback/priceFeedback.do";
    var paras=$('pfbfrm').serialize();
    new Ajax.Request(url,{
        parameters:paras,
        onComplete:completePriceFeedback
    });
}
function completePriceFeedback(oResponse)
{
    var response=oResponse.responseText;    
    if(response=='1')
    {
        closeFloatPage();
    }else
    {
        showFloatResult(response);
    }
}

//标签按钮动态切换
  var defaultNewProductTabId=0;
  var defaultFriendRecommendTabId=0;
  function changeHotProductTab(prefix,id)
  {
  for(var i=1;i<=3;i++)
  {
  document.getElementById(prefix+"_tab"+i).className=prefix+"_tab2";
  document.getElementById(prefix+"_content_"+i).style.display="none";
  }
  document.getElementById(prefix+"_tab"+id).className=prefix+"_tab1";
  document.getElementById(prefix+"_content_"+id).style.display="";
  }
  function changeNewProductTab(id)
  {
  if(id!=defaultNewProductTabId)
  {
  document.getElementById("newproduct_tab_"+id).className="default_tab1";
  document.getElementById("newproduct_content_"+id).style.display="";
  document.getElementById("newproduct_tab_"+defaultNewProductTabId).className="default_tab2";
  document.getElementById("newproduct_content_"+defaultNewProductTabId).style.display="none";
  defaultNewProductTabId=id;
  }
  }
  function changeFriendRecommend(id)
  {
  if(id!=defaultFriendRecommendTabId)
  {
  document.getElementById("friendlyRecommend_tab_"+id).className="default_tab1";
  document.getElementById("friendlyRecommend_content_"+id).style.display="";
  document.getElementById("friendlyRecommend_tab_"+defaultFriendRecommendTabId).className="default_tab2";
  document.getElementById("friendlyRecommend_content_"+defaultFriendRecommendTabId).style.display="none";
  defaultFriendRecommendTabId=id;
  }
  }
  function changeLeftTopTab(id)
  {
  var oldId=id==1?2:1;
  document.getElementById("leftTopTab"+id).className="tag_bg3";
  document.getElementById("leftTopTab"+oldId).className="tag_bg3_";
  document.getElementById("leftTopContent_"+id).style.display="";
  document.getElementById("leftTopContent_"+oldId).style.display="none";
  }
  
  var defaultRightTopTabId=1;
  function changeRightTop(id)
  {
  if(id!=defaultRightTopTabId)
  {
  document.getElementById("rightTopTab_"+id).className="tag_bg3";
  document.getElementById("rightTopTab_"+defaultRightTopTabId).className="tag_bg3_";
  document.getElementById("rightTopContent_"+id).style.display="";
  document.getElementById("rightTopContent_"+defaultRightTopTabId).style.display="none";
  defaultRightTopTabId=id;
  }
  }

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
             ImgD.height=(image.height*FitWidth)/image.width;}
             else
             {
             ImgD.width=image.width;
             ImgD.height=image.height;
             }
             }
            else{
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

//统计来路链接编码
function URLencode(sStr)
{
	return escape(sStr).replace(/\+/g, '%2B').replace(/\"/g,'%22').replace(/\'/g, '%27').replace(/\//g,'%2F');
}

function shiftLeft()
{
    currentPage++;
    if(currentPage>=pages)
    {
        currentPage=0;
    }
    var comName="pic_page_"+currentPage;
    document.getElementById('scrollDiv').scrollLeft=document.getElementById(comName).offsetLeft;
}
function shiftRight()
{
    currentPage--;
    if(currentPage<0)
    {
        currentPage=pages-1;
    }
    var comName="pic_page_"+currentPage;
    document.getElementById('scrollDiv').scrollLeft=document.getElementById(comName).offsetLeft;
}
function openImageWindow()
{
    window.open('productImage.do?product.id='+currentProductID,'','width=650,height=580,scrollbars=0');
}
function switchPic(picID)
{
    var pic_id="pic_"+picID;
    //var userCom="user_"+picID;
    //var userName=$(userCom).value;
    //var userID=$('userid_'+picID).value;
    var imgPath=document.getElementById(pic_id).value;
    var imgbig=document.getElementById('pic800800_'+picID).value;
    //var creatorType=$("user_creator_type_"+picID).value;
    var imgContent="<image src='"+encodeURI(imgPath)+"' border='0' onclick='openImageWindow();' style='cursor:pointer'/>";
    //alert("imagContent:"+imgContent);
    document.getElementById('mainpicimg').src=encodeURI(imgPath);
    document.getElementById('mainpicimg').alt=encodeURI(imgbig);
   // $('picDiv').innerHTML=imgContent;
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

// Set cookie value, currently not used
function addCookie(c_name,value,expiredays)
{
    var exdate=new Date();
    exdate.setDate(exdate.getDate()+expiredays);
    document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
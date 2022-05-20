<%if piaofu=1 then%>
<script>
var brOK=false;
var mie=false;
var vmin=2;
var vmax=5;
var vr=3;
var timer1;
var jobads;


function movechip(chipname)
{
	if(brOK)
	{
		eval("chip="+chipname);
		if(!mie)
		{
			pageX=window.pageXOffset;
			pageW=window.innerWidth;
			pageY=window.pageYOffset;
			pageH=window.innerHeight;
		} 
		else
		{
			pageX=window.document.body.scrollLeft;
			pageW=window.document.body.offsetWidth-8;
			pageY=window.document.body.scrollTop;
			pageH=window.document.body.offsetHeight;
		}
		chip.xx=chip.xx+chip.vx;
		chip.yy=chip.yy+chip.vy;
		chip.vx+=vr*(Math.random()-0.5);
		chip.vy+=vr*(Math.random()-0.5);
		if(chip.vx>(vmax+vmin))  chip.vx=(vmax+vmin)*2-chip.vx;
		if(chip.vx<(-vmax-vmin)) chip.vx=(-vmax-vmin)*2-chip.vx;
		if(chip.vy>(vmax+vmin))  chip.vy=(vmax+vmin)*2-chip.vy;
		if(chip.vy<(-vmax-vmin)) chip.vy=(-vmax-vmin)*2-chip.vy;
		if(chip.xx<=pageX)
		{
			chip.xx=pageX;
			chip.vx=vmin+vmax*Math.random();
		}
		if(chip.xx>=pageX+pageW-chip.w)
		{
			chip.xx=pageX+pageW-chip.w;
			chip.vx=-vmin-vmax*Math.random();
		}
		if(chip.xx>=680)
		{
			chip.xx=chip.xx-20;
			chip.vx=-vmin-vmax*Math.random();
		}
		if(chip.yy<=pageY)
		{
			chip.yy=pageY;
			chip.vy=vmin+vmax*Math.random();
		}
		if(chip.yy>=pageY+pageH-chip.h)
		{
			chip.yy=pageY+pageH-chip.h;
			chip.vy=-vmin-vmax*Math.random();
		}
		if(!mie)
		{
			eval('document.'+chip.named+'.top ='+chip.yy);
			eval('document.'+chip.named+'.left='+chip.xx);
		}
		else
		{
			eval('document.all.'+chip.named+'.style.pixelLeft='+chip.xx);
			eval('document.all.'+chip.named+'.style.pixelTop ='+chip.yy);
		}
		chip.timer1=setTimeout("movechip('"+chip.named+"')",80);
	}
}

function stopme(chipname)
{
	if(brOK)
	{
		eval("chip="+chipname);
		if(chip.timer1!=null)
		{
			clearTimeout(chip.timer1)
		}
	}
}

function jobads()
{
	if(navigator.appName.indexOf("Internet Explorer")!=-1)
	{
		if(parseInt(navigator.appVersion.substring(0,1))>=4) brOK=navigator.javaEnabled();mie=true;
	}
	if(navigator.appName.indexOf("Netscape")!=-1)
	{
		if(parseInt(navigator.appVersion.substring(0,1))>=4) brOK=navigator.javaEnabled();
	}
	jobads.named="jobads";
	jobads.vx=vmin+vmax*Math.random();
	jobads.vy=vmin+vmax*Math.random();
	jobads.w=1;
	jobads.h=1;
	jobads.xx=0;
	jobads.yy=0;
	jobads.timer1=null;
	movechip("jobads");
}

document.write('<div id="jobads" style="height:49px;left:178px;position:absolute;top:1237px;width:70px; z-index:1000">');
document.write('<a href="<%=piaofuurl%>" target="_blank" onmouseover=stopme("jobads"); onmouseout=movechip("jobads");>');
document.write('<img src="<%=piaofupic%>" border="0" alt="<%=piaofutit%>"></a></div>');
jobads();
</script>
<%end if%>

<% ' 弹出次数
if tanchu_time=1 then
if request.cookies("goldweb")("ad")=Request.ServerVariables("REMOTE_ADDR") then tanchu=0
Response.cookies("goldweb").path="/"
response.cookies("goldweb")("ad")=Request.ServerVariables("REMOTE_ADDR")
end if

if tanchu=1 then
%>
<script>
window.open ("<%=tanurl%>", "newwindow", "height=<%=tanheight%>, width=<%=tanwidth%>, top=<%=tantop%>, left=<%=tanleft%>, toolbar=NO, menubar=NO, scrollbars=NO, resizable=NO,location=NO, status=NO")
</script>
<%
end if

if cebian=1 then
%>
<DIV id=ad_dl01 style="Z-INDEX: 1; LEFT: 10px; VISIBILITY: visible; WIDTH: 100px; POSITION: absolute; TOP: 55px">
<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
  <TBODY>
  <TR>
    <TD align=left><A onclick="ad_dl01.style.visibility='hidden'"><IMG height=16 src="../images/small/adbg.gif" width=100 border=0></TD>
  </TR>
  <TR>
    <TD><a href="<%=lefturl%>" target="_blank"><IMG src="<%=leftpic%>" width=100 height=250 border=0></a></TD>
  </TR></TBODY></TABLE>
</DIV>

<DIV id=ad_dl02 style="Z-INDEX: 1; RIGHT: 10px; VISIBILITY: visible; WIDTH: 100px; POSITION: absolute; TOP: 55px">
<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
  <TBODY>
  <TR>
    <TD align=right><A onclick="ad_dl02.style.visibility='hidden'"><IMG height=16 src="../images/small/adbg.gif" width=100 border=0></A></TD>
  </TR>
  <TR>
    <TD><a href="<%=righturl%>" target="_blank"><IMG height=250 src="<%=rightpic%>" width=100 border=0></a></TD>
  </TR></TBODY></TABLE>
</DIV>
<%
end if
%>

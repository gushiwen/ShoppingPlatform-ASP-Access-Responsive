//移动端访问首页跳转到移动首页
function uaredirect(new_url)
{ 
	(function(Switch){
		var switch_pc = window.location.hash;
		if(switch_pc != "#pc"){
			if(/iphone|ipod|ipad|ipad|Android|nokia|blackberry|webos|webos|webmate|bada|lg|ucweb|skyfire|sony|ericsson|mot|samsung|sgh|lg|philips|panasonic|alcatel|lenovo|cldc|midp|wap|mobile/i.test(navigator.userAgent.toLowerCase())){
				Switch.location.href=new_url;
			}
		}
	})(window);
}
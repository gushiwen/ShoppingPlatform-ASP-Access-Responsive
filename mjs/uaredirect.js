//非移动端访问将跳转至PC页
function uaredirect(new_url)
{ 
	(function(Switch){
		var switch_mob = window.location.hash;
		if(switch_mob != "#mobile"){
			if(!/iphone|ipod|ipad|ipad|Android|nokia|blackberry|webos|webos|webmate|bada|lg|ucweb|skyfire|sony|ericsson|mot|samsung|sgh|lg|philips|panasonic|alcatel|lenovo|cldc|midp|wap|mobile/i.test(navigator.userAgent.toLowerCase())){
				Switch.location.href = new_url;
			}
		}
	})(window);
}
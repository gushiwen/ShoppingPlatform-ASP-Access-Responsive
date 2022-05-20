<%
Set rssetup=conn.execute("select * from shopsetup") 
kaiguan=rssetup("kaiguan")
guanbi=rssetup("guanbi")
mosi=rssetup("mosi") 'SSL or not

loginskin=rssetup("loginskin")
yzm_skin=rssetup("yzm_skin")

bodyfixed=rssetup("bodyfixed")

siteurl=rssetup("siteurl")
adm_mail= rssetup("adm_mail")

' Admin text info in both CHN and EN
sitename=rssetup("sitename")
ensitename=rssetup("ensitename")
sitekeywords=rssetup("sitekeywords")
ensitekeywords=rssetup("ensitekeywords")
sitedescription=rssetup("sitedescription")
ensitedescription=rssetup("ensitedescription")
adm_comp=rssetup("adm_comp")
enadm_comp=rssetup("enadm_comp")
adm_name=rssetup("adm_name")
enadm_name=rssetup("enadm_name")
adm_address=rssetup("adm_address")
enadm_address=rssetup("enadm_address")
adm_kf= rssetup("adm_kf")
enadm_kf= rssetup("enadm_kf")

adm_tel=rssetup("adm_tel")
adm_fax=rssetup("adm_fax")
adm_msn=rssetup("adm_msn")

jmail=rssetup("jmail")                  '发信组件
reg_mailyesorno=rssetup("reg_mailyesorno")	'是否发信
mailserver=rssetup("mailserver")		'smtp服务器地址
mailname=rssetup("mailname")			'发信邮箱
mailpassword=rssetup("mailpassword")		'发信邮箱密码

kf_color= rssetup("kf_color")
newsmove=rssetup("newsmove")
news_skin=rssetup("news_skin")

' UserType text info in both CHN and EN
usertype1=rssetup("usertype1")
enusertype1=rssetup("enusertype1")
usertype2=rssetup("usertype2")
enusertype2=rssetup("enusertype2")
usertype3=rssetup("usertype3")
enusertype3=rssetup("enusertype3")
usertype4=rssetup("usertype4")
enusertype4=rssetup("enusertype4")
usertype5=rssetup("usertype5")
enusertype5=rssetup("enusertype5")
usertype6=rssetup("usertype6")
enusertype6=rssetup("enusertype6")

kou1=rssetup("kou1")
kou2=rssetup("kou2")
kou3=rssetup("kou3")
kou4=rssetup("kou4")
kou5=rssetup("kou5")
kou6=rssetup("kou6")

fei1=rssetup("fei1")
fei2=rssetup("fei2")
fei3=rssetup("fei3")
fei4=rssetup("fei4")
fei5=rssetup("fei5")
fei6=rssetup("fei6")
pei1=rssetup("pei1")
pei2=rssetup("pei2")
pei3=rssetup("pei3")
pei4=rssetup("pei4")
pei5=rssetup("pei5")
pei6=rssetup("pei6")
mianyoufei=clng(rssetup("mianyoufei"))
mianyoufei_msg=rssetup("mianyoufei_msg")

enfei1=rssetup("enfei1")
enfei2=rssetup("enfei2")
enfei3=rssetup("enfei3")
enfei4=rssetup("enfei4")
enfei5=rssetup("enfei5")
enfei6=rssetup("enfei6")
enpei1=rssetup("enpei1")
enpei2=rssetup("enpei2")
enpei3=rssetup("enpei3")
enpei4=rssetup("enpei4")
enpei5=rssetup("enpei5")
enpei6=rssetup("enpei6")
enmianyoufei=clng(rssetup("enmianyoufei"))
enmianyoufei_msg=rssetup("enmianyoufei_msg")

pic_xiaogao=rssetup("pic_xiaogao")
lar_color=rssetup("lar_color")		'大类颜色
mid_color=rssetup("mid_color")		'中类颜色
tree_num=rssetup("tree_num")         '首页左侧资讯数量
tree_view=rssetup("tree_view")       '首页左侧资讯列表是否显示
tree_display=rssetup("tree_display")
index_tishi=rssetup("index_tishi")

' Product lated text info in both CHN and EN
globalpriceunit=rssetup("globalpriceunit")
englobalpriceunit=rssetup("englobalpriceunit")
quehuo=rssetup("quehuo")
enquehuo=rssetup("enquehuo")
wujiage=rssetup("wujiage")
enwujiage=rssetup("enwujiage")
huiyuanjia=rssetup("huiyuanjia")
enhuiyuanjia=rssetup("enhuiyuanjia")

prompt_num= rssetup("prompt_num")
newprod_num= rssetup("newprod_num")
renmen_num= rssetup("renmen_num")
fenlei_num= rssetup("fenlei_num")
per_page_num= rssetup("per_page_num")
reg= rssetup("reg")
reg_check= rssetup("reg_check")
bbs= rssetup("bbs")
help_hang=rssetup("help_hang")
menu= rssetup("menu")
lockip= rssetup("lockip")
prod_pic="pic/"
qqonline=rssetup("qqonline")
adm_qq= rssetup("adm_qq")
adm_qq_name= rssetup("adm_qq_name")
whereqq=rssetup("whereqq")
kefuskin=rssetup("kefuskin")
qqskin=rssetup("qqskin")
qqmsg_on=rssetup("qqmsg_on")
qqmsg_off=rssetup("qqmsg_off")
lockip=rssetup("lockip")
ip=rssetup("ip")
topmenu_view=rssetup("topmenu_view")
topmenu=rssetup("topmenu")
music=rssetup("other1")
set rssetup=nothing

If mosi=1 Then 
  FullSiteUrl="https://"&siteurl
Else
  FullSiteUrl="http://"&siteurl
End If 

TimeOffset=+15 '正或负值
AdvDaysSetting="30,90,180,360"

AreaCollectionEN = "Aljunied;;Ang Mo Kio;;Bishan-Toa Payoh;;Bukit Panjang;;Chua Chu Kang;;East Coast;;Holland-Bukit Timah;;Hong Kah North;;Hougang;;Joo Chiat;;Jurong;;Marine Parade;;Moulmein-Kallang;;Mountbatten;;Nee Soon;;Pasir Ris-Punggol;;Pioneer;;Potong Pasir;;Punggol East;;Radin Mas;;Sembawang;;Sengkang;;Tampines;;Tanjong Pagar;;Whampoa;;West Coast;;Yuhua"
AreaCollectionCH = "丰加北;;先驱;;武吉班让;;后港;;如切;;蒙巴登;;波东巴西;;榜鹅东;;拉丁马士;;盛港西;;黄埔和裕华;;荷兰-武吉知马;;摩绵-加冷;;阿裕尼;;碧山大巴窑;;蔡厝港;;东海岸;;裕廊;;马林百列;;义顺;;三巴旺;;淡滨尼;;丹绒巴葛;;西海岸;;宏茂桥;;白沙-榜鹅"
%>

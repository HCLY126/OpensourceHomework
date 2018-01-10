# # -*- coding:utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )
import urllib
import urllib2
import cookielib
from excel import *
from user import *

List=[]
cookie = cookielib.CookieJar()
opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie))
postdata = urllib.urlencode({'zjh':user(0),'mm':user(1)})
loginUrl = 'http://zhjw.dlut.edu.cn/loginAction.do'
result = opener.open(loginUrl,postdata)
gradeUrl = 'http://zhjw.dlut.edu.cn/xkAction.do?actionType=6'
result = opener.open(gradeUrl)
html = etree.HTML(result.read().decode('gbk'))
schedule = html.xpath('//td[@class="pageAlign"]/table[@border="1"]')
write_schedule(cut(get_son(schedule[0],List)))
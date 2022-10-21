from asyncio.windows_events import NULL
from lxml import etree
import re
from urllib import request, response
import requests 

parser = etree.HTMLParser(encoding='utf-8')

# 获取所有的推文时间
def WeuiDesktopMassTime(start, end):
  weuiDesktopMassTime = []
  for i in range(start, end):
    tree = etree.parse("E:/Code/py爬虫/作品表爬虫/"+str(i)+".html", parser=parser)
    html = etree.tostring(tree,encoding="utf-8").decode()
    result = tree.xpath("//*[@class='weui-desktop-mass__time']/text()")
    weuiDesktopMassTime.extend(result)
  #print(weuiDesktopMassTime)
  return weuiDesktopMassTime

# 获取所有的推文链接
def GetSiteList(start, end):
  siteLists = []
  for i in range(start, end):
    tree = etree.parse("E:/Code/py爬虫/作品表爬虫/"+str(i)+".html", parser=parser)
    html = etree.tostring(tree,encoding="utf-8").decode()
    result = tree.xpath("//*[@class='weui-desktop-mass-appmsg__title']/@href")
    siteLists.extend(result)
  return siteLists

# 获取所有推文中的责任姓名
def GetName(siteList):
  headers = {"User-Agent":""}
  for i in range(len(siteList)-1,0,-1):
    dict = {"time":NULL, "activity":NULL, "edit":NULL, "text":NULL, "proofreading":NULL, "picture":0, "video":NULL}
    #print(siteList[i])
    # 错误：请求外网
    response = requests.get(siteList[i], headers=headers).text
    html = etree.HTML(response)
    publish_time = html.xpath("//*[@id='meta_content']/em[1]/@text()")# 时间
    activity_name = html.xpath("//*[@id='activity-name']")# 标题
    dict["activity"] = activity_name[0].text
    print(dict["time"],dict["activity"])
    sectionp = html.xpath("//div[@id='js_content']//*/text()")# 署名
    # pat = r"编辑：(?)"
    #sectionp = re.findall(pat, sectionp)
    print(sectionp)

getSiteList = GetSiteList(1,2)
GetName(getSiteList)
# weuiDesktopMassTime = WeuiDesktopMassTime(1,5)
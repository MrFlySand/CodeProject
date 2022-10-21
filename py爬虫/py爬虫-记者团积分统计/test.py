# -*- coding: utf-8 -*-
from lxml import etree
import re
from urllib import request, response
import requests 

parser = etree.HTMLParser(encoding='utf-8')
headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.34"}
url = "https://mp.weixin.qq.com/s?__biz=MzA5MjU5MjcxMA==&mid=2657508233&idx=1&sn=6a7a1c4e325210e90da630484f936fc0&chksm=8bf834a0bc8fbdb641c80d9a340db3985debeab371f5ca999d8a6c7ca6513165a18b7c6bc475#rd"

response = requests.get(url, headers=headers).text
html = etree.HTML(response)
publish_time = html.xpath("//em[@id='publish_time']/text()")
print(publish_time)

rich_media_title  = html.xpath("//h1[@id='activity-name']")
sectionp = html.xpath("//div[@id='js_content']/section/section/p/text()")
print(rich_media_title[0],sectionp)
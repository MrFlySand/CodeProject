from asyncio.windows_events import NULL
from lxml import etree
import re
from urllib import request, response
import requests 
from lxml import etree
import xlsxwriter

# 获取所有的推文链接
def GetSiteList(start, end):
	siteLists = []
	for i in range(start, end):
		parser = etree.HTMLParser(encoding='utf-8')
		tree = etree.parse("C:/Users/MrFlySand/Desktop/testPy/"+str(i)+".html", parser=parser)
		html = etree.tostring(tree,encoding="utf-8").decode()
		result = tree.xpath("//*[@class='weui-desktop-mass-appmsg__title']/@href")
		siteLists.extend(result)
	return siteLists

# 获取所有推文中的责任姓名
def GetName(siteList):
	# print(siteList)
	headers = {"User-Agent":"Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Mobile Safari/537.36 Edg/105.0.1343.33"}
	# 创建表格
	workbook = xlsxwriter.Workbook('dome.xlsx')
	worksheet = workbook.add_worksheet()
	bold = workbook.add_format({'bold': True})
	worksheet.write('A1', '记者团官方微信公众号 每月作品表汇总', bold)
	worksheet.write('A2', '日期', bold)
	worksheet.write('B2', '选题名称', bold)
	worksheet.write('C2', '编辑', bold)
	worksheet.write('D2', '校对', bold)
	worksheet.write('E2', '文字', bold)
	worksheet.write('F2', '图片', bold)
	worksheet.write('G2', '视频', bold)
	worksheet.write('H2', '音频', bold)
	worksheet.write('I2', '学生主编', bold)
	parser = etree.HTMLParser(encoding='utf-8')
	line = 2

	for i in range(len(siteList)-1,0,-1):
		line = line + 1
		dict = {"time":NULL, "activity":NULL, "edit":NULL, "text":NULL, "proofreading":NULL, "picture":0, "video":NULL}    
		# print(siteList[i])
		# 错误：请求外网
		# response = requests.get(siteList[i], headers=headers).text
		# html = etree.HTML(response)
		# url = r"https://mp.weixin.qq.com/s?__biz=MzA5MjU5MjcxMA==&mid=2657506933&idx=1&sn=7b0c93e80da90ab9fc7eb7d85c98f32a&chksm=8bf8395cbc8fb04a3f4e41492cbc5f8d872d8394ec854f9127e4244d73dc4695d31819efe080#rd"
		print(siteList[i])
		reponse = request.urlopen(siteList[i]).read().decode()
		html = etree.HTML(reponse)
		try:
			date = r"2022-[0-9]+-[0-9]+"
			date = re.search(date, reponse).group()
			worksheet.write('A'+str(line), reponse)
		except Exception as result:
			pass

		activity_name = html.xpath("//h1")[0].text.strip() #标题
		worksheet.write('B'+str(line), activity_name)

		try:
			bianJi = r"主编：[/u4e00-/u9fa5]+"
			bianJi = re.search(bianJi, reponse).group()
			worksheet.write('C'+str(line),bianJi)
			print(bianJi)
		except Exception as result:
			print("13")
			pass

		try:
				jiaoDui = r"校对：([/u4e00-/u9fa5]+)"
				jiaoDui = re.search(jiaoDui, reponse).group()
				worksheet.write('D'+str(line),jiaoDui)
				print(jiaoDui)
		except Exception as result:
				pass

		try:
				bianJi = r"文字：([/u4e00-/u9fa5]+)"
				bianJi = re.search(bianJi, reponse).group()
				worksheet.write('E'+str(line),bianJi)
				print(bianJi)
		except Exception as result:
				pass

		try:
				bianJi = r"图片：([/u4e00-/u9fa5]+)"
				bianJi = re.search(bianJi, reponse).group()
				worksheet.write('F'+str(line),bianJi)
				print(bianJi)
		except Exception as result:
				pass

		try:
				bianJi = r"视频：([/u4e00-/u9fa5]+)"
				bianJi = re.search(bianJi, reponse).group()
				worksheet.write('G'+str(line),bianJi)
				print(bianJi)
		except Exception as result:
				pass

		try:
				bianJi = r"音频：([/u4e00-/u9fa5]+)"
				bianJi = re.search(bianJi, reponse).group()
				worksheet.write('H'+str(line),bianJi)
				print(bianJi)
		except Exception as result:
				pass

		try:
				zhuBian = r"学生主编：([/u4e00-/u9fa5]+)"
				zhuBian = re.search(zhuBian, reponse).group()
				worksheet.write('I'+str(line),zhuBian)
				print(zhuBian)
		except Exception as result:
				pass
				
	# 关闭工作薄
	workbook.close()

getSiteList = GetSiteList(1,2)
# print(GetSiteList)
GetName(getSiteList)

from asyncio.windows_events import NULL
from doctest import Example
import importlib
from pathlib import Path
from lxml import etree
import re
from urllib import request, response
import requests 
from urllib import request
import re #进行数据清洗要导入此模块
from asyncio import sleep
from datetime import date,datetime
# from openpyxl import load_workbook
import time
import pdfkit
import os, sys
import datetime
import pdfkit
import time
import webbrowser
import webbrowser as web
from pynput import mouse
import pyautogui
# import win32api
# import win32con
import driver
from selenium import webdriver
import pyperclip

cur_file_dir = os.path.abspath(__file__).rsplit("\\", 1)[0]

# 获取所有的推文链接
def GetSiteList(start, end,path):
  siteLists = []
  print("\n程序正在运行...")
  for i in range(start, end):
    parser = etree.HTMLParser(encoding='utf-8')
    try:
      tree = etree.parse(path+str(i)+"公众号.html", parser=parser)
    except Exception as result:
      print("错误：你保存的html文件名称错误，正确文件名称为：1.html、2.html、3.html，请重新运行程序。")
      pass 

    html = etree.tostring(tree,encoding="utf-8").decode()
    result = tree.xpath("//*[@class='weui-desktop-mass-appmsg__title']/@href")
    siteLists.extend(result)
  return siteLists

# 导出pdf
def PrintPdf(url):
  wk_path = r'D:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
  config = pdfkit.configuration(wkhtmltopdf=wk_path)
  # web.open(url) #默认浏览器打开url
  
  control = mouse.Controller()
  # 滑动网页到最下方
  driver=webdriver.Edge() #打开edge浏览器
  driver.maximize_window()
  driver.get(url)
  time.sleep(2)
  
  reponse = request.urlopen(url).read().decode()

  try:
    pat1 = r"var ct = \"(\d+)\""        
    date1 = re.search(pat1, reponse).group(1)
    date1 = int(date1)
    #转换为其他日期格式,如:"%Y-%m-%d %H:%M:%S"
    timeArray = time.localtime(date1)
    otherStyleTime = time.strftime("%m月%d日", timeArray)

  except Exception as result:
    try:
      pat2 = r"window.ct = \'(\d+)\'"
      date2 = re.search(pat2, reponse).group(1)
      date2 = int(date2)
      timeArray = time.localtime(date2)
      otherStyleTime = time.strftime("%m月%d日", timeArray)
    except Exception as result:
      pass
  
  # 标题
  try:
    html = etree.HTML(reponse)        
    activity_name = html.xpath("//h1")[0].text.strip() #标题
  except Exception as result:
    pass


  temp_height=0
  while True:
    #循环将滚动条下拉
    driver.execute_script("window.scrollBy(0,300)")
    #sleep一下让滚动条反应一下
    time.sleep(1)
    #获取当前滚动条距离顶部的距离
    check_height = driver.execute_script("return document.documentElement.scrollTop || window.pageYOffset || document.body.scrollTop;")
    #如果两者相等说明到底了
    if check_height==temp_height:
        break
    temp_height=check_height
    # print(check_height)

  time.sleep(1)
  pyautogui.keyDown('ctrl')    # 按下shift
  pyautogui.press('p')    # 按下 4
  pyautogui.keyUp('ctrl')   # 释放 shift

  time.sleep(10)
  pyautogui.keyDown('enter')
  pyautogui.keyUp('enter')

  # time.sleep(1)
  # pyautogui.keyDown('left')
  # pyautogui.keyUp('left')

  
  print(otherStyleTime,",",activity_name)
  pyperclip.copy(otherStyleTime)#将规定复制到系统剪贴板#
  pyautogui.keyDown('ctrl')    # 按下shift
  pyautogui.press('v')    # 按下 4
  pyautogui.keyUp('ctrl')   # 释放 shift

  pyautogui.press('-')

  pyperclip.copy(activity_name)#将规定复制到系统剪贴板#
  pyautogui.keyDown('ctrl')    # 按下shift
  pyautogui.press('v')    # 按下 4
  pyautogui.keyUp('ctrl')   # 释放 shift

  time.sleep(2)
  pyautogui.keyDown('enter')
  pyautogui.keyUp('enter')

  time.sleep(3)

if __name__ == '__main__':
  # 获取当前文件的路径
  if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
  elif __file__:
    application_path = os.path.dirname(__file__)
  path = application_path.replace("\\","/")+"/"

  print("关注微信公众号：小知识酷，获取更多程序内容\n")
  minNum = input("请输入html文件最小起始数：")
  maxNum = input("请输入html文件最大终止数：")
  ## 获取当前文件夹下的所有html文件中文章的url
  getSiteList = GetSiteList(int(minNum),int(maxNum)+1, path)
  # PrintPdf(getSiteList[0])
  # print(len(getSiteList))
  for i in range(0,len(getSiteList)):
   PrintPdf(getSiteList[i])
   if(i%10==0):
    os.system('taskkill /F /IM msedge.exe')

  print("\nExcel文件位置："+path)
  input("\n程序运行完毕")

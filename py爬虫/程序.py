from base64 import encode
import urllib.request
import re
from lxml import etree

# 读取文档中所有的数据
def readAll(url):
  str = open(url,encoding = "utf-8")
  strs = ""
  # 读取每一行
  for line in str.readlines():                             
    strs = line.strip()  + strs     
  return strs

# 正则表达式匹配字符
strs = readAll("C:/Users/MrFlySand/Desktop/1.txt")
pat = re.compile(r'[a-z]+' or '[A-Z][a-z]+')
data = pat.findall(strs)

# 统计每个单词出现的次数并加入到dict字典中
dict = {}
for i in range(0,len(data)):   
  if data[i] in dict:
    dictValue = dict[data[i]]+1
    dict.update({data[i]:dictValue})
  else:
    dict.update({data[i]:1})

# 根据字典的value值进行排序
dict = sorted(dict.items(),  key=lambda dict: dict[1], reverse=True)

for key in dict:
  print(key)


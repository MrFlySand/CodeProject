from urllib import request
import re #进行数据清洗要导入此模块

url = r"https://mp.weixin.qq.com/s?__biz=MzA5MjU5MjcxMA==&mid=2657508233&idx=1&sn=6a7a1c4e325210e90da630484f936fc0&chksm=8bf834a0bc8fbdb641c80d9a340db3985debeab371f5ca999d8a6c7ca6513165a18b7c6bc475#rd"

reponse = request.urlopen(url).read().decode()

# 通过正则表达式进行数据清洗
data = re.findall("<title>(.*?)</title>", reponse)
print((data))

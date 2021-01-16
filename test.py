import tkinter as tk
import tkinter.messagebox
from tkinter import *
import selenium as se
from selenium import webdriver
from selenium.webdriver.remote.command import Command
import urllib3
import time
import openpyxl
import requests
import bs4

"""driver = webdriver.Chrome()
driver.get('https://jobs.51job.com/guangzhou-hzq/128123507.html?s=01&t=0')
time.sleep(5)
#print(driver.current_url)
#print(type(driver.current_url))
msg = driver.find_element_by_class_name('tCompany_main').find_element_by_class_name('bmsg').text
print(msg.strip())
time.sleep(10)
driver.quit()"""

"""try:
 driver.execute(Command.STATUS)
except OSError:
 print("OsError!!!")
except ConnectionRefusedError:
 print("ConnectionRefusedError!!!")
except urllib3.exceptions.NewConnectionError:
 print("NewConnectionError!!!")
except urllib3.exceptions.MaxRetryError:
 print("MaxRetryError!!!")
except Exception as e:
 print(e)"""


"""
root = Tk()
s = Spinbox(root, values = (0,5,10,15,20))
s.pack()
print(type(int(s.get())))
tk.messagebox.showerror('错误', '无法识别该网页!')
if tk.messagebox.askyesno(title = '警告', message = '是否要停止爬取?'):
 print('是')
else:
 print('否')
root.mainloop()"""


"""while True:
 try:
  time.sleep(1)
 except KeyboardInterrupt:
  print('中断')
  break
 else:
  print('else')"""

"""driver = webdriver.Chrome()
driver.set_page_load_timeout(10)
try:
 driver.get('https://jobs.51job.com/guangzhou-hzq/128123507.html?s=01&t=0')
except se.common.exceptions.InvalidArgumentException:
 print('URL参数错误!')
except se.common.exceptions.WebDriverException:
 print('连接超时，请检查网络！')
else:
 print('正常')"""

#wb = openpyxl.open(r'C:\Users\Administrator\Desktop\ddd.xlsx')
#ws = wb.active
#time.sleep(5)
#ws.append([0,1,2,3,4,5,6,7,8,9])
#try:
# wb.save(r'C:\Users\Administrator\Desktop\ddd.xlsx')
#except PermissionError:
# print('Excel表格无权限写入，请关闭已经打开该表格的应用')
#wb.close()

"""headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36'}
while True:
 try:
  res = requests.get('https://jobs.51job.com/guangzhou-yxq/128155496.html?s=01&t=0', headers=headers ,timeout = 5)
 except requests.exceptions.ConnectTimeout:
  print('连接超时！  10秒后重连')
  time.sleep(10)
 except requests.exceptions.ConnectionError:
  print('无法连接该网页，请检查网络！  10秒后重连')
  time.sleep(10)
 else:
  break
print('连接成功')"""

def open_URL(url):                #使用requests模块获取前程无忧网站各工作的职位描述和要求
        tmp_time = 3           #最多重连3次
        bmsg = None
        while tmp_time > 0:
                headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36'}
                while True:
                        try:
                                res = requests.get(url, headers=headers ,timeout = 5)   #5秒后超时
                        except requests.exceptions.ConnectTimeout:
                                print('连接超时！ 5秒后重连')
                                time.sleep(5)
                        except (requests.exceptions.MissingSchema, requests.exceptions.InvalidURL):
                                print('URL参数错误！')
                                return ""
                        except requests.exceptions.ConnectionError:
                                print('连接出错！')
                                return ""
                        else:
                                break
                res.encoding = 'gbk'      #字符编码为gbk
                soup = bs4.BeautifulSoup(res.text,'html.parser')
                bmsg = soup.find('div', class_='bmsg')
                if bmsg == None:
                        print('获取页面信息失败！  5秒后重试')  #bs4找不到该元素
                        time.sleep(5)
                        tmp_time -= 1
                else:
                        break
        if tmp_time == 0:
                return ""
        else:
                return bmsg.text.replace('\n',' ').replace('\xa0',' ').replace('微信分享','')

print(open_URL('http://s'))

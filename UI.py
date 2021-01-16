import tkinter as tk
from tkinter import ttk   #使用Notebook模块
from tkinter import *
import tkinter.messagebox
import threading
from selenium import webdriver
from selenium.webdriver.remote.command import Command
import urllib3
import time
import os
import openpyxl
import requests     #网页请求模块
import bs4          #html解析模块


#全局变量
driver = None            #浏览器驱动变量
sleepTime = 5            #用于间隔休眠的时间，单位秒
startFlag = False        #是否未停止，用于控制爬虫是否停止。一旦按下开始键，是否暂停中或运行中，一律当爬虫未停止。
runFlag = False          #是否运行中，用于开始或暂停爬虫
stopFlag = False         #传递到爬虫内部，用于彻底停止爬虫
chromeOpen = False       #谷歌浏览器是否已打开
file_name = os.getcwd() + r'\data.xlsx'           #数据保存文件名
wb = None                #Excel工作本Workbook
ws = None                #Excel活动表格

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#爬虫实现


#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#界面自定义类
class UiRoot(Tk):
 def __init__(self):
  Tk.__init__(self)

class UiFrame(Frame):
 def run(self, func):
  self.master.run(func)
  
class BackgroundTask(object):
 """Similar to Android's AsyncTask"""
 def __init__(self, ui):
  self.ui = ui       #tk的root

 def doBefore(self):
  """Runs on the main thread, returns arg"""
  pass
 def do(self, arg):
  """Runs on a separate thread, returns result"""
  pass
 def doAfter(self, result):
  """Runs on the main thread again"""
  pass

 def run(self):
  """Invoke this on the main thread only"""
  arg = self.doBefore()
  threading.Thread(target=self._onThread, args=[arg]).start()     #新建一个线程处理去按钮触发的事件

 def _onThread(self, arg):
  result = self.do(arg)
  self.doAfter(result)
  #self.ui.run(lambda: self.doAfter(result))

class ShowTask(BackgroundTask):  #继承BackgroundTask类       打开谷歌浏览器触发事件
 def doBefore(self):
  self.ui.openGoogleButton.config(state=DISABLED)

 def do(self, arg):
  self.ui.log('打开谷歌浏览器')
  global driver
  global chromeOpen
  driver = webdriver.Chrome()
  #driver.get('https://www.baidu.com/')
  while True:
   try:
    driver.execute(Command.STATUS)   #检查浏览器状态
    time.sleep(1)
   except urllib3.exceptions.MaxRetryError:   #出现异常说明浏览器已退出
    self.ui.log('谷歌浏览器已关闭')
    tk.messagebox.showinfo("提示", "Google浏览器已关闭")
    break
   except Exception:
    self.ui.log('无法连接到浏览器，请关闭浏览器重试！')
    tk.messagebox.showwarning("警告", "无法连接到浏览器，请关闭浏览器重试！")
    break
   else:
    if chromeOpen == False:
     self.ui.log("浏览器打开成功")
     chromeOpen = True

 def doAfter(self, result):
  global chromeOpen
  self.ui.openGoogleButton.config(state=NORMAL)
  chromeOpen = False


class OpenBoss(BackgroundTask):            #打开Boss直聘按钮触发事件
 def doBefore(self):
  self.ui.openBossButton.config(state=DISABLED)

 def do(self, arg):
  if chromeOpen == False:                #判断浏览器是否打开
   self.ui.log("Google浏览器未打开")
   tk.messagebox.showwarning("警告", "Google浏览器未打开！")
   return
  self.ui.log("打开Boss直聘网页")
  driver.get('https://www.zhipin.com/')   #打开Boss直聘

 def doAfter(self, result):
  self.ui.openBossButton.config(state=NORMAL)


class OpenJob(BackgroundTask):         #打开前程无忧按钮触发事件
 def doBefore(self):
  self.ui.openJobButton.config(state=DISABLED)

 def do(self, arg):
  if chromeOpen == False:                 #判断浏览器是否打开
   self.ui.log("Google浏览器未打开")
   tk.messagebox.showwarning("警告", "Google浏览器未打开！")
   return
  self.ui.log("打开前程无忧网页")
  driver.get('https://www.51job.com/')   #打开前程无忧

 def doAfter(self, result):
  self.ui.openJobButton.config(state=NORMAL)


class Start_Pause(BackgroundTask):      #开始/暂停按钮触发事件
 def doBefore(self):
  pass

 def do(self, arg):
  global runFlag
  global startFlag
  global driver

  if chromeOpen == False:                           #判断浏览器是否打开
   self.ui.log("Google浏览器未打开")
   tk.messagebox.showwarning("警告", "Google浏览器未打开！")
   return

  url = driver.current_url
  if url.find('zhipin') == -1 and url.find('51job') == -1:     #判断网页是否打开正确
   tk.messagebox.showerror('错误', '无法识别该网页！')
   self.ui.log("无法识别网页！ 请打开Boss直聘或前程无忧网页")
   return

  self.ui.openBossButton.config(state=DISABLED)      #爬取期间禁止再次打开或更换网页
  self.ui.openJobButton.config(state=DISABLED)
  if runFlag == False:     #检查是否在运行
   self.ui.log("开始爬取")
   global sleepTime
   sleepTime = int(self.ui.spinbox.get())  #设置休眠时间
   self.ui.startButton.config(text = '暂停爬取')
   self.ui.spinbox.config(state=DISABLED)
   runFlag = True                 #运行中标志
   startFlag = True                #没停止标志
   while runFlag and not stopFlag:
    self.ui.log('工作中!')
    time.sleep(sleepTime)
  else:
   self.ui.log("暂停爬取")
   self.ui.startButton.config(text = '开始爬取')
   self.ui.spinbox.config(state=NORMAL)
   runFlag = False            #暂停中标志

 def doAfter(self, result):
  pass

class Stop(BackgroundTask):          #停止按钮触发事件
 def doBefore(self):
  self.ui.stopButton.config(state=DISABLED)

 def do(self, arg):
  global runFlag
  global stopFlag
  global startFlag
  if startFlag == False:                 #爬虫未开始却按下停止键
   self.ui.log("爬虫未处于运行或暂停中")
   tk.messagebox.showinfo("提示", "爬虫未处于运行或暂停中！")
   return
  if startFlag == True and tk.messagebox.askyesno(title = '提示', message = '是否要停止爬取？'):          #没停止并且确定要停止
   self.ui.log("停止爬取")
   self.ui.startButton.config(text = '开始爬取')
   self.ui.spinbox.config(state=NORMAL)
   runFlag = False         #没有运行标志
   stopFlag = True         #停止运行中的爬虫标志
   startFlag = False       #已停止标志
   time.sleep(5)
   stopFlag = False
   self.ui.log("爬虫已被停止")
   self.ui.openBossButton.config(state=NORMAL)      #在非工作爬取期间可以打开或更换网页
   self.ui.openJobButton.config(state=NORMAL)

 def doAfter(self, result):
  self.ui.stopButton.config(state=NORMAL)

class initWB(BackgroundTask):       #初始化工作本和工作表，无则创建
 global wb
 global ws
 def do(self, arg):
  time.sleep(2)
  try:
   wb = openpyxl.open(file_name)      #打开Excel文件
  except FileNotFoundError:
   self.ui.log('找不到Excel文件data!')
   self.ui.log('创建一个新的Excel文件')
   wb = openpyxl.Workbook()       #创建一个新的工作本
   ws = wb.active
   ws.append(['职位名称', '最低薪酬(元/月)', '最高薪酬', '平均薪酬', '公司所在地', '经验要求', '学历要求', '公司福利', '公司名称', '链接地址', '公司类型', '公司大小', '业务定位方向', '职位要求和描述']) #首行标题
   ws.freeze_panes = 'A2'  #冻结首行
   wb.save(file_name)      #保存文件
   self.ui.log('Excel文件data创建完成!')
  else:
   self.ui.log('成功找到并打开Excel文件data')
   ws = wb.active


class Check(UiFrame):     #查重界面
 def __init__(self, parent, **kwargs):
  UiFrame.__init__(self, parent, **kwargs)
  #mFrame = Labelframe(self)

class Generate(UiFrame):     #生成图片界面
 def __init__(self, parent, **kwargs):
  UiFrame.__init__(self, parent, **kwargs)
  #mFrame = Labelframe(self)

class Search(UiFrame):     #推荐界面
 def __init__(self, parent, **kwargs):
  UiFrame.__init__(self, parent, **kwargs)
  #mFrame = Labelframe(self)


#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#运行界面
class MainUi(UiRoot):
 """主窗口"""   
 def __init__(self):
  UiRoot.__init__(self)
  self.geometry('400x650')   #窗口大小
  self.maxsize(width = 400, height = 650)      #窗口大小不可变
  self.minsize(width = 400, height = 650)      #窗口大小不可变
  self.title('Selenium控制型爬虫')         #窗口标题
  img1 = PhotoImage(file = 'google.gif')
  img2 = PhotoImage(file = 'boss.gif')
  img3 = PhotoImage(file = '51job.gif')
  
  tabs = ttk.Notebook(self, width = 388, height = 230)
  crawlerFrame = Frame(tabs)
  self.openGoogleButton = Button(crawlerFrame, text = '打开浏览器', image = img1, width = 60, height = 60, command = ShowTask(self).run)  #按钮 
  self.openGoogleButton.grid(row = 0, column = 0, padx = 40, pady = 10)
  self.openBossButton = Button(crawlerFrame, text = '打开Boss直聘', image = img2, width = 60, height = 60, command = OpenBoss(self).run)
  self.openBossButton.grid(row = 0, column = 1, padx = 0, pady = 10)
  self.openJobButton = Button(crawlerFrame, text = '打开前程无忧', image = img3, width = 60, height = 60, command = OpenJob(self).run)
  self.openJobButton.grid(row = 0, column = 2, padx = 40, pady = 10)

  Label(crawlerFrame, text="打开浏览器", font = '12').grid(row = 1, column = 0, padx = 0, pady = 0)    #按钮提示文字 
  Label(crawlerFrame, text="打开Boss直聘", font = '12').grid(row = 1, column = 1, padx = 0, pady = 0)
  Label(crawlerFrame, text="打开前程无忧", font = '12').grid(row = 1, column = 2, padx = 0, pady = 0)

  Label(crawlerFrame, text="爬取每一页面间隔休眠时间(单位:秒)", font = '3').place(x = 30, y = 130)
  self.spinbox = Spinbox(crawlerFrame, values = (5,10,15,20,25,30,35,40,45,50,60,70,80,90,100), state = 'readonly', font = '3', width = 4)
  self.spinbox.place(x = 305, y = 132)

  self.startButton = Button(crawlerFrame, text = '开始爬取', width = 16, height = 2,   font = '13', command = Start_Pause(self).run)  #按钮  
  self.startButton.place(x = 40, y= 174)
  self.stopButton = Button(crawlerFrame, text = '结束爬取', width = 16, height = 2,   font = '13', command = Stop(self).run)
  self.stopButton.place(x = 220, y= 174)
  tabs.add(crawlerFrame, text = '爬虫')
  #--------------------------------------------------------------以上为爬虫界面--------------------------------------------------------------------------------------------------
  tabs.add(Check(self), text = '查重')
  tabs.add(Generate(self), text = '数据可视化')
  tabs.add(Search(self), text = '职位推送')
  tabs.place(x = 5, y = 5)

  #-----------------------------------------------------------------------------------------------------------------------------------------------------------------
  self.text = Text(self, width = 55, height = 29)  #Debug显示框
  self.text.place(x = 5, y = 265)
  self.text.configure(state=DISABLED)

  initWB(self).run()
  self.mainloop()

 def log(self, msg):
  print(msg)
  self.text.configure(state=NORMAL)
  self.text.insert(END, msg + '\n')
  self.text.configure(state=DISABLED)
  self.text.see(END)



if __name__ == '__main__':
 MainUi()

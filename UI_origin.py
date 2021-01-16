from tkinter import *

class show():
 def __init__(self,ui):
  self.ui = ui
  self.i = 0

 def print_click(self):
  print("click!")
  self.i += 1
  self.ui.log("CLICK" + str(self.i) + "\n")

def show0():
 print("click")
def show1():
 print("click")
def show2():
 print("click")
def show3():
 print("click")
def show4():
 print("click")


class UiRoot(Tk):
 def __init__(self):
  Tk.__init__(self)

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
class MainUi(UiRoot):
 """主窗口"""   
 def __init__(self):
  UiRoot.__init__(self)
  self.geometry("400x650")   #窗口大小
  self.title('Selenium控制型爬虫')         #窗口标题
  #mLabel = Label(self,text="窗口程序")
  #mLabel.pack()            #自动调节组件本身尺寸
  google_img = PhotoImage(file = 'google.gif')   #加载图片
  boss_img = PhotoImage(file = 'boss.gif')
  job_img = PhotoImage(file = '51job.gif')

  self.openGoogleButton = Button(self, image = google_img, text = '打开浏览器', width = 60, height = 54, command = show(self).print_click)  #按钮 
  self.openGoogleButton.grid(row = 0, column = 0, padx = 40, pady = 12)
  self.openBossButton = Button(self, image = boss_img, text = '打开Boss直聘', width = 60, height = 54, command = show0)
  self.openBossButton.grid(row = 0, column = 1, padx = 0, pady = 12)
  self.openJobButton = Button(self, image = job_img, text = '打开前程无忧', width = 60, height = 54, command = show0)
  self.openJobButton.grid(row = 0, column = 2, padx = 40, pady = 12)

  Label(self, text="打开浏览器", font = '12').grid(row = 1, column = 0, padx = 0, pady = 0)    #按钮提示文字
  Label(self, text="打开Boss直聘", font = '12').grid(row = 1, column = 1, padx = 0, pady = 0)
  Label(self, text="打开前程无忧", font = '12').grid(row = 1, column = 2, padx = 0, pady = 0)

  self.startButton = Button(self, text = '开始爬取', width = 16, height = 2, font = '13', command = show0)  #按钮
  self.startButton.place(x = 40, y= 125)
  self.stopButton = Button(self, text = '结束爬取', width = 16, height = 2, font = '13', command = show0)
  self.stopButton.place(x = 220, y= 125)

  #text_box = Text(app, width = 55, height = 34)  #Debug显示框
  #text_box.place(x = 5, y = 194)
  #self.scrollbar = Scrollbar(self,width = 2)
  #self.scrollbar.place(x = 398, y = 194)
  self.text = Text(self, width = 55, height = 34)
  self.text.place(x = 5, y = 194)
  self.text.configure(state=DISABLED)
  #self.scrollbar.config(command=self.text.yview)
  #self.redirectStreams()
  self.mainloop()

 def log(self, msg):
  self.text.configure(state=NORMAL)
  self.text.insert(END, msg)
  self.text.configure(state=DISABLED)
  self.text.see(END)




def main():
 MainUi()

if __name__ == '__main__':
 main()

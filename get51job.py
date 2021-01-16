from selenium import webdriver
import selenium as se
import time
import openpyxl
import requests
import bs4

file_name = r'C:\Users\Administrator\Desktop\广州游戏.xlsx'

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
        
def append_list_job(ee, ee_l, ee_d, ee_fl, ee_c, ee_link, ee_t_s, ee_b, ee_msg):  #职位名称  薪资  地区|经验|学历  福利  公司名称  链接  公司类型|大小  业务方向
        l = []
        sal_low = 0
        sal_hight = 0
        sal_avg = 0
        l.append(ee.replace(' ',''))
        ee_l = ee_l.replace(' ','')
        if ee_l.find('-') == -1:
                if ee_l.find('元/日') != -1:
                        sal_low = int(ee_l.split('元/日')[0].strip()) * 22              #统一标准单位元/月
                        sal_hight = int(ee_l.split('元/日')[0].strip()) * 26
                        sal_avg = (sal_low + sal_hight) // 2
                        l.append(sal_low)
                        l.append(sal_hight)
                        l.append(sal_avg)
                elif ee_l.find('万/月') != -1:
                        sal_low = int(float(ee_l.split('万/月')[0].strip()) * 10000)             #统一标准单位元/月
                        sal_hight = sal_low
                        sal_avg = sal_low
                        l.append(sal_low)
                        l.append(sal_hight)
                        l.append(sal_avg)
                elif ee_l.find('千/月') != -1:
                        sal_low = int(float(ee_l.split('千/月')[0].strip()) * 1000)             #统一标准单位元/月
                        sal_hight = sal_low
                        sal_avg = sal_low
                        l.append(sal_low)
                        l.append(sal_hight)
                        l.append(sal_avg)
                elif ee_l.find('万/年') != -1:
                        sal_low = int(float(ee_l.split('万/年')[0].strip()) * 10000 // 12)             #统一标准单位元/月
                        sal_hight = sal_low
                        sal_avg = sal_low
                        l.append(sal_low)
                        l.append(sal_hight)
                        l.append(sal_avg)
                else:
                        l.append(ee_l.strip())
                        l.append("null")
                        l.append("null")
        else:
                if ee_l.find('万/月') != -1:
                        try:
                                sal_low = int(float(ee_l.split('万/月')[0].split('-')[0]) * 10000)      #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('万/月')[0].split('-')[1]) * 10000)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                elif ee_l.find('千/月') != -1:
                        try:
                                sal_low = int(float(ee_l.split('千/月')[0].split('-')[0]) * 1000)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('千/月')[0].split('-')[1]) * 1000)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                elif ee_l.find('万/年') != -1:
                        try:
                                sal_low = int(float(ee_l.split('万/年')[0].split('-')[0]) * 10000 // 12)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('万/年')[0].split('-')[1]) * 10000 // 12)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                elif ee_l.find('元/日') != -1:
                        try:
                                sal_low = int(float(ee_l.split('元/日')[0].split('-')[0]) * 26)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('元/日')[0].split('-')[1]) * 26)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                else:
                        l.append(ee_l.strip())
                        l.append("null")
                        l.append("null")
        
        if ee_d.find('|') == -1:
                l.append(ee_d.strip())
                l.append("")
                l.append("")
        else:
                ee_d = ee_d.replace(' ','')       #去空格
                tmp_index = ee_d.rfind('招')      #从右往左找
                if tmp_index != -1:
                        ee_d = ee_d[0:tmp_index-1]        #去掉招收人数
                if len(ee_d.split('|')) == 3:
                        l.append(ee_d.split('|')[0].strip())       #公司所在地
                        l.append(ee_d.split('|')[1].strip())       #经验
                        l.append(ee_d.split('|')[2].strip())       #学历
                elif len(ee_d.split('|')) == 2:
                        tmp_index = 0
                        while tmp_index < len(ee_d):
                                if ee_d[tmp_index].isdigit():      #判断字符串中是否含有数字
                                        break
                                else:
                                        tmp_index += 1
                        if tmp_index < len(ee_d):               #含有数字
                                l.append(ee_d.split('|')[0].strip())   #公司所在地
                                l.append(ee_d.split('|')[1].strip())   #经验
                                l.append("")
                        else:
                                if ee_d.find('在校生') == -1 and ee_d.find('应届生') == -1 and ee_d.find('无需经验') == -1:   #不包含经验要求
                                        l.append(ee_d.split('|')[0].strip())   #公司所在地
                                        l.append("")
                                        l.append(ee_d.split('|')[1].strip())   #学历
                                else:                                                              #包含经验要求
                                        l.append(ee_d.split('|')[0].strip())   #公司所在地
                                        l.append(ee_d.split('|')[1].strip())   #经验
                                        l.append("")
                elif len(ee_d.split('|')) == 1:
                        l.append(ee_d.split('|')[0].strip())
                        l.append("")
                        l.append("")
        
        l.append(ee_fl.strip())               #福利
        l.append(ee_c.strip())                #公司名称
        l.append(ee_link.strip())             #链接地址
        if ee_t_s.find('|') == -1:
                if ee_t_s.find('-') == -1:
                        l.append(ee_t_s.strip())
                        l.append('null')
                else:
                        l.append('null')
                        l.append(ee_t_s.strip())
        else:
                l.append(ee_t_s.split('|')[0].strip())    #公司类型
                l.append(ee_t_s.split('|')[1].strip())    #公司大小
        l.append(ee_b.strip())                            #公司定位业务方向
        l.append(ee_msg.strip())                          #职位描述和要求
        print(ee)
        return l

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
keyWord = "引擎 客户端 开发 端游 手游 小游戏 Unity U3D u3d U3d UE4 ue4 Ue4 虚幻 cocos Cocos COCOS 2d 2D 3d 3D 测试 策划 美工 美术 设计 UI ui Ui 特效 动画 动作 服务器 维护 维稳 脚本 数据 前端 游戏"               
stopWord = "销售 推广 地产 经理 主管 客服 人事 教育 讲师 分销 电商 投放 运营 翻译 英语 商务 老师 译员 发行 主播 直播 玩家 试玩 帮派 治疗 市场 营销 游戏店 服务员 体验"
sw = stopWord.split(' ')
kw = keyWord.split(' ')
try:
 wb = openpyxl.open(file_name)      #打开文件
except FileNotFoundError:
 print('找不到Excel文件!')
 print('创建一个新的Excel文件')
 wb = openpyxl.Workbook()
 ws = wb.active
 ws.append(['职位名称', '最低薪酬(元/月)', '最高薪酬', '平均薪酬', '公司所在地', '经验要求', '学历要求', '公司福利', '公司名称', '链接地址', '公司类型', '公司大小', '业务定位方向', '职位要求和描述'])
 ws.freeze_panes = 'A2'
 wb.save(file_name)
 print('Excel文件创建完成!')
else:
 print('成功找到并打开Excel文件')
 ws = wb.active
driver = webdriver.Chrome()          #打开Google浏览器
driver.set_page_load_timeout(10)      #设置超时限制时间
while True:
 try:
  driver.get('https://search.51job.com/list/030200,000000,0000,00,9,99,%25E6%25B8%25B8%25E6%2588%258F,2,1.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=0&dibiaoid=0&line=&welfare=')
 except se.common.exceptions.WebDriverException:
  print('连接超时，请检查网络!')
  print('休眠中')
  time.sleep(10)
 else:
  break
print('休眠中，请设置好网页参数')
time.sleep(20)   #留足时间设置网页的搜索参数
url_this = driver.current_url
#tmp_list = []
while True:
 try:
  job_list = driver.find_element_by_class_name('j_joblist').find_elements_by_class_name('e')   #获取列表
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException 休眠中\n")
  time.sleep(15)
 else:
  if len(job_list) > 0:                 #获取列表成功
   break
  else:                                 #列表长度有误
   print("错误：工作列表长度为0!  休眠中")
   time.sleep(10)

for i in range(len(job_list)):
 e = job_list[i].find_element_by_class_name("jname")      #获取工作名称
 tmp = 0
 while tmp < len(sw):
  if e.text.find(sw[tmp]) != -1:             #包含停用词，跳到下一工作
   break
  else:
   tmp += 1
 if tmp < len(sw):
  continue
 tmp = 0
 while tmp < len(kw):
  if e.text.find(kw[tmp]) != -1:             #都不包含关键词，跳到下一工作
   break
  else:
   tmp += 1
 if tmp == len(kw):
  continue
 e_l = job_list[i].find_element_by_class_name("sal")     #工作工资
 e_d = job_list[i].find_element_by_class_name("d")      #工作要求
 e_c = job_list[i].find_element_by_class_name("cname")   #公司名称
 e_fl = ""
 try:
  e_fl = job_list[i].find_element_by_class_name("tags").get_attribute('title')   #福利
 except se.common.exceptions.NoSuchElementException:
  e_fl = ""
 except:
  print("\n获取福利出现其它错误\n")
 e_link = job_list[i].find_element_by_class_name("el").get_attribute('href')    #详情链接地址
 e_type_size = job_list[i].find_element_by_class_name("dc")                     #公司类型和大小
 e_business = job_list[i].find_element_by_class_name("int")                     #业务方向
 e_msg = open_URL(e_link)                                                       #职位描述和要求
 ws.append(append_list_job(e.text, e_l.text, e_d.text, e_fl, e_c.text, e_link, e_type_size.text, e_business.text, e_msg))
 
"""for each_list in tmp_list:
 while True:
  try:
   driver.get(each_list[9])                    #分别打开每个网页
   tmp_msg = driver.find_element_by_class_name('tCompany_main').find_element_by_class_name('bmsg').text.strip()      #获取并添加职位描述和要求
  except se.common.exceptions.WebDriverException:
   print('连接超时，请检查网络!')
   print('休眠10秒后重新连接')
   time.sleep(10)
  except se.common.exceptions.InvalidArgumentException:
   print('URL参数出错，跳过该链接!')                         #URL非法异常
   each_list.append('null')
   break
  except se.common.exceptions.NoSuchElementException:
   print('无法获取职位信息!  休眠中!')    #无法获取元素异常
   time.sleep(10)
  else:
   each_list.append(tmp_msg.replace('\n',' ').replace('微信分享','').replace('【',' ').replace('】',' '))       #去掉一些无关字符
   break
 ws.append(each_list)
 time.sleep(5)"""
 
while True:
 try:
  wb.save(file_name)       #保存文件
 except PermissionError:
  print('表格被其它程序占用中，无法写入数据!')
  print('10秒后重试')
  time.sleep(10)
 else:
  break
#tmp_list.clear()
print('保存成功,休眠中!!!')
"""while True:
 try:
  driver.get(url_this)                             #回到工作列表主页
 except se.common.exceptions.WebDriverException:
  print('连接超时，请检查网络!')
  print('休眠10秒后重新连接')
  time.sleep(10)
 else:
  break"""
time.sleep(10)
while True:
 try:
  driver.find_element_by_class_name("j_page").find_element_by_class_name("next").click()    #下一页
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException 休眠中\n")
  time.sleep(15)
 else:
  break
print('下一页,休眠中!!!')
time.sleep(10)

while driver.current_url != url_this:
 print("循环---------------------------------------------------")
 url_this = driver.current_url
 while True:
  try:
   job_list = driver.find_element_by_class_name('j_joblist').find_elements_by_class_name('e')   #获取列表
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException 休眠中\n")
   time.sleep(15)
  else:
   if len(job_list) > 0:                 #获取列表成功
    break
   else:                                 #列表长度有误
    print("错误：工作列表长度为0!  休眠中")
    time.sleep(10)
    
 #tmp_list.clear()
 for i in range(len(job_list)):
  e = job_list[i].find_element_by_class_name("jname")      #获取工作名称
  tmp = 0
  while tmp < len(sw):
   if e.text.find(sw[tmp]) != -1:             #包含停用词，跳到下一工作
    break
   else:
    tmp += 1
  if tmp < len(sw):
   continue
  tmp = 0
  while tmp < len(kw):
   if e.text.find(kw[tmp]) != -1:            #都不包含关键词，跳到下一工作
    break
   else:
    tmp += 1
  if tmp == len(kw):
   continue
  e_l = job_list[i].find_element_by_class_name("sal")     #工作工资
  e_d = job_list[i].find_element_by_class_name("d")      #工作要求
  e_c = job_list[i].find_element_by_class_name("cname")   #公司名称
  e_fl = ""
  try:
   e_fl = job_list[i].find_element_by_class_name("tags").get_attribute('title')   #福利
  except se.common.exceptions.NoSuchElementException:
   e_fl = ""
  except:
   print("\n其它错误\n")
  e_link = job_list[i].find_element_by_class_name("el").get_attribute('href')    #详情链接地址
  e_type_size = job_list[i].find_element_by_class_name("dc")                     #公司类型和大小
  e_business = job_list[i].find_element_by_class_name("int")                     #业务方向
  e_msg = open_URL(e_link)                                                       #职位描述和要求
  ws.append(append_list_job(e.text, e_l.text, e_d.text, e_fl, e_c.text, e_link, e_type_size.text, e_business.text, e_msg))

 """for each_list in tmp_list:
  while True:
   try:
    driver.get(each_list[9])                    #分别打开每个网页
    tmp_msg = driver.find_element_by_class_name('tCompany_main').find_element_by_class_name('bmsg').text.strip()      #获取并添加职位信息
   except se.common.exceptions.WebDriverException:
    print('连接超时，请检查网络!')                            #连接超时
    print('休眠10秒后重新连接')
    time.sleep(10)
   except se.common.exceptions.InvalidArgumentException:
    print('URL参数出错，跳过该链接!')                         #URL非法异常
    each_list.append('null')
    break
   except se.common.exceptions.NoSuchElementException:
    print('无法获取职位信息!  休眠中!')    #无法获取元素异常
    time.sleep(10)
   else:
    each_list.append(tmp_msg.replace('\n',' ').replace('微信分享','').replace('【',' ').replace('】',' '))        #去掉一些无关字符
    break
  ws.append(each_list)
  time.sleep(5)"""
 
 while True:
  try:
   wb.save(file_name)       #保存文件
  except PermissionError:
   print('表格被其它程序占用中，无法写入数据!')
   print('10秒后重试')
   time.sleep(10)
  else:
   break
 print('保存成功,休眠中!!!')
 """while True:
  try:
   driver.get(url_this)                             #回到工作列表主页
  except se.common.exceptions.WebDriverException:
   print('连接超时，请检查网络!')
   print('休眠10秒后重新连接')
   time.sleep(10)
  else:
   break"""
 time.sleep(10)
 while True:
  try:
   driver.find_element_by_class_name("j_page").find_element_by_class_name("next").click()    #下一页
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException 休眠中\n")
   time.sleep(15)
  else:
   break
 print('休眠中')
 time.sleep(10)
print("结束-------------------------------------------------------")
driver.close()

#影视 视频 后期 拍摄 短视频 剪辑 渲染 修图 摄像

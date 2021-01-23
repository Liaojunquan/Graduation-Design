from selenium import webdriver
import selenium as se
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl

file_name = r'C:\Users\Administrator\Desktop\data.xlsx'

def append_list_boss(ee_n, ee_l, ee_a, ee_d, ee_fl, ee_c, ee_link, ee_t_s, ee_b, ee_msg):  #职位名称  薪资  地区  经验|学历  福利  公司名称  链接  公司类型|大小  业务方向  职位描述和要求
        l = []
        sal_low = 0
        sal_hight = 0
        sal_avg = 0
        l.append(ee_n.strip())
        if ee_l.find('-') == -1:                   #薪酬字符串中不含-
                if ee_l.find('元/天') != -1:
                        sal_low = int(ee_l.split('元/天')[0].strip()) * 22              #统一标准单位元/月
                        sal_hight = int(ee_l.split('元/天')[0].strip()) * 26
                        sal_avg = (sal_low + sal_hight) // 2
                        l.append(sal_low)
                        l.append(sal_hight)
                        l.append(sal_avg)
                elif ee_l.find('K') != -1:
                        sal_low = int(float(ee_l.split('K')[0].strip()) * 1000)             #统一标准单位元/月
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
                elif ee_l.find('元/时') != -1:
                        sal_low = int(float(ee_l.split('元/时')[0].strip()) * 10 * 24)             #统一标准单位元/月
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
                elif ee_l.find('元/时') != -1:
                        try:
                                sal_low = int(float(ee_l.split('元/时')[0].split('-')[0]) * 10 * 24)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('元/时')[0].split('-')[1]) * 10 * 24)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                elif ee_l.find('元/天') != -1:
                        try:
                                sal_low = int(float(ee_l.split('元/天')[0].split('-')[0]) * 24)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('元/天')[0].split('-')[1]) * 24)
                                sal_avg = (sal_low + sal_hight) // 2
                                l.append(sal_low)
                                l.append(sal_hight)
                                l.append(sal_avg)
                        except:
                                l.append(ee_l.strip())
                                l.append("null")
                                l.append("null")
                elif ee_l.find('K') != -1:
                        try:
                                sal_low = int(float(ee_l.split('K')[0].split('-')[0]) * 1000)            #统一标准单位元/月
                                sal_hight = int(float(ee_l.split('K')[0].split('-')[1]) * 1000)
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
        l.append(ee_a.strip().replace(' ',''))             #公司位置
        if ee_d.find('学历不限') != -1:
                l.append(ee_d.split('学历不限')[0])        #经验要求
                l.append("学历不限")                       #学历要求
        elif ee_d.find('高中') != -1:
                l.append(ee_d.split('高中')[0])
                l.append("高中")
        elif ee_d.find('大专') != -1:
                l.append(ee_d.split('大专')[0])
                l.append("大专")
        elif ee_d.find('本科') != -1:
                l.append(ee_d.split('本科')[0])
                l.append("本科")
        elif ee_d.find('硕士') != -1:
                l.append(ee_d.split('硕士')[0])
                l.append("硕士")
        elif ee_d.find('博士') != -1:
                l.append(ee_d.split('博士')[0])
                l.append("博士")
        else:
                l.append(ee_d.strip())
                l.append("null")
        
        l.append(ee_fl.strip())               #福利
        l.append(ee_c.strip())                #公司名称
        l.append("https://www.zhipin.com" + ee_link.strip())             #链接地址
        tmp_s = ee_t_s.split(ee_b.strip())[1]
        if tmp_s.find('已上市') != -1:          #判断是否已上市
                l.append("已上市")              #公司类型
        else:
                l.append("未上市")
        index = 0
        while index < len(tmp_s):
                if tmp_s[index].isdigit():      #判断数字字符最初出现的位置
                        break
                else:
                        index += 1
        if index < len(tmp_s):
                l.append(tmp_s[index:len(tmp_s)])   #通过第一个数字字符位置分割字符串   公司大小
        else:
                l.append("")
        l.append(ee_b.strip())                            #公司定位业务方向
        l.append(ee_msg.replace('\n',' ').replace('【',' ').replace('】',' '))
        print(ee_n)
        return l

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
keyWord = """引擎 客户端 开发 端游 手游 小游戏 Unity U3D u3d U3d UE4 ue4 Ue4 UE 虚幻 cocos Cocos COCOS 2d 平面设计 三维 建模 渲染 视觉 MAYA Maya maya C4D c4d C4d 3DMAX 3dmax cad CAD ZBrush 
2D 3d 3D 测试 策划 美工 美术 设计 UI ui Ui 特效 动画 动作 服务器 维护 维稳 脚本 数据 前端 WEB web Web 游戏 贴图 材质 影视 视频 拍摄 摄影 摄像 导演 编导 前期 后期 剪辑 分镜 AR VR 研发经理 全栈 
小程序 数据库 数据管理 数据库管理 数据分析 DBA sql SQL Sql JAVA Java java C++ c++ C c Windows windows Android android 安卓 Python python 爬虫 图像识别 音频 音视频 VR  Vr vr 虚拟现实 AR ar 
Ar 虚拟仿真 Unreal 交互开发 程序开发 计算机图形 Flash flash FLASH An AN MG mg Mg 二维 分镜 AE Ae ae PR pr Pr 达芬奇 多媒体 Davinci 调色 修图 调光 图片 数字媒体"""

stopWord = """销售 推广 地产 经理 主管 客服 人事 教育 讲师 分销 电商 投放 运营 翻译 英语 商务 玩游戏 陪 质检 猎头 游戏机 装配 美主 gm gs GM GS Gm Gs 生产 * 激光 扫描 机械 公关 理财 
教师 教研 老师 译员 发行 主播 直播 玩家 试玩 帮派 治疗 市场 营销 游戏店 服务员 体验 助教 经纪人 管培 投资 顾问 合伙人 投标 审核 制片人 演员 自动驾驶 售前 售后 院长 企业文化 体验馆 店长 
接待 财务 店员 营业员 内容生态 看房 看楼 讲解 ARM arm 收账 会计 半导体 工艺 交通 能源 客户 射频 SAE 标注 漆 涂料 化工 甲油胶 建筑 建材 油墨 印刷 环保 化学 热转印 汽车 胶 粉体 灯光 CAE FAE 
硬件 嵌入式 自动化 芯片 AE工程师 供应链 AE应用工程师 PR工程师 仪器 AE技术 驱动 电路 免费"""

tagStopWord = "漆 涂料 化工 甲油胶 建筑 建材 油墨 印刷 环保 化学 汽车 胶 粉体 灯光 调研 营销 市场 客户 乙方 射频 销售 公关 电商 创意 推广 广告 运营 项目 执行 策划 创意"
sw = stopWord.split(' ')
kw = keyWord.split(' ')
tsw = tagStopWord.split(' ')
wb = openpyxl.open(file_name)       #打开文件
ws = wb.active
driver = webdriver.Chrome()
driver.set_page_load_timeout(10)      #设置超时限制时间
while True:
 try:
  driver.get('https://www.zhipin.com/')
 except se.common.exceptions.WebDriverException:
  print('连接超时，请检查网络!')
  print('休眠中')
  time.sleep(10)
 else:
  break
print('休眠中，请设置好网页参数')
time.sleep(20)   #留足时间设置网页的搜索参数
url_this = driver.current_url
while True:
 try:
  job_list = driver.find_element_by_class_name('job-list').find_elements_by_class_name('job-primary')   #获取列表
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException 休眠15秒后重试\n")
  time.sleep(15)
 else:
  if len(job_list) > 0:                 #获取列表成功
   break
  else:                                 #列表长度有误
   print("错误：工作列表长度为0!  休眠中")
   time.sleep(10)

for i in range(len(job_list)):
 e_n = job_list[i].find_element_by_class_name("job-name").text      #获取工作名称
 e_a = job_list[i].find_element_by_class_name("job-area").text      #获取公司地区
 tmp = 0
 while tmp < len(sw):
  if e_n.find(sw[tmp]) != -1:    #包含停用词，跳到下一工作
   break
  else:
   tmp += 1
 if tmp  < len(sw):
  continue
 e_tag = ""
 try:
  e_tag = job_list[i].find_element_by_class_name("info-append").find_element_by_class_name("tags").text   #职位标签
  #print(e_tag)
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException\n")
 except:
  print("获取标签出现其它错误!")
 else:
  if e_tag != "":
   tmp = 0
   while tmp < len(tsw):
    if e_tag.find(tsw[tmp]) != -1:             #包含标签停用词，跳到下一工作
     break
    else:
     tmp += 1
   if tmp < len(tsw):
    continue
 tmp = 0
 while tmp < len(kw):
  if e_n.find(kw[tmp]) != -1:             #都不包含关键词，跳到下一工作
   break
  else:
   tmp += 1
 if tmp == len(kw):
  continue
 e_l = job_list[i].find_element_by_class_name("job-limit").find_element_by_class_name("red").text        #薪酬
 e_d = ""
 try:
  e_d = job_list[i].find_element_by_class_name("job-limit").find_element_by_tag_name('p').text            #经验与学历
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException\n")
  e_d = ""
 except:
  print("\n获取经验与学历出现其它错误\n")
  e_d = ""
 e_c = job_list[i].find_element_by_class_name("company-text").find_element_by_class_name("name").text    #公司名称
 e_fl = ""
 try:
  e_fl = job_list[i].find_element_by_class_name("info-append").find_element_by_class_name("info-desc").text     #福利
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException\n")
  e_fl = ""
 except:
  print("\n获取福利出现其它错误\n")
  e_fl = ""
 e_link = job_list[i].find_element_by_class_name("primary-box").get_attribute('href')                                              #详情链接地址
 e_type_size = ""
 try:
  e_type_size = job_list[i].find_element_by_class_name("company-text").find_element_by_tag_name("p").text                           #公司类型和大小
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException\n")
  e_type_size = ""
 except:
  print("\n获取类型和大小出现其它错误\n")
  e_type_size = ""
 e_business = ""
 try:
  e_business = job_list[i].find_element_by_class_name("company-text").find_element_by_tag_name("p").find_element_by_tag_name("a").text    #公司业务定位
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException\n")
  e_business = ""
 except:
  print("\n获取业务定位出现其它错误\n")
  e_business = ""
 tmp_index = 5         #最多5次尝试机会
 e_msg = ""
 while tmp_index > 0:
  primary_box = job_list[i].find_element_by_class_name("primary-box")                      #获取primary-box元素
  ActionChains(driver).move_to_element(primary_box).perform()                              #移动鼠标到primary-box元素
  time.sleep(1)
  try:
   e_msg = job_list[i].find_element_by_class_name("info-detail").find_element_by_class_name("detail-bottom-text").text      #获取职位描述和要求
  except se.common.exceptions.NoSuchElementException:
   print('无法获取职位描述和要求!  重试中')
   time.sleep(2)
   tmp_index -= 1
  else:
   break
 ws.append(append_list_boss(e_n, e_l, e_a, e_d, e_fl, e_c, e_link, e_type_size, e_business, e_msg))
 time.sleep(3)
 
wb.save(file_name)    #保存文件
print('保存成功,休眠中!!!')
time.sleep(10)
while True:
 try:
  driver.find_element_by_class_name("page").find_element_by_class_name("next").click()    #下一页
 except se.common.exceptions.NoSuchElementException:
  print("\nNoSuchElementException 休眠15秒后重试\n")
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
   job_list = driver.find_element_by_class_name('job-list').find_elements_by_class_name('job-primary')   #获取列表
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
  e_n = job_list[i].find_element_by_class_name("job-name").text      #获取工作名称
  e_a = job_list[i].find_element_by_class_name("job-area").text      #获取公司地区
  tmp = 0
  while tmp < len(sw):
   if e_n.find(sw[tmp]) != -1:    #包含停用词，跳到下一工作
    break
   else:
    tmp += 1
  if tmp  < len(sw):
   continue
  e_tag = ""
  try:
   e_tag = job_list[i].find_element_by_class_name("info-append").find_element_by_class_name("tags").text   #职位标签
   print(e_tag)
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException\n")
  except:
   print("获取标签出现其它错误!")
  else:
   if e_tag != "":
    tmp = 0
    while tmp < len(tsw):
     if e_tag.find(tsw[tmp]) != -1:             #包含标签停用词，跳到下一工作
      break
     else:
      tmp += 1
    if tmp < len(tsw):
     continue
  tmp = 0
  while tmp < len(kw):
   if e_n.find(kw[tmp]) != -1:             #都不包含关键词，跳到下一工作
    break
   else:
    tmp += 1
  if tmp == len(kw):
   continue
  e_l = job_list[i].find_element_by_class_name("job-limit").find_element_by_class_name("red").text        #薪酬
  e_d = ""
  try:
   e_d = job_list[i].find_element_by_class_name("job-limit").find_element_by_tag_name('p').text            #经验与学历
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException\n")
   e_d = ""
  except:
   print("\n获取经验与学历出现其它错误\n")
   e_d = ""
  e_c = job_list[i].find_element_by_class_name("company-text").find_element_by_class_name("name").text   #公司名称
  e_fl = ""
  try:
   e_fl = job_list[i].find_element_by_class_name("info-append").find_element_by_class_name("info-desc").text     #福利
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException\n")
   e_fl = ""
  except:
   print("\n获取福利出现其它错误\n")
   e_fl = ""
  e_link = job_list[i].find_element_by_class_name("primary-box").get_attribute('href')                                              #详情链接地址
  e_type_size = ""
  try:
   e_type_size = job_list[i].find_element_by_class_name("company-text").find_element_by_tag_name("p").text                           #公司类型和大小
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException\n")
   e_type_size = ""
  except:
   print("\n获取类型和大小出现其它错误\n")
   e_type_size = ""
  e_business = ""
  try:
   e_business = job_list[i].find_element_by_class_name("company-text").find_element_by_tag_name("p").find_element_by_tag_name("a").text    #公司业务定位
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException\n")
   e_business = ""
  except:
   print("\n获取业务定位出现其它错误\n")
   e_business = ""
  tmp_index = 5           #最多5次尝试机会
  e_msg = ""
  while tmp_index > 0:
   primary_box = job_list[i].find_element_by_class_name("primary-box")                      #获取primary-box元素
   ActionChains(driver).move_to_element(primary_box).perform()                              #移动鼠标到primary-box元素
   time.sleep(1)
   try:
    e_msg = job_list[i].find_element_by_class_name("info-detail").find_element_by_class_name("detail-bottom-text").text      #获取职位描述和要求
   except se.common.exceptions.NoSuchElementException:
    print('无法获取职位描述和要求!  重试中')
    time.sleep(2)
    tmp_index -= 1
   else:
    break
  ws.append(append_list_boss(e_n, e_l, e_a, e_d, e_fl, e_c, e_link, e_type_size, e_business, e_msg))
  time.sleep(3)
 wb.save(file_name)   #保存文件
 print('保存成功,休眠中!!!')
 time.sleep(10)
 while True:
  try:
   driver.find_element_by_class_name("page").find_element_by_class_name("next").click()    #下一页
  except se.common.exceptions.NoSuchElementException:
   print("\nNoSuchElementException 休眠15秒后重试\n")
   time.sleep(15)
  else:
   break
 print('休眠中')
 time.sleep(10)
print("结束-------------------------------------------------------")
wb.close()
driver.close()
driver.quit()

#影视 视频 后期 拍摄 短视频 剪辑 渲染 修图 摄像

from selenium import webdriver
import time

driver = webdriver.Chrome()
driver.get('https://www.zhipin.com/c101280100/?query=%E5%BD%B1%E8%A7%86&page=1&ka=page-1')
url_this = driver.current_url
ee = driver.find_elements_by_class_name("job-name")      #获取工作名称
ee_l = driver.find_elements_by_class_name("job-limit")   #工作限制要求
for i in range(len(ee)):
	print(ee[i].text + "  " + ee_l[i].text.replace("\n","  "))   #打印数据

time.sleep(5)
driver.find_element_by_class_name("page").find_element_by_class_name("next").click()    #下一页
while driver.current_url != url_this:
    url_this = driver.current_url
    try:
        ee = driver.find_elements_by_class_name("job-name")      #获取工作名称
        ee_l = driver.find_elements_by_class_name("job-limit")   #工作限制要求
    except NoSuchElementException:
        print("NoSuchElementException")
        time.sleep(15)
        continue
    if len(ee) == 0:
        time.sleep(10)
        print("长度为0")
        continue
    else:
        for i in range(len(ee)):
            print(ee[i].text + "  " + ee_l[i].text.replace("\n","  "))   #打印数据
    try:
        driver.find_element_by_class_name("page").find_element_by_class_name("next").click()    #下一页
    except NoSuchElementException:
        print("NoSuchElementException")
        time.sleep(15)
        continue
    time.sleep(10)
print("结束")
driver.close()

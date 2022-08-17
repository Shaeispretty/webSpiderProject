from pydoc import stripid
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import xlwt
from selenium import webdriver


options = webdriver.ChromeOptions()

# 处理SSL证书错误问题
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')

# 忽略无用的日志
options.add_experimental_option(
    "excludeSwitches", ['enable-automation', 'enable-logging'])

search = input('搜索： ').strip()
wd = webdriver.Chrome(options=options)
wd.get('https://search.jd.com/Search?keyword='+search)
wd.maximize_window()
js = 'var q=document.documentElement.scrollTop=400'
wd.execute_script(js)
sleep(2)
# 输入框输入商品名称

names = []
nums = []
CPUs = []
for j in range(1,4):
    for i in range(1, 4):
        n = i
        div = wd.find_element(
            By.XPATH, '//*[@id="J_goodsList"]/ul/li[{key}]/div/div[1]/a/img'.format(key=n))

        ActionChains(wd).move_to_element(
            div).move_by_offset(5, 5).click().perform()
        sleep(2)

        windows = wd.window_handles
        wd.switch_to.window(windows[-1])
        name = wd.find_element(
            By.XPATH, '//*[@id="detail"]/div[2]/div[1]/div[1]/ul[3]/li[1]').get_attribute('textContent')
        num = wd.find_element(
            By.XPATH, '//*[@id="detail"]/div[2]/div[1]/div[1]/ul[3]/li[2]').get_attribute('textContent')
        CPU = wd.find_element(
            By.XPATH, '//*[@id="detail"]/div[2]/div[1]/div[1]/ul[3]/li[5]').get_attribute('textContent')
        names.append(name)
        nums.append(num)
        CPUs.append(CPU)
        wd.close()
        windows = wd.window_handles
    wd.switch_to.window(windows[0])
    wd.find_element(By.XPATH, '//*[@id="J_bottomPage"]/span[1]/a[9]')
    sleep(2)
    print("第" + str(j) + "页完成")
col = ('names', 'nums', 'CPUs')

workbook = xlwt.Workbook(encoding='utf-8')
# 创建一个workbook 设置编码
worksheet = workbook.add_sheet("手机规格", cell_overwrite_ok=True)

for i in range(0, 4):
    worksheet.write(0, i, col[i])
for i in range(0, 300):
    worksheet.write(i+1, 0, names[i])
for i in range(0, 300):
    worksheet.write(i+1, 1, nums[i])
for i in range(0, 300):
    worksheet.write(i+1, 2, CPUs[i])

workbook.save('手机规格.xls')
print("WorkDone")
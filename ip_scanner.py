from selenium import webdriver
import time
import ctypes, sys

from selenium.common.exceptions import ElementNotVisibleException


def is_admin():
  try:
    return ctypes.windll.shell32.IsUserAnAdmin()
  except:
    return False
# 前台开启浏览器模式
def openChrome():
    # 加启动配置
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')
    option.add_argument('headless')
    # 打开chrome浏览器
    driver = webdriver.Chrome(chrome_options=option)
    return driver

# 授权操作
def operationAuth(driver):
    url = "http://192.168.1.64/"
    driver.get(url)

    try:
        elem_pw = driver.find_elements_by_xpath("//input[@type='password']")
        for elem in elem_pw:
            elem.send_keys("Buchou123")
        # class="aui_state_highlight" class="aui_close"  placeholder="用户名"
        driver.find_element_by_xpath("//button[@class='aui_state_highlight']").click()
        time.sleep(1)
        driver.find_element_by_xpath("//div[@class='aui_titleBar']//a").click()
    except ElementNotVisibleException:
        print("已注册")
        driver.find_element_by_xpath("//input[@type='password']").clear()
    try:
        driver.find_element_by_xpath("//input[@type='password']").send_keys("Buchou123")
        driver.find_element_by_xpath("//input[@placeholder='用户名']").send_keys("admin")
        driver.find_element_by_xpath("//button").click()
    except:
        pass
    print("登陆成功")
    time.sleep(2)
    driver.find_element_by_xpath("//a[@ng-bind='oLan.config']").click()
    time.sleep(1)
    driver.find_element_by_xpath("//span[@class='menu-icon network-icon']").click()
def change_ip(driver,ip):

  time.sleep(1)

  driver.find_element_by_xpath('//input[@ng-model="oEthernetParams.szIpv4"]').clear()
  driver.find_element_by_xpath('//input[@ng-model="oEthernetParams.szIpv4"]').send_keys(ip)
  time.sleep(1)
  try:
    driver.find_element_by_xpath('//i[@class ="success"]')
  except:
    print("请输入正确格式的ip，如:192.168.1.65(注意点号是英文标点)")
    ip=input()
    change_ip(driver, ip)
    return 0
  driver.find_element_by_xpath('//input[@ng-click="testIP()"]').click()
  time.sleep(10)
  print(driver.find_element_by_xpath('//td[@class="aui_main"]//div').get_attribute('textContent'))
  if(driver.find_element_by_xpath('//td[@class="aui_main"]//div').get_attribute('textContent')=="该地址尚未被使用"):
    driver.find_element_by_xpath('//button[@class="aui_state_highlight"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//span[@ng-bind="oLan.save"]').click()
    print("修改完成，关闭程序即可")
    c=input()
  else:
    print("重新输入ip，如:192.168.1.65(注意点号是英文标点)")
    ip=input()
    change_ip(driver,ip)
    return 0
# 该地址尚未被使用
# 方法主入口
if __name__ == '__main__':

    if is_admin():
        print("已使用管理员权限打开")
        # 加启动配置
        driver = openChrome()
        operationAuth(driver)
        print("请输入ip后按回车键，如:192.168.1.65(注意点号是英文标点)")
        ip=input()
        change_ip(driver,ip)

    else:
        print("未使用管理员权限打开，自动获取管理员权限")
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)

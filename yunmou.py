
# encoding=utf8
from selenium import webdriver
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
import time
import openpyxl

def get_index(column_name,sheet):
    index = 0
    for name in sheet.rows.__next__():

        if (name.value==column_name):
            return index
        index += 1
def load_data():
    book = openpyxl.load_workbook('养户摄像头对应表.xlsx')
    sheet = book.worksheets[0]
    get_index("序列号",sheet)
    devices = []
    first_row_flag =0
    for rows in sheet.rows:
        if(first_row_flag==0):
            first_row_flag=1
        else:
            single_device = []
            if(rows[get_index("序列号",sheet)].value==None):
                break
            single_device.append(rows[get_index("公司", sheet)].value)
            single_device.append(rows[get_index("地区", sheet)].value)
            single_device.append(rows[get_index("服务部", sheet)].value)
            single_device.append(rows[get_index("序列号",sheet)].value)
            single_device.append(rows[get_index("设备名称",sheet)].value)
            # print(single_device)
            devices.append(single_device)

    return devices
def yunmou():

    # 加启动配置
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')

    # 打开chrome浏览器
    driver = webdriver.Chrome(chrome_options=option)
    url ="https://www.hik-cloud.com/chain/login/index.html#/login"
    driver.get(url)
    time.sleep(1)
    driver.find_element_by_xpath('//input[@placeholder="请输入用户名"]').send_keys("13792886247")
    driver.find_element_by_xpath('//input[@placeholder="请输入密码"]').send_keys("Buchou@0306")
    driver.find_element_by_xpath('//div[@class="login-btn-div"]').click()
    time.sleep(3)
    try:
        driver.find_element_by_xpath('//button[@class="el-button el-button--default el-button--primary "]').click()
    except Exception :
        pass
    devices=load_data()
    time.sleep(3)
    url = 'https://www.hik-cloud.com/chainX/index.html#/system/equipment/management'
    driver.get(url)
    for device_info in devices:

        print(device_info)
        time.sleep(1)
        title1=device_info[0]
        time.sleep(1)
        print(title1)
        driver.find_elements_by_xpath('//span[@title="' + title1 + '"]/../span')[0].click()
        time.sleep(1)
        title2 =device_info[1]
        time.sleep(1)
        driver.find_elements_by_xpath('//span[@title="' + title2 + '"]/../span')[0].click()
        time.sleep(1)
        title3=device_info[2]
        if(title3==None):
            driver.find_elements_by_xpath('//span[@title="' + title2 + '"]')[0].click()
        else:
            driver.find_element_by_xpath()
        time.sleep(1)
        driver.find_element_by_xpath('//i[@class="h-icon-add"]').click()
        time.sleep(1)
        inputs = driver.find_elements_by_xpath('//div[@class="el-form-item__content"]//div[@class="el-input"]//input[@autocomplete="off"]')
        device_num=device_info[3]
        device_verify_code= "12341234q"
        device_name = device_info[4]
        inputs[0].send_keys(device_num)
        inputs[1].send_keys(device_verify_code)
        inputs[2].send_keys(device_name+device_num)
        driver.find_elements_by_xpath('//div[@class="el-dialog__wrapper"]//div[@class="el-dialog__footer"]//div[@class="dialog-footer"]//button[@class="el-button el-button--primary"]//span')[3].click()
        try:
            driver.find_elements_by_xpath('//i[@class="el-dialog__close el-icon h-icon-close"]')[3].click()
        except:
            pass

        time.sleep(1)
        # try:
        # driver.find_element_by_xpath('//div[text()="%s"]//..//..//..//td[@class="el-table_1_column_8   is-hidden"]//div//span'%device_num).click()
        # print(len(driver.find_elements_by_xpath('//div[text()="%s"]//..//..//..//td[@class="el-table_1_column_8   is-hidden"]//div//span[@class="el-tooltip operator open"]'%device_num)))
        try:
            driver.find_elements_by_xpath('//div[text()="%s"]//..//..//..//td//div//span[@class="el-tooltip operator open"]'%device_num)[2].click()
            driver.find_element_by_css_selector("#HKClosePasswordDialog > form > div.el-form-item.is-required-right > div > div > input").send_keys("12341234q")
            driver.find_element_by_xpath('//span[text()="关闭加密"]').click()
            driver.find_element_by_xpath('//span[text()="确定"]').click()
        except:
            print('关闭加密出现问题，设备已经加密过或者设备离线')
        driver.refresh()
        time.sleep(1)
    print("添加成功")
    url = "https://www.hik-cloud.com/chain/login/index.html#/login"
    driver.get(url)
# 方法主入口
if __name__ == '__main__':
    yunmou()


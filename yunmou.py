# encoding=utf8
from selenium import webdriver
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
import time
import openpyxl


def get_index(column_name, sheet):
    index = 0
    for name in sheet.rows.__next__():

        if name.value == column_name:
            return index
        index += 1


# 读取设备信息
def load_data():
    book = openpyxl.load_workbook('养户摄像头对应表.xlsx')
    sheet = book.worksheets[0]
    get_index("序列号", sheet)
    devices = []
    first_row_flag = 0
    for rows in sheet.rows:
        if first_row_flag == 0:
            first_row_flag = 1
        else:
            single_device = []
            if rows[get_index("序列号", sheet)].value is None:
                break
            single_device.append(rows[get_index("公司", sheet)].value)
            single_device.append(rows[get_index("地区", sheet)].value)
            single_device.append(rows[get_index("算法组名", sheet)].value)
            single_device.append(rows[get_index("序列号", sheet)].value)
            single_device.append(rows[get_index("设备名称", sheet)].value)
            single_device.append(rows[get_index("服务部", sheet)].value)
            # print(single_device)
            devices.append(single_device)

    return devices


# 录入单个设备
def record_device(driver, device_info):
    # 第一层try catch表示如果出错就重新录入一次
    try:
        url = 'https://www.hik-cloud.com/chainX/index.html#/system/equipment/management'
        driver.get(url)
        time.sleep(1)
        title1 = device_info[0]
        title2 = device_info[1]
        title3 = device_info[5]
        print(device_info[3])
        driver.find_elements_by_xpath('//span[@title="' + title1 + '"]/../span')[0].click()
        driver.find_elements_by_xpath('//span[@title="' + title1 + '"]')[0].click()
        time.sleep(3)
        driver.find_elements_by_xpath('//span[@title="' + title2 + '"]/../span')[0].click()
        driver.find_elements_by_xpath('//span[@title="' + title2 + '"]')[0].click()

        time.sleep(2)
        if title3 is not None:
            print(title3)
            try:
                driver.find_elements_by_xpath('//span[@title="' + title3 + '"]')[0].click()
            except:
                pass
            try:
                driver.find_elements_by_xpath('//span[@title="' + title3 + '"]')[-1].click()
            except:
                pass

        time.sleep(1)

        driver.find_element_by_xpath('//i[@class="h-icon-add"]').click()
        time.sleep(1)
        inputs = driver.find_elements_by_xpath(
            '//div[@class="el-form-item__content"]//div[@class="el-input"]//input[@autocomplete="off"]')
        device_num = device_info[3]
        device_verify_code = "12341234q"
        device_name = device_info[4]
        inputs[0].send_keys(device_num)
        inputs[1].send_keys(device_verify_code)
        inputs[2].send_keys(device_name + device_num)
        try:
            driver.find_elements_by_xpath(
                '//div[@class="el-dialog__wrapper"]//div[@class="el-dialog__footer"]//div[@class="dialog-footer"]//button[@class="el-button el-button--primary"]//span')[
                3].click()
        except:
            pass
        time.sleep(1)
        try:
            driver.find_elements_by_xpath(
                '//div[@class="el-dialog__wrapper"]//div[@class="el-dialog__footer"]//div[@class="dialog-footer"]//button[@class="el-button el-button--default"]//span')[
                3].click()
        except:
            pass

        time.sleep(1)
        driver.find_element_by_xpath('//input[@placeholder="设备序列号/名称"]').send_keys(device_num)
        time.sleep(1)
        driver.find_elements_by_css_selector('.el-input__icon.h-icon-search')[1].click()
        time.sleep(4)
        try:
            driver.find_elements_by_css_selector('.el-tooltip.operator.open')[2].click()

            driver.find_element_by_css_selector(
                "#HKClosePasswordDialog > form > div.el-form-item.is-required-right > div > div > input").send_keys(
                "12341234q")
            time.sleep(1)
            driver.find_element_by_xpath('//span[text()="关闭加密"]').click()
            time.sleep(1)
            driver.find_elements_by_xpath('//span[text()="确定"]')[1].click()
            # print("关闭加密成功")
        except:
            pass
        time.sleep(1)
        try:
            driver.find_elements_by_css_selector('.el-tooltip.operator.close')[2].click()
            time.sleep(1)
            driver.find_elements_by_xpath('//span[contains(text(),"确认")]')[1].click()
            # print("开启加密成功")
        except Exception as e:
            print(e)
            print('开启加密出现问题')
        time.sleep(1)
        try:
            driver.find_elements_by_css_selector('.el-tooltip.operator.open')[2].click()

            driver.find_element_by_css_selector(
                "#HKClosePasswordDialog > form > div.el-form-item.is-required-right > div > div > input").send_keys(
                "12341234q")
            time.sleep(1)
            driver.find_element_by_xpath('//span[text()="关闭加密"]').click()
            time.sleep(1)
            driver.find_elements_by_xpath('//span[text()="确定"]')[1].click()
            # print("关闭加密成功")
        except:
            print(e)
            print('关闭加密出现问题')
        time.sleep(1)
    except Exception as e:
        print("出现问题,", "重新录入一次")
        record_device(driver, device_info)


def yunmou():
    # 加启动配置
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')

    # 打开chrome浏览器
    driver = webdriver.Chrome(chrome_options=option)
    driver.maximize_window()
    url = "https://www.hik-cloud.com/chain/login/index.html#/login"
    driver.get(url)
    time.sleep(1)
    driver.find_element_by_xpath('//input[@placeholder="请输入用户名"]').send_keys("ym18363963202")
    driver.find_element_by_xpath('//input[@placeholder="请输入密码"]').send_keys("Buchou123")
    driver.find_element_by_xpath('//div[@class="login-btn-div"]').click()
    time.sleep(3)
    try:
        driver.find_element_by_xpath('//button[@class="el-button el-button--default el-button--primary "]').click()
    except:
        pass
    devices = load_data()
    time.sleep(3)
    for device_info in devices:
        record_device(driver,device_info)

    return driver


def Ai_inspect(driver):
    url = "https://www.hik-cloud.com/AI-inspect/index.html#/intelligent/inspectionConfig/retail"
    driver.get(url)
    devices = load_data()
    name_cache = None
    for device_info in devices:

        sevice_departament = device_info[2]
        # 如果换算法组了，就刷新页面重新选择
        if name_cache != sevice_departament:
            try:
                # 结束当前算法组
                driver.find_element_by_xpath('//span[text()="确 定"]').click()
                time.sleep(0.5)
                driver.find_element_by_xpath('//span[text()="下一步"]').click()
                time.sleep(0.5)
                driver.find_element_by_xpath('//span[text()="完成"]').click()
                time.sleep(0.5)
            except:
                pass
            time.sleep(1)
            url = "https://www.hik-cloud.com/AI-inspect/index.html#/intelligent/inspectionConfig/retail"
            driver.get(url)
            name_cache = sevice_departament
            print(sevice_departament + "算法组正在添加中")
            time.sleep(0.5)
            driver.find_element_by_xpath('//input[@placeholder="请输入任务名称"]').send_keys(sevice_departament)
            time.sleep(0.5)
            driver.find_element_by_css_selector('.el-input__icon.h-icon-search').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//span[text()="' + sevice_departament + '"]').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//span[text()="编辑任务"]').click()
            time.sleep(2)
            driver.find_element_by_xpath('//span[text()="选择"]').click()

        driver.find_element_by_xpath('//input[@placeholder="输入监控点名称"]').clear()
        time.sleep(0.5)
        driver.find_element_by_xpath('//input[@placeholder="输入监控点名称"]').send_keys(device_info[3])
        time.sleep(1)
        driver.find_element_by_css_selector('.el-input__icon.h-icon-search').click()
        time.sleep(1)
        try:
            check_boxes = driver.find_elements_by_css_selector('.el-checkbox')
            # 点击未被选中的checkbox
            for check_box in check_boxes:
                cname = check_box.get_attribute('class')
                if 'is-checked' in cname:
                    pass
                else:
                    check_box.click()
                    break
            print(device_info[3] + '添加成功')
        except Exception as e :
            print(e)
            print(device_info[3] + '此前已被添加')
        time.sleep(0.5)
    # 结束当前服务部
    driver.find_element_by_xpath('//span[text()="确 定"]').click()
    time.sleep(0.5)
    driver.find_element_by_xpath('//span[text()="下一步"]').click()
    time.sleep(0.5)
    driver.find_element_by_xpath('//span[text()="完成"]').click()
    time.sleep(0.5)


# 方法主入口
if __name__ == '__main__':
    driver = yunmou()
    # Ai_inspect(driver)
    c = input()

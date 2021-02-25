from selenium import webdriver
import time
def login_system():    #输入用户名密码登录系统
    browser=webdriver.Firefox()
    browser.get('http://www.farmsoft.com.cn/lord/')
    farm_element=browser.find_element_by_css_selector('.textbox-icon')
    farm_element.click()
    time.sleep(1)
    nkzj_element=browser.find_element_by_css_selector('#_easyui_tree_129 > span:nth-child(1)')
    nkzj_element.click()
    time.sleep(1)
    nc_element=browser.find_element_by_css_selector('#_easyui_tree_143 > span:nth-child(4)')
    nc_element.click()
    user_element=browser.find_element_by_css_selector('.top > div:nth-child(2) > div:nth-child(1) > span:nth-child(2) > input:nth-child(1)')
    user_element.send_keys('Username')#输入用户名
    psw_element=browser.find_element_by_css_selector('.top > div:nth-child(3) > div:nth-child(1) > span:nth-child(2) > input:nth-child(1)')
    psw_element.send_keys('Password')#输入密码
    login_element=browser.find_element_by_css_selector('#login')
    login_element.click()

def turn2Dgzzjl():#进入稻谷种植记录
    time.sleep(1)
    xxjlgl_element=browser.find_element_by_css_selector('div.panel:nth-child(2) > div:nth-child(1) > div:nth-child(1)')
    xxjlgl_element.click()
    dgzzjl_element=browser.find_element_by_css_selector('#_easyui_tree_7 > span:nth-child(3) > div:nth-child(1)')
    dgzzjl_element.click()
#进入填写表单界面，需要打开新的网页
def turn2Newpage():
    js='window.open("http://www.farmsoft.com.cn/lord/BizForm/../Form/DimEdit?formId=19&parentId=");'
    browser.execute_script(js)
    browser.switch_to_window(browser.window_handles[-1])
def FillTheForm():
    #填写表单
    time.sleep(2)
    browser.find_element_by_css_selector('#datagrid-row-r1-2-0 > td:nth-child(2) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > span:nth-child(2) > span:nth-child(1) > a:nth-child(1)').click()
    browser.find_element_by_css_selector('.tree-hit').click()
    time.sleep(5)
    #根据人名排序选择点击人名
    zzh=browser.find_element_by_css_selector('div.combo-p:nth-child(9) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > ul:nth-child(2) > li:nth-child(1)')
    zzh.click()#根据人名排列顺序index（从1到最后）
    time.sleep(1)
    #填入品种
    pz=browser.find_element_by_css_selector('#datagrid-row-r1-2-0 > td:nth-child(4) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > span:nth-child(2) > input:nth-child(2)')
    pz.send_keys('龙粳31')
    #填入日期
    bzrq=browser.find_element_by_css_selector('.datebox > input:nth-child(2)')
    bzrq.send_keys('2020-04-01')
    #填入种植面积
    mianji=browser.find_element_by_css_selector('input.validatebox-invalid:nth-child(1)')
    mianji.send_keys('150')
    #填入计量单位
    jiliangdanwei=browser.find_element_by_css_selector('#datagrid-row-r1-2-0 > td:nth-child(8) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > span:nth-child(2) > span:nth-child(1) > a:nth-child(1)')
    jiliangdanwei.click()
    mushu=browser.find_element_by_css_selector('#_easyui_combobox_i1_0')
    mushu.click()
    #填入种子来源
    zhongzilaiyuan=browser.find_element_by_css_selector('#datagrid-row-r1-2-0 > td:nth-child(10) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > span:nth-child(2) > span:nth-child(1) > a:nth-child(1)')
    zhongzilaiyuan.click()
    #种子来源选择倒数第二项
    zhongzigongsi=browser.find_element_by_css_selector('#_easyui_tree_53')
    zhongzigongsi.click()

def Saveandnew():
    #保存表单并继续
    Save=browser.find_element_by_css_selector('#btnSaveAndNew')
    Save.click()

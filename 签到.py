#西北工业大学疫情自动填报
#By SYJ
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import time
import smtplib
import datetime
from email.mime.text import MIMEText
from email.header import Header

if __name__ == '__main__':
    #这里改成你的统一认证用户名和密码
    user_name = '2018301177'
    pwd = 'syj18761782628'

    # 加上这两句话不打开浏览器
    option = webdriver.ChromeOptions()
    # option.add_argument('headless') # 设置option
    # 用浏览器打开打卡的网址
    browser = webdriver.Chrome(options=option)
    browser.get("http://yqtb.nwpu.edu.cn/wx/xg/yz-mobile/index.jsp")
    # 输用户名和密码
    user_name_input = browser.find_element_by_id("username")
    user_name_input.send_keys(user_name)
    user_pwd_input = browser.find_element_by_id("password")
    user_pwd_input.send_keys(pwd)
    #登录
    login_button = browser.find_element_by_name("submit")
    ActionChains(browser).move_to_element(login_button).click(login_button).perform()
    print('点击登陆')
    #健康登记
    jkdj = browser.find_element_by_xpath("/html/body/div[1]/div[4]/ul/li[1]/a/span")
    ActionChains(browser).move_to_element(jkdj).click(jkdj).perform()
    #当前所在位置 根据实际情况选择 用不上的注释掉
    # 在国内（西安市以外的其他地区）
    weizhi = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[3]/label[3]')
    ActionChains(browser).move_to_element(weizhi).click(weizhi).perform()
    #选择省份地区(我的是江苏省南通市崇川区）

    browser.find_element_by_xpath('//*[@id="province"]/option[11]').click()
    browser.find_element_by_xpath('//*[@id="city"]/option[7]').click()
    browser.find_element_by_xpath('//*[@id="district"]/option[3]').click()

    # #在学校（由学校统一安排住宿的学生）
    # school = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[3]/label[1]')
    # ActionChains(browser).move_to_element(school).click(school).perform()
    #
    # #在西安（租住校内房屋或住在校外的学生）
    # xian = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[3]/label[2]')
    # ActionChains(browser).move_to_element(xian).click(xian).perform()
    #近15天是否前往或经停中高风险地区，或其他有新冠患者病例报告的社区？
    fou1 = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[6]/label[1]')
    ActionChains(browser).move_to_element(fou1).click(fou1).perform()
    #近15天接触过出入或居住在中高风险地区的人员，以及其他有新冠患者病例报告的发热或呼吸道症状患者？
    fou2 = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[8]/label[1]')
    ActionChains(browser).move_to_element(fou2).click(fou2).perform()
    #近15天您或家属是否接触过疑似或确诊患者，或无症状感染患者（核酸检测阳性者）？
    fou3 = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[10]/label[1]')
    ActionChains(browser).move_to_element(fou3).click(fou3).perform()
    #今天的体温范围
    xiaoyu37 = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[12]/label[1]')
    ActionChains(browser).move_to_element(xiaoyu37).click(xiaoyu37).perform()
    #您或家属有无疑似症状?（可多选）
    meiyou = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[14]/label[1]/div[1]')
    ActionChains(browser).move_to_element(meiyou).double_click(meiyou).perform() #此处要双击，否则会取消选择
    #您或家属当前健康状态
    zhenchang = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[17]/label[1]')
    ActionChains(browser).move_to_element(zhenchang).click(zhenchang).perform()
    #您或家属是否正在隔离？（隔离是根据上级单位、医院相关要求的居家或封闭性隔离，宅在家的不属隔离)
    zcxxss = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[20]/label[1]')
    ActionChains(browser).move_to_element(zcxxss).click(zcxxss).perform()
    #提交填报信息
    submit = browser.find_element_by_xpath('//*[@id="rbxx_div"]/div[24]')
    ActionChains(browser).move_to_element(submit).click(submit).perform()
    time.sleep(1)
    # 郑重承诺
    swear = browser.find_element_by_xpath('//*[@id="qrxx_div"]/div[2]/div[25]')
    ActionChains(browser).move_to_element(swear).click(swear).perform()
    # 确认提交
    confirm = browser.find_element_by_partial_link_text('确认提交')
    ActionChains(browser).move_to_element(confirm).click(confirm).perform()
    time.sleep(2)
    #关闭浏览器
    browser.close()

    # 发送给个人邮箱
    # 用于构建邮件头
    # 获取时间
    now_time1 = datetime.datetime.now().strftime('%Y-%m-%d')
    now_time2 = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 发信方的信息：发信邮箱，QQ 邮箱授权码
    from_addr = '1559650057@qq.com'
    # 进入qq邮箱->设置->账户->找到stmp服务，点击开启。验证后会给你一个授权码，直接复制，填入下方即可
    password = 'udyekijuibaahgdg'

    # 收信方邮箱
    to_addr = 'syj20000910@gmail.com'

    # 发信服务器
    smtp_server = 'smtp.qq.com'

    # 邮箱正文内容，第一个参数为内容（正文部分），第二个参数为格式(plain 为纯文本)，第三个参数为编码
    msg = MIMEText('今日已经填报好健康信息,填写时间' + str(now_time2), 'plain', 'utf-8')

    # 邮件头信息
    msg['From'] = Header(from_addr)
    msg['To'] = Header(to_addr)
    msg['Subject'] = Header(str(now_time1) + '疫情自动填报情况')

    # 开启发信服务，这里使用的是加密传输
    server = smtplib.SMTP_SSL(smtp_server)
    server.connect(smtp_server, 465)
    # 登录发信邮箱
    server.login(from_addr, password)
    # 发送邮件
    server.sendmail(from_addr, to_addr, msg.as_string())
    # 关闭服务器
    server.quit()

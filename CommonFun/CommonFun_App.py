# 初始化和封装一下基本方法
import csv
import logging
import logging.config
import os
import smtplib
import time
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import yaml
from appium import webdriver
from appium.webdriver.common.multi_action import MultiAction
from appium.webdriver.common.touch_action import TouchAction
from selenium.webdriver.support.wait import WebDriverWait

base_path = os.path.dirname(os.path.dirname(__file__))
log_config_path = base_path + '/config/log.conf'
logging.config.fileConfig(log_config_path)
logging = logging.getLogger()

caps_yaml_path = base_path + '/config/360_rent_APK_caps.yaml'


class CommonFun_App(object):

    def openApp(self):

        with open(caps_yaml_path, 'r', encoding='utf-8') as file:
            data = yaml.load(file, Loader=yaml.FullLoader)

        desired_caps = {'platformName': data['platformName'], 'platformVersion': data['platformVersion'],
                        'deviceName': data['deviceName'], 'app': data['app'], 'appPackage': data['appPackage'],
                        'appActivity': data['appActivity'], 'noReset': data['noReset'],
                        'unicodeKeyboard': data['unicodeKeyboard'], 'resetKeyboard': data['resetKeyboard']}

        logging.info('Start App')
        self.driver = webdriver.Remote('http://' + str(data['ip']) + ':' + str(data['port']) + '/wd/hub', desired_caps)
        self.driver.implicitly_wait(10)
        return self.driver

    def find_element(self, loc):

        if "=>" not in loc:
            raise NameError("Positioning syntax errors, lack of '=>'.")
        by = loc.split("=>")[0].lower()
        value = loc.split("=>")[1]

        if by == 'id':
            element = self.driver.find_element_by_id(value)
        elif by == 'name':
            element = self.driver.find_element_by_name(value)
        elif by == 'class':
            element = self.driver.find_element_by_class_name(value)
        elif by == 'link_text':
            element = self.driver.find_element_by_link_text(value)
        elif by == 'xpath':
            element = self.driver.find_element_by_xpath(value)
        elif by == 'css':
            element = self.driver.find_element_by_css_selector(value)

        else:
            self.get_Screen_Shot('Fail to find the element')
            logging.error('Fail to find element')
            raise NameError(
                "Please enter the correct targeting elements,'id','name','class','link_text','xpath','css'.")

        return element

    def get_window_size(self):
        """
        获取窗口大小
        :return:窗口长和宽
        """
        x = self.driver.get_window_size()['width']
        y = self.driver.get_window_size()['height']
        logging.info('The windows’ width is %d：' % x + ',height is %d' % y)
        return x, y

    def wait(self, seconds):
        try:
            self.driver.implicitly_wait(int(seconds))
            logging.info('wait for %d seconds.' % seconds)
        except:
            logging.error('Fail to wait')
            self.get_Screen_Shot('Fail to wait')

    def back(self):
        try:
            self.driver.keyevent(4)
        except:
            logging.error('Fail to back')
            self.get_Screen_Shot('Fail to back')

    def long_press(self, loc):
        try:
            element = self.find_element(loc)
            TouchAction(self.driver).long_press(element).perform()
        except:
            logging.error('Fail to long press the element')
            self.get_Screen_Shot('Fail to long press the element')

    def tap(self, x, y):
        try:
            self.driver.tap([(x, y)])
        except:
            logging.error('Fail to tap %s,%s' % (x, y))
            self.get_Screen_Shot("Fail to tap %s,%s" % (x, y))

    def shake(self):
        logging.info('Shake the mobile')
        self.driver.shake()

    def close(self):
        try:
            self.driver.close_app()
            logging.info('Close the window')
        except:
            logging.error('Fail to close the window')
            self.get_Screen_Shot('Fail to close the window')

    def pinch(self):
        try:
            x = self.driver.get_window_size()['width']
            y = self.driver.get_window_size()['height']
            action1 = TouchAction(self.driver)
            action2 = TouchAction(self.driver)
            zoom_action = MultiAction(self.driver)
            action1.press(x=x * 0.2, y=y * 0.2).wait(1000).move_to(x=x * 0.4, y=y * 0.4).wait(1000).release()
            action2.press(x=x * 0.8, y=y * 0.8).wait(1000).move_to(x=x * 0.6, y=y * 0.6).wait(1000).release()
            logging.info('Start pinch')
            zoom_action.add(action1, action2)
            zoom_action.perform()
        except:
            logging.error('Fail to pinch')
            self.get_Screen_Shot('Fail to pinch')

    def zoom(self):
        try:
            x = self.driver.get_window_size()['width']
            y = self.driver.get_window_size()['height']
            action1 = TouchAction(self.driver)
            action2 = TouchAction(self.driver)
            zoom_action = MultiAction(self.driver)
            action1.press(x=x * 0.4, y=y * 0.4).wait(1000).move_to(x=x * 0.2, y=y * 0.2).wait(1000).release()
            action2.press(x=x * 0.6, y=y * 0.6).wait(1000).move_to(x=x * 0.8, y=y * 0.8).wait(1000).release()
            logging.info('Start zoom')
            zoom_action.add(action1, action2)
            zoom_action.perform()
        except:
            logging.error('Fail to zoom')
            self.get_Screen_Shot('Fail to zoom')

    def quit(self):
        try:
            self.driver.quit()
            logging.info('Quit the browser')
        except:
            logging.error('Fail to quit the browser')
            self.get_Screen_Shot('Fail to quit the browser')

    def getTime(self):
        """
        获取当前时间
        """
        try:
            now = time.strftime("%Y-%m-%d %H_%M_%S")
            print(now)
            logging.info('Current time is：' + now)
            return now
        except:
            logging.error('Fail to get current time')
            self.get_Screen_Shot('Fail to get current time')

    def get_Screen_Shot(self, filename):
        """
        截图
        :param filename:截图名称

        """
        try:
            image_file = base_path + '/screenshots/' + filename + '.png'
            logging.info('get screenshot as %s' % filename)
            self.driver.get_screenshot_as_file(image_file)
        except:
            logging.error('Fail to take screenshot')

    def click_element(self, loc):

        try:
            self.wait_element(loc)
            self.find_element(loc).click()
            logging.info('Click the element')
        except:
            logging.error('Fail to click the element')
            self.get_Screen_Shot('Fail to click the element')

    def clear_element(self, loc):

        logging.info('Clear the element')
        try:
            self.wait_element(loc)
            self.find_element(loc).clear()

        except:
            logging.error('Fail to clear the element')
            self.get_Screen_Shot('Fail to clear the element')

    def type_element(self, loc, text):

        # element_name = self.find_element(loc).text
        try:
            self.wait_element(loc)
            self.clear_element(loc)
            self.find_element(loc).send_keys(text)
            logging.info('Type %s  in the element' % text)

        except:
            logging.error('Fail to type %s element' % text)
            self.get_Screen_Shot('Fail to type %s element' % text)

    def swipe(self, direction):

        if direction.lower() == 'left':
            logging.info("Swipe left")
            l = self.get_window_size()
            x1 = int(l[0] * 0.1)
            x2 = int(l[0] * 0.9)
            y = int(l[1] * 0.5)
            self.driver.swipe(x2, y, x1, y, 1000)
        elif direction.lower() == 'right':
            logging.info("Swipe Right")
            l = self.get_window_size()
            x1 = int(l[0] * 0.1)
            x2 = int(l[0] * 0.9)
            y = int(l[1] * 0.5)
            self.driver.swipe(x1, y, x2, y, 1000)
        elif direction.lower() == 'up':
            logging.info("Swipe Up")
            l = self.get_window_size()
            x = int(l[0] * 0.5)
            y1 = int(l[1] * 0.1)
            y2 = int(l[2] * 0.9)
            self.driver.swipe(x, y1, x, y2, 1000)
        elif direction.lower() == 'down':
            logging.info("Swipe Down")
            l = self.get_window_size()
            x = int(l[0] * 0.5)
            y1 = int(l[1] * 0.1)
            y2 = int(l[2] * 0.9)
            self.driver.swipe(x, y2, x, y1, 1000)
        else:
            logging.error('Can not swipe')
            self.get_Screen_Shot('Can not swipe')

    def wait_element(self, loc, seconds=10):

        if "=>" not in loc:
            raise NameError("Positioning syntax errors, lack of '=>'.")
        by = loc.split("=>")[0].lower()
        value = loc.split("=>")[1]

        if by == 'id':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_id(value))
        elif by == 'name':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_name(value))
        elif by == 'class':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_class_name(value))
        elif by == 'link_text':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_link_text(value))
        elif by == 'xpath':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_xpath(value))
        elif by == 'css':
            WebDriverWait(self.driver, seconds).until(lambda x: x.find_element_by_css_selector(value))
        else:
            raise NameError(
                "Please enter the correct targeting elements,'id','name','class','link_text','xpath','css'.")

    def isElementExist(self, loc):

        logging.info('check whether the element is displayed')

        try:
            assert self.find_element(loc).is_displayed()
            print('The element is displayed')
        except AssertionError:
            print('The element is not displayed')
            self.get_Screen_Shot(' The element is not displayed')

    def get_text(self, loc):
        """

        :param loc:
        :return:
        """

        self.wait_element(loc)
        element_text = self.find_element(loc).text
        return element_text

    def isEnabled(self, loc):

        logging.info('check whether the element is enabled')

        try:
            self.wait_element(loc)
            assert self.find_element(loc).is_enabled()
            print('The  element is enabled')
        except AssertionError:
            print('The element is not enabled')
            self.get_Screen_Shot('The element is not enabled')

    def isSelected(self, loc):

        logging.info('Check whether the element is selected')
        try:
            self.wait_element(loc)
            assert self.find_element(loc).is_selected()
            print('The element is selected')
        except AssertionError:
            print('The element is not selected')
            self.get_Screen_Shot('The element is not selected')


    def alterAccept(self):
        logging.info('Accept alter')
        try:
            self.driver.switch_to.alert.accept()
        except:
            logging.error('Can not accept alter')
            self.get_Screen_Shot('Can not accept alter')

    def alterDismiss(self):
        logging.info('Dismiss alter')
        try:
            self.driver.switch_to.alert.dismiss()
        except:
            logging.error('Can not dismiss alter')
            self.get_Screen_Shot('Can not dismiss alter')

    def switchFrame(self, loc):
        logging.info('Enter to frame')
        try:
            self.driver.switch_to.frame(self.find_element(loc))
        except:
            logging.error('Can not enter to frame')
            self.get_Screen_Shot('Can not enter frame')

    def switchFrameOut(self):
        logging.info('Leave frame')
        try:
            self.driver.switch_to.default_content()
        except:
            logging.error('Can not leave the frame')
            self.get_Screen_Shot('Can not leave frame')

    def select_by_value(self, loc, value):
        logging.info('select by value')
        try:
            self.wait_element(loc)
            self.find_element(loc).select_by_value(value)
        except:
            logging.error('Can not select by %s' % value)
            self.get_Screen_Shot('Can not select by %s' % value)

    def get_csv_data(self, csv_file, line):
        logging.info('=====get_csv_data======')
        with open(csv_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            for index, row in enumerate(reader, 1):
                if index == line:
                    return row

    def latest_report(self, report_dir):
        lists = os.listdir(report_dir)
        print(lists)
        lists.sort(key=lambda fn: os.path.getatime(report_dir + '\\' + fn))
        print(lists[-1])
        file = os.path.join(report_dir, lists[-1])
        return file

    def send_mail(self, latest_report):
        f = open(latest_report, 'rb')
        mail_content = f.read()
        f.close()
        yamlfile = base_path + '/config/' + 'email.yaml'
        file = open(yamlfile, encoding='utf-8')
        data = yaml.load(file, Loader=yaml.FullLoader)

        smtpserver = data['smtpserver']

        user = data['user']
        password = data['password']

        sender = data['sender']
        receives = data['receives']

        subject = data['subject']
        body_text = data['body_text']
        filename = data['filename']
        msg = MIMEMultipart()
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = sender
        msg['To'] = receives
        msg['date'] = time.strftime('%a, %d %b %Y %H:%M:%S %z')
        text_msg = MIMEText(body_text, 'plain', 'utf-8')
        msg.attach(text_msg)
        file_msg = MIMEApplication(mail_content)
        file_msg.add_header('content-disposition', 'attchment', filename=filename + '.html')

        msg.attach(file_msg)

        smtp = smtplib.SMTP_SSL(smtpserver, 465)
        smtp.helo(smtpserver)
        smtp.ehlo(smtpserver)
        smtp.login(user, password)

        logging.info("Start send email....")
        smtp.sendmail(sender, receives, msg.as_string())
        smtp.quit()
        logging.info("Send email end")

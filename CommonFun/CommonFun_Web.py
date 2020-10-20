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
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait

base_path = os.path.dirname(os.path.dirname(__file__))
log_config_path = base_path + '/config/log.conf'
logging.config.fileConfig(log_config_path)
logging = logging.getLogger()


class CommonFun_Web(object):

    def open(self, browser_type, url):
        """

        :param browser_type:
        :param url:
        :return:
        """
        logging.info('Open website')
        if browser_type.lower() == 'firefox':
            self.driver = webdriver.Firefox()
            self.driver.get(url)
        elif browser_type.lower() == 'chrome':
            self.driver = webdriver.Chrome()
            self.driver.get(url)
        elif browser_type.lower() == 'ie':
            self.driver = webdriver.Ie()
            self.driver.get(url)
        elif browser_type.lower() == 'chrome_headless':
            option = webdriver.ChromeOptions()
            option.add_argument("headless")
            self.driver = webdriver.Chrome(chrome_options=option)
            self.driver.get(url)
        elif browser_type.lower() == 'edge':
            self.driver = webdriver.Edge()
            self.driver.get(url)
        elif browser_type.lower() == 'opera':
            self.driver = webdriver.Opera()
            self.driver.get(url)
        elif browser_type.lower() == 'phantomjs':
            self.driver = webdriver.PhantomJS()
            self.driver.get(url)
        else:
            self.get_Screen_Shot('fail to open browser')
            logging.error('Fail to open website with %s' % NameError)
            raise NameError(
                "Not found %s browser,You can enter 'ie', 'firefox', 'opera', 'phantomjs', 'edge' or 'chrome'." % browser_type)

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
            self.get_Screen_Shot('fail to find the element')
            logging.error('fail to find element')
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

    def set_window_size(self, wide, high):
        try:
            self.driver.set_window_size(width=int(wide), height=int(high))
            logging.info('Set windows size is %d of width and %d of height' % (wide, high))
        except:
            logging.error('Fail to set windows size as %s of width and %s of height' % (wide, high))

    def maxWindows(self):
        try:
            self.driver.maximize_window()
            logging.info('Set max window')
        except:
            logging.error('Fail to set maxWindows')
            self.get_Screen_Shot('Fail to set maxWindow')

    def forward(self):
        try:
            self.driver.forward()
            logging.info('Click forward on current page.')
        except:
            logging.error('Fail to  click forward')
            self.get_Screen_Shot('Fail to click forward')

    def back(self):
        try:
            self.driver.back()
            logging.info('Click back on current page.')
        except:
            logging.error('Fail to click back')
            self.get_Screen_Shot('Fail to click back')

    def wait(self, seconds):
        try:
            self.driver.implicitly_wait(int(seconds))
            logging.info('wait for %d seconds.' % seconds)
        except:
            logging.error('Fail to wait')
            self.get_Screen_Shot('Fail to wait')

    def close(self):
        try:
            self.driver.close()
            logging.info('Close the window')
        except:
            logging.error('Fail to close the window')
            self.get_Screen_Shot('Fail to close the window')

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
            self.click_element(loc)
            self.clear_element(loc)
            self.find_element(loc).send_keys(text)
            logging.info("Type '%s'  in the element" % text)

        except:
            logging.error("Fail to type '%s' in element" % text)
            self.get_Screen_Shot("Fail to type '%s' element" % text)

    def rightClick(self, loc):

        try:
            self.wait_element(loc)
            ActionChains(self.driver).context_click(self.find_element(loc)).perform()
            logging.info('Right click the element')
        except:
            logging.error('Fail to right click element')
            self.get_Screen_Shot('Fail to right click the element')

    def doubleClick(self, loc):
        try:
            self.wait_element(loc)
            ActionChains(self.driver).double_click(self.find_element(loc)).perform()
            logging.info('Double click the element')
        except:
            logging.error('Fail to double click element')
            self.get_Screen_Shot('Fail to double click element')

    def move_to_element(self, loc):

        try:
            self.wait_element(loc)
            ActionChains(self.driver).move_to_element(self.find_element(loc)).perform()
            logging.info('move to the element')
        except:
            logging.error('Fail to move to the element')
            self.get_Screen_Shot('Fail to move to the element')

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

        logging.info('check whether %s element is displayed')

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

    def get_title(self):
        title = self.driver.title
        try:
            logging.info('The website title is: %s' % title)
            return title
        except:
            logging.error('Fail to get the title of website')
            self.get_Screen_Shot('Fail to get the title of website')

    def get_currentUrl(self):
        current_url = self.driver.current_url
        try:
            logging.info('The current url is: %s' % current_url)
            return current_url
        except:
            logging.error('Fail to get current url')
            self.get_Screen_Shot('Fail to get current url')

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
        body_text = data['Body_text']
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

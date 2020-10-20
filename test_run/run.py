import logging
import time
import unittest
from HTMLTestRunnerNew import HTMLTestRunner

import yaml
from CommonFun.CommonFun_Web import *
from CommonFun.CommonFun_App import *

from CommonFun import CommonFun_App
from CommonFun.CommonFun_App import base_path
from CommonFun.CommonFun_Web import base_path


class run(CommonFun_App):
    report_dir = base_path + '/reports/'
    test_dir = base_path + '/test_case/'
    report_yaml_dir = base_path + '/config/test_report.yaml'
    discover = unittest.defaultTestLoader.discover(test_dir, pattern='Test_*')
    now = time.strftime('%Y-%m-%d %H_%M_%S')
    report_name = report_dir + '/' + now + '_result.html'

    with open(report_yaml_dir, 'r', encoding='utf-8') as file:
        data = yaml.load(file, Loader=yaml.FullLoader)

    with open(report_name, 'wb') as file:
        runner = HTMLTestRunner(stream=file, title=data['title'], tester=data['tester'])
        logging.info('Start run testcase')
        runner.run(discover)
    po = CommonFun_Web()
    latest_report = po.latest_report(report_dir)

    po.send_mail(latest_report)

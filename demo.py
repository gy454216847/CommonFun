from CommonFun.CommonFun_App import CommonFun_App
from CommonFun.CommonFun_Web import CommonFun_Web

# demo = CommonFun_Web()
# demo.open('chrome', 'http://www.baidu.com')
# demo.wait(10)
# class test(unittest.TestCase):
#
#     def __init__(self):
#         self.driver = CommonFun_Web().open('chrome', 'https://www.baidu.com')
#
#     def test_demo(self):
#         self.driver
input_box = 'Id=>kw'
submit = 'id=>su'
po = CommonFun_Web()
po.open('chrome', 'http://www.baidu.com')
po.type_element(input_box, 'selenium')
po.click_element(submit)
app = CommonFun_App()

import unittest
from pyunitreport import HTMLTestRunner
from selenium import webdriver

class HellowWorld(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Chrome(executable_path = r'C:\Users\mruiz\Documents\pythonscripts\chatbot\chromedriver_win32\chromedriver.exe')
        driver = self.driver
        driver.implicitly_wait(10)

    def test_hello_world(self):
        driver = self.driver
        driver.get('http://www.platzi.com')

    def tearDown(self):
        self.driver.quit()

if __name__ == "__main__":
    unittest.main(verbosity = 2, testRunner = HTMLTestRunner(output = 'reportes',report_name = 'hello-world-report'))
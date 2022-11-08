import unittest
from pyunitreport import HTMLTestRunner
from selenium import webdriver

class HellowWorld(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Chrome(executable_path = r'C:\Users\mruiz\Documents\pythonscripts\chatbot\chromedriver_win32\chromedriver.exe')
        driver = cls.driver
        driver.implicitly_wait(10)

    def test_hello_world(cls):
        driver = cls.driver
        driver.get('http://www.platzi.com')

    def test_visit_wiki(cls):
        cls.driver.get('https://www.wikipedia.org')
    
    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

if __name__ == "__main__":
    unittest.main(verbosity = 2, testRunner = HTMLTestRunner(output = 'reportes',report_name = 'hello-world-report'))
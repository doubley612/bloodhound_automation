import os
import subprocess
import getpass
import time
import logging
import glob
from singleton_decorator import singleton
from selenium import webdriver
from selenium.webdriver.remote.errorhandler import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from config import config
from automation_utils import get_domains_in_forest


WEBDRIVER_ARGUMENTS = ['--headless', '--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu',
                       '--remote-debugging-port=9222', r'--user-data-dir=%APPDATA%\BloodHound']
UPLOAD_ICON_ELEMENT = 'fa-tasks'
UPLOAD_FIELD = "//input[@type='file']"
SHARPHOUND_RESULT_POSTFIX = '_BloodHound.zip'
NO_RUNNING_TASK_TEXT = 'INFO: No tasks running with the specified criteria.'
UPLOAD_FINISHED_MESSAGE = 'All uploads finished, closing BloodHound UI.'


@singleton
class BloodHoundUploader(object):
    def __init__(self):
        self._kill_uploader_processes()
        options = webdriver.ChromeOptions()
        options.binary_location = config['bloodhound_binary']
        for argument in WEBDRIVER_ARGUMENTS:
            options.add_argument(argument)
        self.driver = webdriver.Chrome(config['chromedriver_binary'], options=options)
        logging.info('Waiting for BloodHound UI to load.')
        time.sleep(10)

    @staticmethod
    def _latest_results_zip(name):
        zip_file_list = glob.glob(f'{config["sharphound_folder"]}\\*.zip')
        zip_file_list.sort(key=os.path.getctime, reverse=True)
        for zip_file in zip_file_list:
            zip_file_name = zip_file.split('\\')[-1]
            if zip_file_name.endswith("_{}.zip".format(name.lower())) or zip_file_name.endswith("_{}.zip".format(name.upper())):
                return zip_file
        return None

    @property
    def _processes_names(self):
        return [binary.split('\\')[-1] for binary in [config['bloodhound_binary'], config['chromedriver_binary']]]

    def _kill_uploader_processes(self):
        logging.info('Checking if previous uploader processes are running.')
        processes_found = False
        current_user = getpass.getuser()
        for process in self._processes_names:
            result = subprocess.check_output(f'taskkill /f /im {process} /fi "USERNAME eq {current_user}"')
            result = result.decode("utf-8").replace(r'\r\n', '')
            if result != NO_RUNNING_TASK_TEXT:
                processes_found = True
                logging.info(result)
        if not processes_found:
            logging.info('No processes found.')

    def _wait_for_upload_icon(self, interval=10):
        element_found_count = 0
        while element_found_count < 3:
            try:
                self.driver.find_element_by_class_name(UPLOAD_ICON_ELEMENT)
                element_found_count += 1
            except NoSuchElementException:
                element_found_count = 0
            time.sleep(interval)

    def _wait_for_login(self, interval=10):
        element_found_count = 0
        while element_found_count < 3:
            try:
                self.driver.find_element_by_class_name("green-icon-color")
                element_found_count += 1
            except NoSuchElementException:
                element_found_count = 0
            try:
                self.driver.find_element_by_class_name(UPLOAD_ICON_ELEMENT)   # already logged in
                return True
            except NoSuchElementException:
                pass
            time.sleep(interval)
        return False

    def _login(self):
        allready_logged_in = self._wait_for_login()
        if allready_logged_in:
            return
        self.driver.find_element_by_xpath("//form/div/div[2]/input[@class='form-control login-text']").send_keys(config['neo4j_user'])
        self.driver.find_element_by_xpath("//form/div/div[3]/input[@class='form-control login-text']").send_keys(config['neo4j_password'], Keys.RETURN)

    def upload_data(self):
        self._login()
        domains_dict = get_domains_in_forest()
        if domains_dict is None:
            domains_dict = config['DEFAULT_DOMAIN_DICT']
            logging.error("There was a problem with executing 'nltest /domain_trusts' - using default domains")
        for domain_name in domains_dict.keys():
            self._wait_for_upload_icon()
            logging.info('Uploading {} SharpHound results.'.format(domain_name))
            upload_field = self.driver.find_element_by_xpath(UPLOAD_FIELD)
            results_file = self._latest_results_zip(domain_name)
            if results_file is None:
                logging.info('{} zip file could not be found. Moving on.'.format(domain_name))
                continue
            upload_field.send_keys(results_file)
            logging.info(f'Uploading results file: {results_file}')
            self._wait_for_upload_icon(interval=60)
            self.driver.execute_script("document.querySelector(\"input[type='file']\").value = ''")  #fixes appending files and uploading them multiple times
            logging.info(f'Finished uploading {results_file}')
        logging.info(UPLOAD_FINISHED_MESSAGE)
        self.driver.quit()
        time.sleep(5)

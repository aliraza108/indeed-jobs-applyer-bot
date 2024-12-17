from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as webdriver
import time
import openpyxl
from selenium.webdriver.support.select import Select

list_text = []



wb = openpyxl.load_workbook(r'data.xlsx')
sheet = wb.active
profile = sheet['E2'].value
job_search = sheet['A2'].value
city = sheet['B2'].value
# job_contains1 = 
job_contains2 = sheet['C3'].value
job_contains3 = sheet['C4'].value

# words which can not be included in the job title
word1 = sheet['D2'].value
word2 = sheet['D3'].value
word3 = sheet['D4'].value
word4 = sheet['D5'].value
word5 = sheet['D6'].value
word7 = sheet['D7'].value


options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=C:\\Users\\HP\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument('--profile-directory=Default')
# options.add_experimental_option("detach", True)

driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)



driver.get("https://pk.indeed.com/")
job =  driver.find_element(By.ID, "text-input-what")
time.sleep(3)
# job.clear()
job.send_keys(Keys.CONTROL + "a")
job.send_keys(Keys.DELETE)
job.send_keys(job_search)
driver.implicitly_wait(3)       

location =  driver.find_element(By.ID, "text-input-where")
location.send_keys(Keys.CONTROL + "a")
location.send_keys(Keys.DELETE, city, Keys.ENTER)

check_Title = []
check_Title = driver.find_elements(By.CLASS_NAME, "css-dekpa")
driver.maximize_window()

# this is the class of Job Title

for i in check_Title:
        if (job_contains2, job_contains3 in i.text):
                # if(word1, word2 , word3 , word4, word5, word7 not in i.text):
                        i.click()
                        def job_apply_process():    
                                    # Indeed apply button with ID & Xpath
                                try:
                                        apply_btn = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/span/div[4]/div[5]/div[2]/div/section/div/div/div[2]/div[2]/div[1]/div/div[3]/div/div[2]/div/div/span/div/button")
                                        apply_btn.click()
                                except:
                                        print("")

                                # Click on continue btn after resume submit / with Xpath and Class
                                try:
                                        after_resume_btn = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[3]/div/button")
                                        after_resume_btn.click()
                                except:
                                        print("")

                                        
                                
                                # Mid Lavel Questtions 
                                
                                # Are you authorized to work in the United States?
                                try:
                                        Q_1  = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[2]/div/div/fieldset/label[1]")
                                        Q_1.click()
                                except:
                                        print("")

                                try:
                                        continue_after = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[3]/div/button0")
                                        continue_after.click()
                                except:
                                        print("")












                                # adding experience in with class and Xpath
                                try:
                                        Job_Title_Input = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[2]/div/fieldset/div/div[1]/div/div[1]/div/span/input")
                                        Job_Title_Input.send_keys("Python Developer")

                                        companyName_Input = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[2]/div/fieldset/div/div[2]/div/div[1]/div/span/input")
                                        companyName_Input.send_keys("Troodox")
                                except:
                                        print("")
                                # Button of continue after experience and job title with class and Xpath
                                try:
                                        exp_Continue = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[3]/div/button")
                                        exp_Continue.click()
                                except:
                                        print("")







                                # final Submit Button
                                try:
                                        submit = driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/div/div[2]/div[2]/div/div/main/div[3]/div/button")
                                        submit.click()
                                except:
                                        print("")

                                if(driver.current_window_handle == "https://smartapply.indeed.com/beta/indeedapply/form/post-apply"):
                                        print("Yessssssss we Applied and found last page")   
        job_apply_process()                             
print("bot runs final")
# time.sleep(1000)
driver.quit()
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from selenium.webdriver.support.select import Select

# Fatching Data From Excel File Ema Karter
wb = openpyxl.load_workbook(r'data.xlsx')
sheet = wb.active
profile = sheet['E2'].value
job_search = sheet['A2'].value
city = sheet['B2'].value
job_contains1 = sheet['C2'].value
job_contains2 = sheet['C3'].value
job_contains3 = sheet['C4'].value

# words which can not be included in the job title
word1 = sheet['D2'].value
word2 = sheet['D3'].value
word3 = sheet['D4'].value
word4 = sheet['D5'].value
word5 = sheet['D6'].value
word7 = sheet['D7'].value

# str(job_contains3,job_contains1, job_contains2)

options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=C:\\Users\\HP\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument('--profile-directory=Default')
driver = webdriver.Chrome(options=options, use_subprocess=True)



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

# Class of UL
Li_jobs_cards = []
check_Title = []
Job_title = []
# this is the class of li cards of jobs
Li_jobs_cards = driver.find_elements(By.CLASS_NAME, "css-5lfssm")
check_Title = driver.find_elements(By.CLASS_NAME, "css-dekpa")

# this is the class of Job Title

for i in check_Title:
        if (job_contains1, job_contains2, job_contains3 in i.text):
                if(word1, word2 , word3 , word4, word5, word7 not in i.text):
                        i.click()
                        
                        try:
                                def job_apply_process():
                                        # Indeed Apply Button
                                        try:
                                                apply_btn = driver.find_element(By.ID,"indeedApplyButton")
                                                apply_btn.click()
                                                time.sleep(5)
                                        except:
                                                print("")
                                        try:
                                                driver.switch_to.window(driver.window_handles[1])
                                                continu_BTN = driver.find_element(By.CLASS_NAME, "7ca0508c20f5a21bf434cbc283a01e32365e1d8f93833c7133eadd015a12d2d3")
                                                continu_BTN.click()
                                        except:
                                                print("")

                                        # continue button after resume with interview date 2-3 option
                                        try:    
                                                continue_btn_opt = driver.find_element(By.CLASS_NAME, "8f2157bb6b37636a29a4ac60bd03bb72091b86ccf5f62147ece6dd0d68bf8753")
                                                continue_btn_opt.click()
                                        except:
                                                print("")
                                                
                                        try:
                                                continueBtn = driver.find_element(By.CLASS_NAME, "1c111d09109bcb9d78c39640b659cd232377fb09c2b176bdb18eaf688872d80a")
                                                continueBtn.click()
                                        except:
                                                print("")
                                        # try conditiosns

                                        # for Experience box
                                        try: 
                                                experience = driver.find_element(By.ID, "input-q_371d7bc37e4dc3248e6a5e1a940ad3ca")
                                                experience.send_keys(2) 
                                        except:
                                                print("experience except")

                                        # experience Box 2 
                                        try:
                                                exp2 = driver.find_element(By.ID, "input-q_2f9e867c2a89016afab1a03d3c67849a")
                                                exp2.send_keys(2)
                                        except:
                                                print("")
                                        # API Intergration Experience 
                                        try:
                                                api_experience = driver.find_element(By.ID,"input-q_2f9e867c2a89016afab1a03d3c67849a")
                                                api_experience.send_keys(2)
                                        except:
                                                print("")
                                        try:
                                                edu_dropdown = Select(driver.find_element(By.ID, "input-q_728ef64b2478d3a57679e47cc060b743")) 
                                                edu_dropdown.select_by_value("Intermediate")
                                        except: 
                                                print("")
                                        try:
                                                continue_BTN = driver.find_element(By.CLASS_NAME, "9842d54b7991a3ad744dda86c70be9998b3be821c1ad1982ef21db209c8d8a2b")
                                                continue_BTN.click()
                                        except:
                                                print("")


                                        try: 
                                                continue_second = driver.find_element(By.CLASS_NAME, "263e9f6f5d498fe0cdd3de932c4d1c670eba41407f17fd7cefdeab6086e6fff2")
                                                continue_second.click()
                                        except:
                                                print("")
                                                #drop down manu of notice period
                                        try: 
                                                notice_period = Select(driver.find_element(By.CLASS_NAME, "input-q_f60c74d2cf3a1c86a2c749d5c35a9ec1"))
                                                notice_period.select_by_index(2)
                                        except:
                                                print("")
                                        # final button with experience company name and desination
                                        try:
                                                jobTitle = driver.find_element(By.ID,"jobTitle")
                                                jobTitle.send_keys(Keys.CONTROL + "a")
                                                jobTitle.send_keys("Python Developer")
                                                companyName = driver.find_element(By.ID, "companyName")
                                                companyName.send_keys(Keys.CONTROL + "a")
                                                companyName.send_keys("Troodox")
                                        except:
                                                        print("")
                                        try:
                                                # click on submit btn
                                                btn_exp = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/div/div[2]/div[2]/div/div/main/div[3]/div/button")
                                                btn_exp.click()
                                                driver.implicitly_wait(5)
                                        except:
                                                print("")
                                        try:    
                                                final_continue = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div/div/div[2]/div[2]/div/div/main/div[3]/div/button")
                                                final_continue.click()
                                        except:
                                                print("")
                                        try:    
                                                final_continue = driver.find_element(By.CLASS_NAME, "b1dd560845dd35c62cd32d2eda4ecc33b039fdd170fe1e4a75307c58d31a1676")
                                                final_continue.click()
                                        except:
                                                print("")

                                        
                                        # driver.close()
                                        time.sleep(5)
                                        driver.switch_to.window(driver.window_handles[0])
                                                        # time.sleep(2)  # Adding a small delay to ensure the element is loaded properly

                                        try:
                                                submit_btn = driver.find_element(By.CLASS_NAME, "263e9f6f5d498fe0cdd3de932c4d1c670eba41407f17fd7cefdeab6086e6fff2").click() 
                                        except:
                                                print("")

                                # run while loop for rerun the casses in case it re assaign the 
                                a = 0
                                while (a<4):
                                        job_apply_process()
                                        a+=1
                                # break
                        except:
                                print("")
        else:
                print("")

print("bot runs final")
time.sleep(1000)
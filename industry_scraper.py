from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from linkedin_scraper import actions
from bs4 import BeautifulSoup
import time, openpyxl, random, csv
pause_time = random.randint(15, 30)
pause_time2 = random.randint(5, 10)

# # Get past LinkedIn authentication wall
# driver = webdriver.Chrome('/Users/alex/Downloads/chromedriver')
# actions.login(driver, email='kawaiipandas@yahoo.com', password='Ale123yan')
# # actions.login(driver, email='alexyan.isnt@gmail.com', password='#Swagmaster@13x')
# # actions.login(driver, email='meiko@inawe.com', password='Hello123')
# print(f'''URL: {driver.command_executor._url}
# session ID: {driver.session_id}''')

driver = webdriver.Remote('http://127.0.0.1:51428')
driver.close()
driver.session_id = 'd8d251f108f0523b0e801a06b73d51ae'

wait = WebDriverWait(driver, 10)

# Reference spreadsheet
wb = openpyxl.load_workbook('/Users/alex/Downloads/Current Orgs.xlsx')
sheet = wb['Data']
new_file = open('/Users/alex/Downloads/Organizations Industries Scraped.csv', 'w', newline='')
nf_writer = csv.DictWriter(new_file, ['Organization ID', 'Organization Name', 'LinkedIn Profile', 'Industry', 'LI Industry'])
nf_writer.writeheader()

for i in range(20, 351):
    org_id = sheet.cell(row=i, column=1).value
    org_name = sheet.cell(row=i, column=2).value
    li_url = sheet.cell(row=i, column=4).value
    industry = sheet.cell(row=i, column=3).value
    li_industry = None
    if li_url is None or 'http' not in li_url or 'linkedin' not in li_url:
        li_industry = 'NONE'
        # nf_writer.writerow({'Organization ID': org_id, 'Organization Name': org_name, 'LinkedIn Profile': li_url, 'Industry': industry, 'LI Industry': ' '})
        # continue
    elif 'school' in li_url:
        li_industry = 'Higher Education'
        # nf_writer.writerow({'Organization ID': org_id, 'Organization Name': org_name, 'LinkedIn Profile': li_url, 'Industry': industry, 'LI Industry': 'Higher Education'})
    elif 'showcase' in li_url or 'company' in li_url:
        driver.get(li_url)
        if 'unavailable' in str(driver.current_url) or '/pub/' in str(driver.current_url) or '/feed/' in str(driver.current_url):
            li_industry = 'NONE'
            # nf_writer.writerow({'Organization ID': org_id, 'Organization Name': org_name, 'LinkedIn Profile': li_url,
            #                     'Industry': industry, 'LI Industry': 'NONE'})
            # time.sleep(pause_time)
            # continue
        else:
            if 'Something went wrong' in driver.page_source:
                li_industry = 'None'
            else:
                wait_page = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[1]/section')))
                html = driver.page_source
                soup = BeautifulSoup(html, 'lxml')
                summary_info_div = soup.find('div', {'class': 'org-top-card-summary-info-list__info-item'})
                try:
                    li_industry = summary_info_div.get_text().strip()
                except AttributeError:
                    li_industry = 'NONE'

    if industry != li_industry:
        nf_writer.writerow(
            {'Organization ID': org_id, 'Organization Name': org_name, 'LinkedIn Profile': li_url, 'Industry': industry,
             'LI Industry': li_industry})
        print(org_name)
        time.sleep(pause_time)
    elif industry == li_industry:
        time.sleep(pause_time2)
        continue

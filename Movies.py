from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.chrome.options import Options

#Set webdriver path
PATH = "C:\Program Files (x86)\chromedriver.exe"
options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(PATH,chrome_options=options)

names = []
BoxOffices = []
years = []
hrefs = []
company = []
genre1 = []
genre2 = []
genre3 = []
genre4 = []
page = list(range(0,900,200))
#Loop through 5 pages
for p in page:
    url = 'https://www.boxofficemojo.com/chart/top_lifetime_gross_adjusted/?adjust_gross_to=2022&offset={}'.format(p)
    driver.get(url)
    #Get the link for each movie
    elems = driver.find_elements(By.CSS_SELECTOR, 'a.a-link-normal')
    links = [elem.get_attribute('href') for elem in elems]
    prefix = 'https://www.boxofficemojo.com/title/tt'

    #Extract and append the name, box office, and year from every movie
    for i in range(2,202):
        name = driver.find_elements(By.XPATH, '//*[@id="table"]/div/table[2]/tbody/tr[{}]/td[2]/a'.format(i))[0]
        names.append(name.text)
        BoxOffice = driver.find_elements(By.XPATH, '//*[@id="table"]/div/table[2]/tbody/tr[{}]/td[3]'.format(i))[0]
        BoxOffices.append(BoxOffice.text)
        year = driver.find_elements(By.XPATH, '//*[@id="table"]/div/table[2]/tbody/tr[{}]/td[6]'.format(i))[0]
        years.append(year.text)
    #Append links
    for link in links:
        if prefix in link:
            hrefs.append(link)
    #Get the company name
    for href in hrefs:
        driver.get(href)
        t = driver.find_elements(By.XPATH, "//*[@id='a-page']/main/div/div[3]/div[4]/div[1]/span[2]")[0]
        h = t.text.replace('\nSee full company information', '')
        company.append(h)
        print(h)
        results = driver.find_elements(By.XPATH, "//*[@id='a-page']/main/div/div[3]/div[4]")[0]
        x = results.text.split("\n")
        genre = "Genres"
        if genre in x:
            n = x.index(genre)
            num = n + 1
            type = x[num]
            x = type.split()
            try:
                genre1.append(x[0])
                genre2.append(x[0])
                genre3.append(x[0])
                genre4.append(x[0])
            except:
                genre4.append("NA")


print(len(names))
print(len(BoxOffices))
print(len(years))
print(len(hrefs))
print(len(company))

#Combine all columns to dataframe & export to Excel
data = {'Movie_Name': names,
        'Box_Office': BoxOffices,
        'Year': years,
        'Company':company,
        'Genre1':genre1,
        'Genre2':genre2,
        'Genre3':genre3,
        'Genre4':genre4,

        }
df = pd.DataFrame(data)

df.to_excel('enter your file path here', index=False)

import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException  
from bs4 import BeautifulSoup

class Scraper:
  def __init__(self):
    print("Init:     Scraper")
    
    # initialize Chrome driver
    self.driver = webdriver.Chrome()

    # url to scrape from
    self.url = "https://my.sunysullivan.edu/ICS/default.aspx?portlet=Course_Schedules&screen=Advanced+Course+Search&screenType=next"
  
  def getDepartment(self, dep="ALL"):
    # load main page
    self.driver.get(self.url)

    # find department dropdown select
    department_dropdown = self.getSelect(By.ID, "pg0_V_ddlDept")

    if department_dropdown == None:
      print("Error:      No Department select found")
      return None

    if dep == "ALL":
      # set select value to index 0 (ALL)
      department_dropdown.select_by_index(0)
    else:
      # set select value to dep
      department_dropdown.select_by_value(dep)

    # find division dropdown select
    division_dropdown = self.getSelect(By.ID, "pg0_V_ddlDivision")

    if division_dropdown == None:
      print("Error:      No Division select found")
      return None

    # set select value to UG (Undergrauate)
    division_dropdown.select_by_value("UG")

    # find search button
    search_button = self.getElement(By.ID, "pg0_V_btnSearch")

    if search_button == None:
      print("Error:      No Search button found")
      return None

    # submit search
    search_button.click()

    # wait for results to populate
    time.sleep(3)

    return self.getClasses()
  
  def getClasses(self):
    # get page source
    source = self.driver.page_source

    # initialize soup from source
    soup = BeautifulSoup(source, 'html.parser')

    # select courses table from soup
    table = soup.select_one("#pg0_V_dgCourses")

    # get a list of dataframe from table source (HTML)
    dataframes = pd.read_html(str(table))

    # pop first dataframe as main dataframe
    # all remaining dataframes will be Textbooks dataframes
    df = dataframes.pop(0)

    # there is not 1 dataframe remaining per row in the main dataframe.
    # fill Textbook for each row in main dataframe with all Book Titles
    # for its corresponding textbook dataframe
    for i in range(len(dataframes)):
      # get textbook dataframe at index i
      textbook_df = dataframes[i]

      # remove NaN values
      textbook_df = textbook_df[pd.isna(textbook_df["Book Title"]) == False]

      # get Book Titles as strings
      title_df = textbook_df["Book Title"].astype(str)

      # set Textbook for main dataframe at index i to list of Book Title values
      df.iloc[i, 1] = ", ".join(title_df.values)

    # find letter navigation if present
    letter_nav = self.getElement(By.CLASS_NAME, "letterNavigator")

    # return df now if no letter navigation exist
    if letter_nav == None:
      return df
    
    # find all links in the letter navigation
    links = letter_nav.find_elements(By.TAG_NAME, "a")

    # check is Next page link is in links
    if "Next" in links[-1].get_attribute("innerHTML"):
      # click next link
      links[-1].click()

      # wait for page to load
      time.sleep(3)

      # make recursive call and append results to main dataframe
      df = df.append(self.getClasses(), ignore_index=True)
    
    return df
  
  def getSelect(self, by, value):
    select = self.getElement(by, value)
    if select == None:
      return select
    
    return Select(select)
  
  def getElement(self, by, value):
    try:
      element = self.driver.find_element(by, value)
    except NoSuchElementException:
      return None
      
    return element
  
  def close(self):
    print("Close:    Scraper")

    # close the Chrome driver
    self.driver.quit()
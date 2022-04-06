from Scraper import Scraper

if __name__ == "__main__":
  print("----------------------")
  scraper = Scraper()
  print("----------------------")

  # get dataframe for department
  df = scraper.getDepartment()

  scraper.close()

  print("----------------------")

  # export dataframe 
  df.to_excel("Scraped.xlsx", index=False)

  print("Export:   Scraped data exported to /Scraped.xlsx")
  print("----------------------")
import os
import pandas as pd
import numpy as np

from Cleaner import Cleaner
from Exporter import Exporter

# Change to data directory
os.chdir("data")

if __name__ == "__main__":
	print("----------------------")
	cleaner = Cleaner()
	print("----------------------")
	exporter = Exporter()
	print("----------------------")

	df = pd.DataFrame()

	# iterate through all files
	for file in os.listdir():
		# Check whether file is in xlsx format or not
		if file.endswith(".xlsx"):
			# ignore ~ files
			if file.startswith("~"):
				continue

			# get cleaned dataframe
			cleaned = cleaner.cleanFile(file)

			if cleaned is not None:
				# add df to combined df
				df = df.append(cleaned, ignore_index=True)

				# export excel file to cleaned data folder
				exporter.exportXLSX(cleaned, "../cleaned data/" + file)

				print("Export:   Cleaned data exported to /cleaned data/" + file)
			
			print("----------------------")
	
	# export combined excel file
	exporter.exportXLSX(df, "../Master.xlsx")

	print("Export:   Cleaned data exported to /Master.xlsx")
	print("----------------------")

	# export sql file
	exporter.exportSQL(df, "../db.sql")

	print("Export:   Cleaned data exported to /db.sql")
	print("----------------------")
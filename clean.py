import os
import pandas as pd
import numpy as np

from Cleaner import Cleaner

# Change to data directory
os.chdir("data")

if __name__ == "__main__":
	print("----------------------")
	cleaner = Cleaner()
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

				# export to cleaned data folder
				cleaned.to_excel("../cleaned data/" + file, index=False)

				print("Export:   Cleaned data exported to /cleaned data/" + file)
			
			print("----------------------")
	
	# export combined
	df.to_excel("../Master.xlsx", index=False)

	print("Export:   Cleaned data exported to /Master.xlsx")
	print("----------------------")
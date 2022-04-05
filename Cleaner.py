import datetime
import pandas as pd
import numpy as np

class Cleaner:
	def __init__(self):
		print("Init:     Cleaner")

		self.blocks = pd.read_excel("../Blocks.xlsx", header=None)
		self.blocks.columns = ["Block", "Day", "Start", "End"]

		# strip whitespace from string columns
		self.blocks[["Block", "Day"]] = 	self.blocks[["Block", "Day"]].apply(lambda x: x.str.strip())

		# uppercase columns
		self.blocks[["Block", "Day"]] = 	self.blocks[["Block", "Day"]].apply(lambda x: x.str.upper())

		print(self.blocks)
	
	def cleanFile(self, file):
		print("Cleaning:", file)

		# get file as dataframe
		df = pd.read_excel(file, header=2)

		# remove empty rows
		df = df[df["Course"].notna()]

		# remove SAMPLE rows
		df = df[df["Block"] != "SAMPLE"]

		# rename Dep, Course, and Section columns
		for i in range(len(df.columns)):
			if df.columns[i] == "Course":
				df.rename({"Course": "Dep"}, axis=1, inplace=True)
				df.rename({"Unnamed: " + str(i+1): "Course"}, axis=1, inplace=True)
				df.rename({"Unnamed: " + str(i+2): "Section"}, axis=1, inplace=True)
				break
		
		# rename Max Size and Comment I columns
		df.rename({"Max Size": "Size"}, axis=1, inplace=True)
		df.rename({"Comment I": "Comments"}, axis=1, inplace=True)

		# convert columns to correct type
		df.loc[pd.isna(df["Bldg"]), "Bldg"] = "TBD"
		df.loc[pd.isna(df["Loc"]), "Loc"] = "TBD"
		df.loc[pd.isna(df["Rm"]), "Rm"] = "TBD"

		# convert columns to correct type
		df["Course"] = df["Course"].astype(int).astype(str)
		df["Size"] = df["Size"].astype(int)
		df["Rm"] = df["Rm"].astype(str)

		# strip whitespace from string columns
		string_columns = ["Block", "Dep", "Course", "Section", "Title", "Instructor", "Day", "Bldg", "Loc", "Rm", "Comments"]
		df[string_columns] = df[string_columns].apply(lambda x: x.str.strip())

		# uppercase columns
		upper_columns = ["Block", "Dep", "Section", "Day", "Bldg", "Loc"]
		df[upper_columns] = df[upper_columns].apply(lambda x: x.str.upper())

		# remove unused columns
		all_columns = ["Block", "Dep", "Course", "Section", "Title", "Instructor", "Day", "Start", "End", "Size", "Bldg", "Loc", "Rm", "Comments"]
		df = df[all_columns]

		# find indices where Day, Start, End, Bldg, Rm, or Loc is Off Campus or Online
		indices = (df["Day"].str.upper() == "OFF CAMPUS") | (df["Day"].str.upper() == "ONLINE") | \
			(df["Start"].str.upper() == "OFF CAMPUS") | (df["Start"].str.upper() == "ONLINE") | \
			(df["End"].str.upper() == "OFF CAMPUS") | (df["End"].str.upper() == "ONLINE") | \
			(df["Bldg"].str.upper() == "OFF CAMPUS") | (df["Bldg"].str.upper() == "ONLINE") | \
			(df["Rm"].str.upper() == "OFF CAMPUS") | (df["Rm"].str.upper() == "ONLINE") | \
			(df["Loc"].str.upper() == "OFF CAMPUS") | (df["Loc"].str.upper() == "ONLINE")
		
		# erase Day, Start, End, Bldg, Rm, and Loc for these indices
		df.loc[indices, ["Day", "Start", "End", "Bldg", "Rm", "Loc"]] = ""
		
		# set Block for these indices to "Z"
		df.loc[indices, "Block"] = "Z"

		# erase NaN values
		df = df.fillna("")

		# reset row index after drops
		df = df.reset_index(drop=True)

		# validate rows
		for row_index in range(df.shape[0]):
			# validate Block
			try:
				df.loc[row_index, "Block"] = self.getBlock(df.loc[row_index]["Block"])
			except:
				print("ERROR:    Invalid Block (", df.loc[row_index]["Block"], ")")
				print(df.loc[row_index])
				return None
      
      # validate Dep
			try:
				df.loc[row_index, "Dep"] = self.getDepartment(df.loc[row_index]["Dep"])
			except:
				print("ERROR:    Invalid Dep (", df.loc[row_index]["Dep"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Section
			try:
				df.loc[row_index, "Section"] = self.getSection(df.loc[row_index]["Section"])
			except:
				print("ERROR:    Invalid Section (", df.loc[row_index]["Section"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Title
			try:
				df.loc[row_index, "Title"] = self.getTitle(df.loc[row_index]["Title"])
			except:
				print("ERROR:    Invalid Title (", df.loc[row_index]["Title"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Instructor
			try:
				df.loc[row_index, "Instructor"] = self.getInstructor(df.loc[row_index]["Instructor"])
			except:
				print("ERROR:    Invalid Instructor (", df.loc[row_index]["Instructor"], ")")
				print(df.loc[row_index])
				return None
			
			# online classes don't need time/location validation
			if self.isOnline(df.loc[row_index]["Block"]):
				continue

			# validate Day
			try:
				df.loc[row_index, "Day"] = self.getDay(df.loc[row_index]["Day"])
			except:
				print("ERROR:    Invalid Day (", df.loc[row_index]["Day"], ")")
				print(df.loc[row_index])
				return None

			# validate Start time
			try:
				df.loc[row_index, "Start"] = self.getTime(df.loc[row_index]["Start"])
			except:
				print("ERROR:    Invalid Start (", df.loc[row_index]["Start"], ")")
				print(df.loc[row_index])
				return None
			
			# validate End time
			try:
				df.loc[row_index, "End"] = self.getTime(df.loc[row_index]["End"])
			except:
				print("ERROR:    Invalid End (", df.loc[row_index]["End"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Building
			try:
				df.loc[row_index, "Bldg"] = self.getBuilding(df.loc[row_index]["Bldg"])
			except:
				print("ERROR:    Invalid Building (", df.loc[row_index]["Bldg"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Location
			try:
				df.loc[row_index, "Loc"] = self.getLocation(df.loc[row_index]["Loc"])
			except:
				print("ERROR:    Invalid Location (", df.loc[row_index]["Loc"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Room
			try:
				df.loc[row_index, "Rm"] = self.getRoom(df.loc[row_index]["Rm"])
			except:
				print("ERROR:    Invalid Room (", df.loc[row_index]["Rm"], ")")
				print(df.loc[row_index])
				return None
			
			# validate Block based on validated Day, Start, and End
			# TODO: resolve issues where mulitple Blocks match for a given Day, Start, Time
			matching_blocks = self.blocks[(self.blocks["Day"] == df.loc[row_index, "Day"]) & (self.blocks["Start"] == df.loc[row_index, "Start"]) & (self.blocks["End"] == df.loc[row_index, "End"])]
			if matching_blocks.shape[0] > 0:
				if df.loc[row_index, "Block"] not in matching_blocks["Block"].values:
					print("WARNING:  Wrong Block for Day, Start, End (", df.loc[row_index, "Block"], ") should be (", matching_blocks.iloc[0, 0], ")")
					df.loc[row_index, "Block"] = matching_blocks.iloc[0, 0]
			else:
				print("WARNING:  Day, Start, End is not a valid Block (", df.loc[row_index, "Day"], df.loc[row_index, "Start"], df.loc[row_index, "End"], ")")
			
		return df
	
	def isOnline(self, loc):
		if pd.isna(loc) or loc == "":
			return False
		
		return loc== "Z" 
	
	def getBlock(self, loc):
		if pd.isna(loc) or loc == "":
			return ""

		if not loc in self.blocks["Block"].values:
			print("WARNING: ", "Invalid Block (", loc, ")")
			return ""
		
		return loc
  
	def getDepartment(self, loc):
		if pd.isna(loc) or loc == "":
			raise KeyError

		return loc
	
	def getSection(self, loc):
		if pd.isna(loc) or loc == "":
			raise KeyError
		
		sec = loc.replace(" ", "")
		if len(sec) != 4:
			raise KeyError
		
		return sec[:2] + " " + sec[2:]
	
	def getTitle(self, loc):
		if pd.isna(loc) or loc == "":
			raise KeyError
		
		return loc
	
	def getInstructor(self, loc):
		if pd.isna(loc) or loc == "":
			print("WARNING: ", "Invalid Instructor (", loc, ")")
			return "TBD"
		
		return loc

	def getDay(self, loc):
		if pd.isna(loc) or loc == "":
			raise KeyError

		if not loc in self.blocks["Day"].values:
			print("WARNING: ", "Invalid Day (", loc, ")")

		return loc
	
	def getTime(self, loc):
		if pd.isna(loc) or loc == "":
			raise KeyError

		if isinstance(loc, datetime.time):
			return loc
		else:
			return datetime.datetime.strptime(loc.replace(" ", ""), "%I:%M%p").time()
	
	def getBuilding(self, loc):
		if pd.isna(loc) or loc == "":
			print("WARNING: ", "Invalid Bldg (", loc, ")")
			return "TBD"
		
		return loc
	
	def getLocation(self, loc):
		if pd.isna(loc) or loc == "":
			print("WARNING: ", "Invalid Loc (", loc, ")")
			return "TBD"
		
		return loc
	
	def getRoom(self, loc):
		if pd.isna(loc) or loc == "":
			print("WARNING: ", "Invalid Rm (", loc, ")")
			return "TBD"
		
		return loc

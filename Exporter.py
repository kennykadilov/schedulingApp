import datetime
import pandas as pd
import numpy as np

class Exporter:
    def __init__(self):
        print("Init:     Exporter")

    def exportXLSX(self, df, fileName):
        # create xlsx file from df
        df.to_excel(fileName, index=False)
    
    def exportSQL(self, df, fileName):
        # create sql string for creating table
        sql = '''
CREATE TABLE `fall_2022` (
    `block` varchar(16),
    `dep` varchar(16) NOT NULL,
    `course` varchar(16) NOT NULL,
    `section` varchar(16) NOT NULL,
    `title` varchar(256) NOT NULL,
    `instructor` varchar(64) NOT NULL,
    `day` varchar(16) NOT NULL,
    `start` time,
    `end` time,
    `size` int NOT NULL DEFAULT 0,
    `bldg` varchar(16) NOT NULL,
    `loc` varchar(16) NOT NULL,
    `rm` varchar(16) NOT NULL,
    `comments` varchar(1024)
);
'''

        # add sql string for inserting values
        sql += "INSERT INTO `fall_2022` VALUES\n"

        # get 2d array of values for df
        rows = df.values.tolist()

        # iterate over each row of the rows
        for row in range(len(rows)):
            # add sql string for row values
            sql += "("

            # iterate over each col of the row
            for col in range(len(rows[row])):
                # get value of col as a string
                strValue = str(rows[row][col])

                # if strValue is an empty string ("") return NULL,
                # otherwise return strValue wrapped in single quotes
                sql += "NULL" if strValue == "" else "'" + strValue + "'"

                # add comma if not the last col
                if col != len(rows[row]) - 1:
                    sql += ", "
            
            # end sql string for row values.
            # add comma if not the last row
            if row == len(rows) - 1:
                sql += ");\n"
            else:
                sql += "),\n"
        
        # create sql file from sql string
        db = open(fileName, "w")
        db.write(sql)
        db.close
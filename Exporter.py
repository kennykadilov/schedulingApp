import datetime
import pandas as pd
import numpy as np

class Exporter:
    def __init__(self):
        print("Init:     Exporter")

    def exportXLSX(self, df, fileName):
        df.to_excel(fileName, index=False)

    def getValueString(self, value, isLast):
        return "'" + str(value) + "'" + ("" if isLast else ", ")
    
    def exportSQL(self, df, fileName):
        sql = '''
CREATE TABLE fall_2022 (
    block varchar(16),
    dep varchar(16) NOT NULL,
    course varchar(16) NOT NULL,
    section varchar(16) NOT NULL,
    title varchar(256) NOT NULL,
    instructor varchar(64) NOT NULL,
    day varchar(16) NOT NULL,
    start time,
    end time,
    size int NOT NULL DEFAULT 0,
    bldg varchar(16) NOT NULL,
    loc varchar(16) NOT NULL,
    rm varchar(16) NOT NULL,
    comments varchar(1024)
);

INSERT INTO fall_2022 VALUES
'''
        rows = df.values.tolist()
        for row in range(len(rows)):
            sql += "("
            for col in range(len(rows[row])):
                sql += self.getValueString(rows[row][col], col == len(rows[row]) - 1)

            if row == len(rows) - 1:
                sql += ");\n"
            else:
                sql += "),\n"

        db = open(fileName, "w")
        db.write(sql)
        db.close
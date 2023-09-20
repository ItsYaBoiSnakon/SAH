#GB_1-7-2023_15-7-2023.xls.xlsx
#custom_1-8-2023 13_27_32

import pandas as pd
import os
import csv
import datetime
import matplotlib.pyplot as plt

class Main:

    def __init__(self, filename="GB_1-7-2023_15-7-2023.xls.xlsx", sheetname="custom_1-8-2023 13_27_32"):
        self.data = []
        self.filename = filename
        self.sheetname = sheetname
        self.csvFileName = "CSV.csv"
        pass

    def run(self):
        self.createCSV()
        self.CSVtoArray()
        
        data = self.arrayToFinalData()
        
        self.printFinalData(data)
        self.printAllData(data)
        self.makeBarChart(data)

    def createCSV(self):
        if not os.path.isfile("CSV.csv"):
            excel = pd.read_excel(self.filename, sheet_name=self.sheetname)
            excel.to_csv (self.csvFileName, index = None, header=True)
            print("not found")
            return True
        print("found")
        return False

    
    # Deze functie maakt een array van de CSV, en print vervolgens de inhoud.
    # ///// DATA LOCATIONS //////
    # 0  :  code of object. 
    #       206 = geen beschikbaarheid
    #       205 = geen beschikbaarheid + whatsappverzoek
    # 1  :  datum van object
    # 2  :  tijd van object
    # 3  :  Werknemer die object heeft gemaakt
    # 4  :  Achternaam Klant
    # 5  :  Klantnummer
    # 6  :  Adres + Postcode Klant
    # 7  :  Stad Klant
    # ( de situatie voor het object die de werknemer heeft geselecteerd )
    # 8  :  BESCHIKBAARHEID IS AL VOL
    # 9  :  GEEN (JUISTE) BESCHIKBAARHEID
    # 10 :  WEL STUDENTEN BUITEN WERKGEBIED
    def CSVtoArray(self):
        csvFile = open(self.csvFileName, 'r+')
        csvRead = csv.reader(csvFile, delimiter=",")
        skipFirst = next(csvRead)
        self.data = list(csvRead)
        self.data = sorted(self.data,key=lambda x:x[7])
        
        
        
    def arrayToFinalData(self):
        finaldata = []
        for i in self.data:
            if i[9] == "1" and (i[8] == "0" or i[10] == "0"):
                for j in finaldata:
                    if i[7] == j[0]:
                        j[1] = j[1] + 1
                        break
                else:
                    finaldata.append([i[7], 1])
        return finaldata
        
    def printAllData(self, data):
        for i in self.data:
                print(f" {i[0]:5}| {i[1]:10}| {i[2]:10}| {i[3]:30}| {i[4]:25}| {i[5]:10}| {i[6]:50}| {i[7]:30}| {i[8]:2}| {i[9]:2}| {i[10]:2}")
    
    def printFinalData(self, data):
        if input("Sort name (1) or total (2) : ") == "1":
            data = sorted(data,key=lambda x:x[0])
        else:
            data = sorted(data,key=lambda x:x[1], reverse=True)      
        
        for i in data:
                print(f"{i[0]:30} | {i[1]:2}")
        
    
    
    
    def makeBarChart(self, data):
        xNum = 1
        x = []       
        stadValue = []
        stadName = []
        for i in data:
            if i[1] >= 5:
                x.append(xNum)
                xNum += 1
                stadValue.append(i[1])
                stadName.append(i[0])
        plt.bar(x, stadValue, tick_label = stadName, width = 0.5, color = ['red', 'blue'])
        plt.xlabel("Stads namen")
        plt.ylabel("Aantal meldingen geen beschikbaarheid")
        plt.title("Data")
        plt.show()
        
        
main = Main()
main.run()


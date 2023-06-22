from cProfile import label # Importing Libraries
from tkinter import *
from tkinter import simpledialog
import pyodbc
import matplotlib.pyplot as plt
import pandas as pd
import datetime
from tkinter import messagebox

all_excels = []
NumberOfBatches = int(input("Number of batches to compares?"))

for i in range(NumberOfBatches):

    class MyDialog(simpledialog.Dialog): # Define a class to create dialog box
        def body(self, master):
        
            Label(master, text = "VesselName:").grid(row = 0) # creating heading before text box
            Label(master, text = "").grid(row = 1) # add empty row to create space between two text box
            Label(master, text = "StartDateTime(YYYY-MM-DD HH:MM:SS):").grid(row = 2) # creating heading before text box
            Label(master, text = "").grid(row = 3) # add empty row to create space between two text box
            Label(master, text = "EndDateTime(YYYY-MM-DD HH:MM:SS):").grid(row = 4) # creating heading before text box
            Label(master, text = "").grid(row = 5) # add empty row to create space between two text box
            Label(master, text = "FileName:").grid(row = 6) # creating heading before text box 
        
            self.e1 = Entry(master)
            self.e2 = Entry(master)
            self.e3 = Entry(master)
            self.e4 = Entry(master)
            self.e5 = Entry(master)
            self.e6 = Entry(master)
            self.e7 = Entry(master)
        
            self.e1.grid(row=0, column=1) # creating text box
            #self.e2.grid(row=1, column=1) {empty space between two text box}
            self.e3.grid(row=2, column=1) # creating text box
            #self.e4.grid(row=3, column=1) {empty space between two text box}
            self.e5.grid(row=4, column=1) # creating text box
            #self.e6.grid(row=5, column=1) {empty space between two text box}
            self.e7.grid(row=6, column=1) # creating text box
            return self.e1
    
        def apply(self):
            Report_Name = self.e1.get() # taking input from user
            Start_Date = self.e3.get() # taking input from user
            End_Date = self.e5.get() # taking input from user
            File_Name = self.e7.get() # taking input from user
        
            connection = pyodbc.connect(driver = '{SQL Server Native Client 11.0}', host = 'EWS1', database = f'{Report_Name}_ReportDB', Trusted_Connection = 'yes') # connecting with SQL server

            sqlQuery = f"select *  FROM [{Report_Name}_ReportDB].[dbo].[ProcessData] where (Date_Time BETWEEN '{Start_Date}.000' AND '{End_Date}.000') AND (Log_Mi = 1 OR OP_Comment IS NOT NULL) order by Date_Time asc" # giving SQL Query to do task


            df = pd.read_sql_query(sqlQuery, connection) # read SQL query with pandas data frame

            df.to_csv(f'{File_Name}.csv') # Save file as csv format

            all_excels.append(File_Name)
              
    root = Tk() # creating message box  
    root.title('File Downloader')
    root.config(bg='black')
    root.attributes('-fullscreen', True)
    d = MyDialog(root) 

main = Tk()
main.geometry("300x200")

w=Label(main, text='Have a nice day :)', font='50')
w.pack()

messagebox.showinfo("Have a nice dasy :)", "Your file is ready") # pops up after your file is being downloaded



df0 = pd.read_csv(f"CSV/{all_excels[0]}.csv")
ax = df0.plot(x="Log Hour", y="DO %", kind = "line", label = None)


for i in all_excels:

    xl = pd.ExcelFile(f'{i}.xlsx') #select excel file

    df = pd.read_excel(xl, skiprows=13) #read excel file after 13 rows

    remove_cols = [col for col in df.columns if 'Unnamed' in col] #select all blank column

    df.drop(remove_cols, axis='columns', inplace=True) #delete all blank column

    df.drop('Date & Time', axis=1, inplace=True) #delete date time column

    df["Log Hour"] = df["Log Hour"].str.replace(":" , ".") # replace ":" to "." in log hour
    
    df.to_csv(f'CSV/{i}.csv', index=False)#save CSV file

    df = pd.read_csv(f"CSV/{i}.csv")#read CSV file

    df.plot(x="Log Hour", y="DO %", kind = "line", label=f'{i}', ax = ax)#plot Graph
      
  
plt.xlabel("log hour") #Label graph
plt.ylabel("DO%")
plt.title("logH vs DO")
plt.legend()


plt.show() #show Graph



































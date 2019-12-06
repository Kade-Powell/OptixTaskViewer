from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox
import xlrd
import time
import pandas as pd
import webbrowser
import csv


# Button functions
def getFile():
    global window
    global selectedFile
    selectedFile =  filedialog.askopenfilename(filetypes = (("spread sheet","*.csv"),("Any files","*.*")))
    Label(window, text=selectedFile, width=50,).grid(row=5, column=1)

def getExportView():
    global var3
    global emailEntry
    userEmail = emailEntry.get()
    listType = var3.get()
    
    try:
        print("main try block")
        driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
        ##start##
        driver.implicitly_wait(30)
        driver.get("https://optix.cox.com/prime")
        elem = driver.find_element_by_name("loginfmt")
        elem.clear()
        elem.send_keys(userEmail)
        elem.send_keys(Keys.RETURN)
        select = Select(driver.find_element_by_name('view'))
        print(listType)
        if listType == "All Assigned Projects": 
            select.select_by_index(5)
        else:
            select.select_by_index(1)
        driver.find_element_by_xpath("//*[contains(text(), 'Export View')]").click()
        time.sleep(7)
        driver.quit()
    except Exception as e:
        print(e)
        pass


def displayTasks():
    global selectedFile
    global var1
    global var2
    global var3

    taskType = var1.get()
    market = var2.get()
    listType = var3.get()
    print (market)
    print(taskType)
   
    df1 = pd.read_csv(selectedFile)    
    df1['ProjID/CirID'] = df1['ProjID/CirID'].apply(lambda x: f"https://optix.cox.com/prime/accounts/proj/view.asp?projid={x}")
        
    try:
        if listType == "All Assigned Projects":
            rawData=df1[["Assigned Tech","ProjID/CirID", "Company Name","Task Type", "District","Scheduled Date","Duration in Hours"]]
        else:
            rawData=df1[["ProjID/CirID", "Company Name","Task Type", "District","Scheduled Date","Duration"]]

        if market == "Any" and taskType == "Any":
            print("if with any any")
            output_data = rawData
        elif market == "Any":
            output_data = rawData[(rawData['Task Type'] == taskType)]
        elif taskType == "Any":
            output_data = rawData[(rawData['District'] == market)] 
        else:
            print("hit the else")
            output_data = rawData[(rawData['District'] == market)&(rawData['Task Type'] == taskType)]
        
        table = output_data.to_html(index=False,table_id='myTable', render_links=True)
        h1tag=f"<h1>{listType}: {taskType} in {market}</h1>"

        #create header template for htmll file
        html = '''
        <!DOCTYPE html>
        <html lang="en">

        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <meta http-equiv="X-UA-Compatible" content="ie=edge">
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
            <link href="https://cdn.datatables.net/1.10.10/css/jquery.dataTables.min.css" rel="stylesheet">
            <link rel="stylesheet" href="style.css"><!-- Custom Styling / Override Bootstrap -->

            <title>Optix Task Viewer</title>
        </head>
        
        <body>
        <div class="container-fluid">
        '''
        html += h1tag
        html += table
        #create footer
        html+='''

        <footer>
        <!-- Bootstrap core JavaScript -->
        <!-- Placed at the end of the document so the pages load faster -->
        <!-- Optional JavaScript -->
        <!-- jQuery first, then Popper.js, then Bootstrap JS -->

        <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js" integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1" crossorigin="anonymous"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js" integrity="sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM" crossorigin="anonymous"></script>
        <script src="https://cdn.datatables.net/1.10.10/js/jquery.dataTables.min.js"></script>

        <!-- User JavaScript Section -->
        <!-- Placed at the end of the document so the pages load faster -->
        <script>
        $(document).ready( function() {
            // Turn existing table addressed by its id `myTable` into datatable
            $("#myTable").DataTable();
        });
        
        </script>
        <!-- User JavaScript Section END -->
        <!-- Placed at the end of the document so the pages load faster -->
        </footer>
        </div>
        </body>
        '''
    except Exception as e:
        html = str(e)
        print(e)


    displayFile = open("display.html", "w")
    displayFile.write(html)
    displayFile.close()
    webbrowser.open_new_tab('display.html')
    

def main():
    global output
    global window
    global emailEntry
    global var1
    global var2
    global var3

    # Main Window
    window = Tk()
    #"window width x window height + position right + position down(from top left corner)"

    window.title('Optix Unassigned Task Viewer')

    # get user info
    Label(window, text="Enter Your CORP Email", font="none 11 bold").grid(row=1, column=1,padx = 1, pady = 3)
    emailEntry = StringVar(window)
    email = Entry(window, textvariable=emailEntry).grid(row=1, column=2,padx = 1, pady = 3)

    #select the unassigned or assigned 
    var3 = StringVar(window)
    Label(window, text="Select Unassigned or Assigned Projects:", font="none 11 bold").grid(row=2, column=1,padx = 1, pady = 3)
    listType = OptionMenu(window, var3, "All Assigned Projects","All Assigned Projects","All Unassigned Projects").grid(row=2 , column=2,padx = 1, pady = 3)
    
    # get csv from optix
    Label(window, text="Get the export view from optix:", font="none 11 bold").grid(row=3, column=1,padx = 1, pady = 3)
    submitButton = Button(window, text='Go', width=15, command=getExportView).grid(row=3, column=2,padx = 1, pady = 3)

    #get file
    Label(window, text="Select the export view from your downloads:", font="none 11 bold").grid(row=4, column=1,padx = 1, pady = 3)
    fileButton = Button(window, text='Select a File', width=15, command=getFile).grid(row=4, column=2,padx = 1, pady = 3)

    #select the task
    var1 = StringVar(window)
    Label(window, text="Select the Tasks to View:", font="none 11 bold").grid(row=6, column=1,padx = 1, pady = 3)
    taskType = OptionMenu(window, var1, "Any","Any","Equipment Staging","Start EWP", "Network Provisioning","Deactivation & Test").grid(row=6, column=2,padx = 1, pady = 3)
    
    #select the market
    var2 = StringVar(window)
    Label(window, text="Select the Market View:", font="none 11 bold").grid(row=7, column=1,padx = 1, pady = 3)
    market = OptionMenu(window, var2, "Any","Any","New Orleans","Lafayette","Baton Rouge","PNS","FWB", "Georgia","Gainesville","Ocala").grid(row=7, column=2,padx = 1, pady = 3)

    #display text
    Label(window, text="Submit and View Tasks:", font="none 11 bold").grid(row=8, column=1,padx = 1, pady = 3)
    fileButton = Button(window, text='Submit', width=15, command=displayTasks).grid(row=8, column=2,padx = 1, pady = 3)

    # run the main loop
    window.mainloop()

if __name__=="__main__":
    main()
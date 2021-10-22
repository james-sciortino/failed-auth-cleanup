import pandas as pd
import os
from os import listdir
from os.path import isfile, join
import sys
import threading
import itertools
import time

class Spinner:

    def __init__(self, message, delay=0.1):
        self.spinner = itertools.cycle(['-', '/', '|', '\\'])
        self.delay = delay
        self.busy = False
        self.spinner_visible = False
        sys.stdout.write(message)

    def write_next(self):
        with self._screen_lock:
            if not self.spinner_visible:
                sys.stdout.write(next(self.spinner))
                self.spinner_visible = True
                sys.stdout.flush()

    def remove_spinner(self, cleanup=False):
        with self._screen_lock:
            if self.spinner_visible:
                sys.stdout.write('\b')
                self.spinner_visible = False
                if cleanup:
                    sys.stdout.write(' ')       # overwrite spinner with blank
                    sys.stdout.write('\r')      # move to next line
                sys.stdout.flush()

    def spinner_task(self):
        while self.busy:
            self.write_next()
            time.sleep(self.delay)
            self.remove_spinner()

    def __enter__(self):
        if sys.stdout.isatty():
            self._screen_lock = threading.Lock()
            self.busy = True
            self.thread = threading.Thread(target=self.spinner_task)
            self.thread.start()

    def __exit__(self, exception, value, tb):
        if sys.stdout.isatty():
            self.busy = False
            self.remove_spinner(cleanup=True)
        else:
            sys.stdout.write('\r')

def create_menu(menu_items): # Create menu options. Result is a dictionary of each site, and a dictionay of the number of sites
    x = 1
    y = 0
    numbers = []
    new_menu = []
    for item in menu_items:
        y = y + x
        numbers.append(str(y))
        new_menu.append(str(y) + ": " + item)
    menu_dict = dict(zip(numbers, menu_items))
    return menu_dict, new_menu
####
def file_menu():       # Present your menu options using the dictionary made from create_menu()
        print(30 * "-" , "File Selection" , 30 * "-") 
        for file in file_string:
            print(file)
        print("0: Exit")
        print(67 * "-")
####
def file_selection(): # Define the logic behind the menu choices. Here we match the numeric value of your menu choice to the dictionary key of each site we created earlier. 
    max = len(file_string)
    min = 1
    add = range(min, max+1)
    total = []
    for i in add:
        total.append(i)
    loop=True
    while loop:  ## While loop which will keep going until loop = False
        file_menu()    ## Displays menu        
        selection = input("Which site do you want to manage? Enter your choice ") #1        
        if selection in str(total):     
            choice = file_dict[selection] #United States
            print(choice + " has been selected")
            return choice
            ## You can add your code or functions here
        elif selection=="0":
            print("Exit has been selected")
            loop=False # This will make the while loop to end as not value of loop is set to False
        else:
            # Any integer inputs other than values 1-5 we print an error message
            input("Wrong option selection. Enter any key to try again..") 
####
def report_filter():
    df = pd.read_excel(xlsx)
    df = df[df["'ENDPOINTMATCHEDPROFILE'"] !=  "''"]
    df = df[df["'ENDPOINTMATCHEDPROFILE'"] != "'Axis-Device'"]
    df = df.drop_duplicates(subset=["'USER_NAME'"])
    buildings = []
    extract = df["'NETWORK_DEVICE_NAME'"].tolist()
    for x in extract:
        buildings.append(x.split('-')[0])
    df["'BUILDINGS'"] = buildings
    return df

thisdir = os.getcwd()
files = [f for f in listdir(thisdir) if isfile(join(thisdir, f))]
menu_create = create_menu(files)
file_dict = menu_create[0]
file_string = menu_create[1]
file_pick = file_selection()
with Spinner("Loading Excel File..."):
    xlsx = pd.ExcelFile(file_pick)
    time.sleep(3)
with Spinner("Cleaning Excel File..."):
    dataframe = report_filter()
    time.sleep(3)
new_report = input("Input the name of the new report:" )
create_report = dataframe.to_excel(new_report + ".xlsx")

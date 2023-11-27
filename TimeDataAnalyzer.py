# File: Main.py
# Author: Steven Heid
# Date: 21 Nov 2023
# Description:  This program processes data from 2 files to compare employee hours. Ouput is stored as .txt
#               File 1 - .xls data that contains Data/Time entries from CrossChex software (physical time tracking).
#               File 2 - .pdf file that contains Date and Hour entries for payroll (manual hour entries).
# Version: 1.1

import fitz
import pandas as pd
import re
from fuzzywuzzy import fuzz
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog


OutputUser = "Justin Rodgers"

def print_processing_bar(num_dots, total_width):
    num_dashes = total_width - num_dots
    bar = "Processing [" + "o" * num_dots + "-" * num_dashes + "]"
    print(bar, end='\r')
total_width =54

print('Select Files.', end='\r')

def select_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return file_path

def select_save_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title=title, filetypes=filetypes)
    return file_path

CrossChexExcel = select_file("Select CrossChex Excel File", [("Excel Files", "*.xls")])
if not CrossChexExcel:
    print("CrossChex Excel file not selected. Exiting.")
    exit()

print_processing_bar(1, total_width)

# Name mapping between files. For names outside of fuzzywuzzy score.
name_mapping = {
    "Ben Zapadenko": "Benjamin Zapadenko",
    "Mathew Jacobs": "Matthew Scott Jacobs",
    "Ben Kilgore": "Benjamin Kilgore",
    "Mattew VanConant": "Matthew VanConant"
}

TimekeeperPDF = select_file("Select Timekeeper PDF File", [("PDF Files", "*.pdf")])
if not TimekeeperPDF:
    print("Timekeeper PDF file not selected. Exiting.")
    exit()


CrossChexExcel = pd.read_excel(CrossChexExcel) 
CrossChexExcel['Name'] = CrossChexExcel['Name'].replace(name_mapping)
Unique_CrossChex_Names = CrossChexExcel['Name'].unique()
CrossChexExcel['Date/Time'] = pd.to_datetime(CrossChexExcel['Date/Time']).dt.strftime("%m/%d/%Y %H:%M:%S")
start_date = pd.to_datetime(CrossChexExcel['Date/Time']).min().strftime("%m/%d/%Y")
week_list = [((pd.to_datetime(start_date) + timedelta(days=day)).strftime("%m/%d/%Y")).lstrip("0").replace("/0", "/") for day in range(7)]
#print(week_list)
days_to_check = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
threshold = 82
date_format = "%m/%d/%Y"

# Function to extract usable TimeKeeper list from pdf.
def extract_single_PDFlist(timekeeper_file):
    pdf_document = fitz.open(timekeeper_file)
    timekeeper_list = [] 

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        page_text = page.get_text().split('\n')
        timekeeper_list.extend(page_text)

    omitted_words = ['No', 'Craft', 'Perdiem', 'QCT']   
    
    filtered_list = [item for item in timekeeper_list if len(item) > 1 and all(omitted_word not in item for omitted_word in omitted_words)]
    return filtered_list

TimeKeeper_List = extract_single_PDFlist(TimekeeperPDF)

"""
### Use this section to export the TimeKeeper_List to a readable .txt file. Used for troubleshooting.
timekeeper_list_output_path = select_save_file("Save TimeKeeper List File", [("Text Files", "*.txt")])
if not timekeeper_list_output_path:
    print("TimeKeeper List file not specified. Exiting.")
    exit()

with open(timekeeper_list_output_path, 'w') as timekeeper_list_file:
    for item in TimeKeeper_List:
        timekeeper_list_file.write(f"{item}\n")
"""


class User:
    def __init__(self, username):
        self.username = username
        self.CrossChex_entries = {}
        self.TimeKeeper_entries = {}
        self.Cleaned_CrossChex = {}
        self.Cleaned_TimeKeeper = {}
        self.TimeKeeper_THours = ''
        self.CrossChex_THours = ''

    def add_CrossChex_entry(self, timestamp):
        
        date_str, time_str = timestamp.split(' ')
        date_dt = datetime.strptime(date_str, "%m/%d/%Y").date()
        time_dt = datetime.strptime(time_str, "%H:%M:%S").time()
        
####    Adjust time cutoff here: 
        if time_dt < datetime.strptime("03:30:00", "%H:%M:%S").time():
            date_dt -= timedelta(days=1)

        adjusted_date_str = date_dt.strftime("%m/%d/%Y")

        if adjusted_date_str not in self.CrossChex_entries:
            self.CrossChex_entries[adjusted_date_str] = []

        self.CrossChex_entries[adjusted_date_str].append(timestamp)
        
    def add_TimeKeeper_entry(self, timestamp, paytype):
        date = timestamp.strftime("%m/%d/%Y")
        paytype = paytype.rstrip(' ')
        formatted_entry = (timestamp.strftime('%H:%M'), paytype)
        if date not in self.TimeKeeper_entries:
            self.TimeKeeper_entries[date] = []     
        self.TimeKeeper_entries[date].append(formatted_entry)

    def format_date(self, date):
        return date.strftime("%m/%d/%Y").lstrip("0").replace("/0", "/")

    def get_Cleaned_CrossChex(self):
        result_dict = {}
        
        for date, times in self.CrossChex_entries.items():
            if len(times) == 1:
                time = datetime.strptime(times[0], "%m/%d/%Y %H:%M:%S")
                result_dict[self.format_date(time)] = f"1 Time @ {time.strftime('%H:%M:%S')}"
            else:
                time_objects = [datetime.strptime(time, "%m/%d/%Y %H:%M:%S") for time in times]
                time_difference = max(time_objects) - min(time_objects)
                
### Adjust break parameters here:
                if time_difference > timedelta(hours=6, minutes=45):
                    time_difference -= timedelta(minutes=30)

                hours, remainder = divmod(time_difference.total_seconds(), 3600)
                minutes, seconds = divmod(remainder, 60)
                time_diff_str = f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"

                formatted_date = self.format_date(min(time_objects))
                result_dict[formatted_date] = time_diff_str

        self.Cleaned_CrossChex = result_dict
        return self.Cleaned_CrossChex

    def get_Cleaned_TimeKeeper(self):
        result_dict = {}

        for date, entries in self.TimeKeeper_entries.items():
            total_time = timedelta(hours=0, minutes=0)

            for entry in entries:
                time, category = entry
                if category in ('ST', 'OT', 'DT'):
                    time_parts = time.split(':')
                    hours, minutes = int(time_parts[0]), int(time_parts[1])
                    total_time += timedelta(hours=hours, minutes=minutes)

            total_time_str = f"{total_time.seconds // 3600:02d}:{(total_time.seconds // 60) % 60:02d}"
            formatted_date = self.format_date(datetime.strptime(date, "%m/%d/%Y"))
            result_dict[formatted_date] = total_time_str

        self.Cleaned_TimeKeeper = result_dict
        return self.Cleaned_TimeKeeper
        
    def get_TimeKeeper_Total(self):

        total_minutes = sum(int(hours) * 60 + int(minutes) for _, time in self.Cleaned_TimeKeeper.items() for hours, minutes in (time.split(':'),))
        hours, minutes = divmod(total_minutes, 60)
        self.TimeKeeper_THours = f"{hours:02d}:{minutes:02d}:00"
       
        return self.TimeKeeper_THours
    
    def get_CrossChex_Total(self):
        total_seconds = 0
        
        for date, time_str in self.Cleaned_CrossChex.items():
            if date in week_list:           ####################
            
                time_pattern = r'^\d{2}:\d{2}:\d{2}$'

                if re.match(time_pattern, time_str):
                    time_parts = time_str.split(':')
                    hours, minutes, seconds = map(int, time_parts)
                    total_seconds += hours * 3600 + minutes * 60 + seconds
                else:
                    pass  
        
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        self.CrossChex_THours = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return self.CrossChex_THours
    
    def __str__(self):
        return f'Username: {self.username}'

print_processing_bar(2, total_width)

##############
user_data = {}

### Initialize users using CrossChex data.
for username in Unique_CrossChex_Names:
    user_data[username] = User(username)
    user_entries = CrossChexExcel[CrossChexExcel['Name'] == username]
    
    for _, row in user_entries.iterrows():
        timestamp = row['Date/Time'] 
        user_data[username].add_CrossChex_entry(timestamp)

print_processing_bar(3, total_width)
dots = 4

### Populate users with TimeKeeper data.
for username, user_instance in user_data.items():
    idx_list = []   # Indexes that match user name.
    purge_range_list = []    
    
    for index, item in enumerate(TimeKeeper_List):
        similarity_score = fuzz.token_set_ratio(username, item)     
                 
        if similarity_score >= threshold and index < len(TimeKeeper_List) -1 and TimeKeeper_List[index + 1] != 'SUPERVISOR':
            idx_list.append(index)
            
                        ############################################################################################
            if username == OutputUser:            
                print(f"{username} @ {index}: data:{TimeKeeper_List[index]}")
    
    for idx in idx_list:
        for index, entri in enumerate(TimeKeeper_List[idx:], start=idx):
            
            if entri in week_list and TimeKeeper_List[index - 1] in days_to_check:  # Date found
                datef = datetime.strptime(entri, date_format)  # Store date
                hours = float(TimeKeeper_List[index + 1])  # Store next entry as hours
                fhours, fminutes = divmod(int(hours * 60), 60)
                time_delta = timedelta(hours=fhours, minutes=fminutes)  # Formated time
                new_datetime = datef + time_delta
                pay_type = TimeKeeper_List[index + 2]  # Store pay type (ST, OT, DT, PERH)
                user_data[username].add_TimeKeeper_entry(new_datetime, pay_type)

            # Purge data from list after processing (for optimization).
            if entri == 'Overtime:':
                start_idx = idx
                stop_idx = index
                purge_range = [idx, index]
                purge_range_list.append(purge_range)
                
                
                if username == OutputUser:
                    print(f"\n{username} start index: {start_idx} data: {TimeKeeper_List[start_idx]}")
                    print(f"\n{username} stop index: {stop_idx} data: {TimeKeeper_List[stop_idx]}")
                    print(TimeKeeper_List[start_idx:stop_idx])
                #TimeKeeper_List = TimeKeeper_List[:start_idx] + TimeKeeper_List[stop_idx:]
                break_outer_loop = True
                break
            
            elif entri == 'WEEK OF':
                break_outer_loop = True
                break
            
        if break_outer_loop:
            break_outer_loop = False
            break
    if username == OutputUser:
                    print(f"\n purge_range_list {purge_range_list}")
        
    for start, stop in purge_range_list:
        TimeKeeper_List = TimeKeeper_List[:start] + TimeKeeper_List[stop:]
        
       
    print_processing_bar(dots, total_width)  
    dots += 1


if OutputUser in user_data:
    # Access the User instance for the desired username
    user_inst = user_data[OutputUser]
    print(f"\n_________________________{user_inst.username}\n")

    cleaned_CrossChex_data = user_inst.Cleaned_CrossChex
    cleaned_TimeKeeper_data = user_inst.Cleaned_TimeKeeper
 
    for day in week_list:
        cnt = False
        date_obj = datetime.strptime(day, "%m/%d/%Y")
        DoW = date_obj.strftime("%a")
        for date, entries in user_inst.get_Cleaned_CrossChex().items():
            if date in day:
                print(f"{DoW} {date}\t{entries}")
                cnt = True
        if cnt == False:
            print(f"{DoW} {day}\t---")
    print(f"CrossChex Hours: {user_inst.get_CrossChex_Total()}") 
    
    print(f"get_Cleaned_CrossChex: {user_inst.get_Cleaned_CrossChex()} \n")
    
    for day in week_list:
        date_obj = datetime.strptime(day, "%m/%d/%Y")
        DoW = date_obj.strftime("%a")
        cnt2 = False
        for date, entries in user_inst.get_Cleaned_TimeKeeper().items():
            if date in day:
                print(f"{DoW} {date}\t{entries}")
                cnt2 = True
        if cnt2 == False:
            print(f"{DoW} {day}\t---")
    print(f"TimeKeeper Hours: {user_inst.get_TimeKeeper_Total()}")
    
    print(f"get_Cleaned_TimeKeeper:{user_inst.get_Cleaned_TimeKeeper()}")
    
output_file_path = select_save_file("Save Output File", [("Text Files", "*.txt")])
if not output_file_path:
    print("Output file not specified. Exiting.")
    exit()

with open(output_file_path, 'w') as output_file:
    for username, user_instance in user_data.items():
        tk_string = ""
        cc_string = ""
#CrossChex       
        output_file.write(f"\n_________________________{user_instance.username}\n")
        for day in week_list:
            counter = False
            date_obj = datetime.strptime(day, "%m/%d/%Y")
            DoW = date_obj.strftime("%a")
            for date, entries in user_instance.get_Cleaned_CrossChex().items():
                if date in day:
                    cc_string += f"{DoW} {date}\t{entries}\n"
                    counter = True
            if counter == False:
                cc_string += f"{DoW} {day}\t---\n"
        
        output_file.write(f"\nCrossChex Hours: {user_instance.get_CrossChex_Total()}\n")
        output_file.write(cc_string)
#TimeKeeper
        for day in week_list:
            date_obj = datetime.strptime(day, "%m/%d/%Y")
            DoW = date_obj.strftime("%a")
            counter2 = False
            for date, entries in user_instance.get_Cleaned_TimeKeeper().items():
                if date in day:
                    tk_string += f"{DoW} {date}\t{entries}\n"
                    counter2 = True
            if counter2 == False:
                tk_string += f"{DoW} {day}\t---\n"

        output_file.write(f"\nTimeKeeper Hours: {user_instance.get_TimeKeeper_Total()}\n")
        output_file.write(tk_string)


print(f'\nDone.')

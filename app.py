import streamlit as st
import openpyxl 
import pandas as pd 
from itertools import islice
import datetime
import os 
import glob

def start_row(worksheet):
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Extended Role CRT":
                start_row = cell.row
                
    return start_row

def crt_names(worksheet):
    crt_names = {}
    for row in worksheet.iter_rows(min_row=(start_row(worksheet))+1, min_col=0, max_col=1 ):
        for cell in row:
            if isinstance(cell.value, str):
                crt_names[cell.value] = cell 
    return crt_names

def time_format(time):
    hours = int(time)
    minutes = int((time - hours) * 100)
    
    time_obj = datetime.time(hour=hours, minute=minutes)
    
    return time_obj

def time_in_hours(timing):
    list_hour = []
    for i in range(0,len(timing),2):
        time1 = timing[i]
        time2 = timing[i+1]
        
        if time2 < time1:
            dt1 = datetime.datetime.combine(datetime.date.today(), time1)
            dt2 = datetime.datetime.combine(datetime.date.today(), time2) + datetime.timedelta(days=1) 
        else:
            dt1 = datetime.datetime.combine(datetime.date.today(), time1)
            dt2 = datetime.datetime.combine(datetime.date.today(), time2)
            
        delta = (dt2 - dt1).total_seconds()
        hours = round(delta/3600, 2)
        
        if hours >= 7: 
            list_hour.append(hours - 1)
        else: 
            list_hour.append(hours)
    
    return list_hour

def crt_hours_dict(worksheet):
    study_hour_dict = {}
    crt = crt_names(worksheet)
    for key in crt:
        first_row = crt[key].row
        second_row = crt[key].offset(row=1, column=0).row
        
        sliced_worksheet = list(worksheet.iter_rows(min_row= first_row, max_row=second_row, min_col= 2, values_only= True))
        
        timing = [time_format(x) for x in sliced_worksheet[0] if x is not None and x != "-" and not isinstance(x, str)]
        timing = time_in_hours(timing)
        studies= [x for x in sliced_worksheet[1] if x is not None]
        
        for i, study in enumerate(studies):
            list_of_studies = study.split("/")
            list_of_studies = [substring.strip() for substring in list_of_studies]
            
            for s in list_of_studies: 
                if s in study_hour_dict:
                    study_hour_dict[s] += timing[i]/(len(list_of_studies))
                else:
                    study_hour_dict[s] = timing[i]/(len(list_of_studies))
    
    return study_hour_dict
  
def main(files):
  excel_files = files
  st.write()
  if len(excel_files) > 0:
    for file_ in excel_files:
      workbook = openpyxl.load_workbook(file_)
      worksheet = workbook["Sheet1"]
      results = crt_hours_dict(worksheet)
      st.write("-"*60)
      st.write(f'''Now tabulating for this excel roster: {file_.name}''',"\n")
      for key, value in results.items():
        st.write(key, ":", value)
      st.write()
    
    st.write()
    st.write("="*60)
    st.write(f''' We are done with a total of {len(excel_files)} excel files \U0001F601 ''')
  else:
    st.write("\nThere are no excel files in the folder path you just entered. Please check again!!!")
  return 

st.set_page_config(page_title = "CRT Hours Tabulator")
st.title("\U0001F4C8 CRT Hours and Study Allocated \U0001F4C9")
st.subheader("Input Excel Files")

st.write("\u2757 Please remember to remove date columns that are out of the month of interest first before running \u2757")

uploaded_files = st.file_uploader("Choose the Excel Files to Upload (.xlsx)", 
                                  type ="xlsx", accept_multiple_files=True)

if uploaded_files:
    main(uploaded_files)
    
    

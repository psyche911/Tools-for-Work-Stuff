# -*- coding: utf-8 -*-
"""
Created on Wed Sep 1, 2021

@author: Denise Mao
"""

import pandas as pd
import regex as re
import time
import datetime
import os
import shutil
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as msg
from tkinter.filedialog import askdirectory


def filterNumber(num):
    if(len(num) >= 5):
        return True
    else:
        return False

def lst_to_str(str):
    x = re.findall('[0-9]+', str)
    finalx = list(filter(filterNumber, x))
    if len(finalx) == 0:
        return ""
    else:
        return list(filter(filterNumber, x))[0].lstrip("0")
        

class extract_vpn:
   
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name
        self.path = StringVar()

    def selectFile(self):
        self.file_name = filedialog.askopenfilename(initialdir='./',
                                                    title = 'Select an Excel file',
                                                    filetypes = [("xlsx", "*.xlsx"), ('CSV', '*.csv'), 
                                                                 ('PDF', '*.pdf'), ('TEXT', '*.txt')])

    def error(self):    # pop up window
        msg.showinfo(title="Warning", message="No File Selected")
        
    def uptime(self):   # update time
        TimeLabel["text"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S:') + "%d" % (datetime.datetime.now().microsecond // 100000)
        self.init_window_name.after(100, self.uptime)        
    
    def set_init_window(self):
        # set Window
        self.init_window_name.title("VPN Extraction")               # Window Title Bar
        self.init_window_name.geometry('300x300+600+300')           # Window size: 300x300，left margin: +600，top margin: +300
        self.init_window_name.resizable(width=FALSE,height=FALSE)   # fixed window size
        Label(self.init_window_name,text="Extracting VPN from Free Text",bg="SkyBlue",fg="Gray").place(x=70,y=10)  # label component
        Button(self.init_window_name,text="Select File",command=self.selectFile,bg="SkyBlue").place(width=200,height=50,x=50,y=45)   # button component to trigger functions
        Button(self.init_window_name,text="Process",bg="SkyBlue",command=self.processFile).place(width=200,height=50,x=50,y=115)
        Button(self.init_window_name,text="Exit",bg="SkyBlue",command=self.init_window_name.destroy).place(width=200,height=50, x=50, y=185)
        self.init_window_name["bg"] = "SkyBlue"     # window background colour
        self.init_window_name.attributes("-alpha")  #  opacity/transparency
        global TimeLabel
        TimeLabel = Label(text="%s%d" % (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S:'),datetime.datetime.now().microsecond // 100000),bg="SkyBlue")
        TimeLabel.place(x=80, y=260)
        self.init_window_name.after(100, self.uptime)
        
    def processFile(self):
        data = pd.ExcelFile(self.file_name)
        df1 = pd.read_excel(data, 'Sheet1', header=1)
        df2 = pd.read_excel(data, 'Sheet2', header=1)

        if(len(df1) == 0 and len(df2) == 0):      
            msg.showinfo('No Rows Selected', 'Excel has no rows')
        else:
            # Remove space in string
            df1['SHORT TEXT'] = df1['SHORT TEXT'].str.replace(" ","")
            df2['SHORT TEXT'] = df2['SHORT TEXT'].str.replace(" ","")
            
            df1['SHORT TEXT'] = df1['SHORT TEXT'].fillna("FreeText")
            df2['SHORT TEXT'] = df2['SHORT TEXT'].fillna("FreeText")
            
            # Perform extraction
            df1['extrVPN'] = df1['SHORT TEXT'].apply(lst_to_str)
            df2['extrVPN'] = df2['SHORT TEXT'].apply(lst_to_str)

            # Create a Pandas Excel writer using XlsxWriter as the engine
            writer = pd.ExcelWriter('Wurth & Galvins Plumbing_extr.xlsx', engine='xlsxwriter')

            # Write each DataFrame to a specific sheet
            df1.to_excel(writer, sheet_name='Sheet1', index = False)
            df2.to_excel(writer, sheet_name='Sheet2', index = False)

            # Close the Pandas Excel writer and output the Excel file
            writer.save()
            
            msg.showinfo(title="Well", message="Done")
            
    def get_current_time(self):
        # Get current time
        current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        return current_time 
    
                
if __name__ == '__main__':
    init_window = Tk()    # instantiate a parent window
    VPN_PORTAL = extract_vpn(init_window)
    VPN_PORTAL.set_init_window()
    init_window.mainloop()

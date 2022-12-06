# -*- coding: utf-8 -*-
#pip install pandas

import pandas as pd
import tkinter as tk
from tkinter.filedialog import asksaveasfilename, askopenfilename
import tkinter.ttk as ttk
import os

def output_xlsx(res):
    filename=asksaveasfilename(defaultextension="*.xlsx",filetypes=(("Excel檔", "*.xlsx"),("All files", "*.*")))
    if filename=="":
        return
    with open(filename,'w') as output:
        res.to_excel(filename,index=False,freeze_panes=(1,0))
    os.startfile(filename)

def check_column_name(): #核對二個檔案的columns名稱是否完全依樣
    new_file=pd.read_excel(new_path_entry.get(),dtype=str)
    old_file=pd.read_excel(old_path_entry.get(),dtype=str)
    a_column=list(new_file.columns) #取出二個檔案的columns名稱
    b_column=list(old_file.columns)
    if (a_column==b_column)==False:
        tk.messagebox.showinfo('標題列不符','二個檔案的欄位名稱不相同，請重新選擇檔案')
        new_path_entry.focus()
    else:
        do_leave_not_duplicate(new_file,old_file)
        
def read_excel_path(input_entry):
    path=askopenfilename(filetypes=(("Excel檔", "*.xlsx"),("All files", "*.*")))
    input_entry.delete(0,'end')
    input_entry.insert(0,path)
    input_entry.focus()
    
def read_columns_name_tolist(): #把舊檔案的column名稱讀出來，做成list給下拉選單用
    temp=pd.read_excel(old_path_entry.get())
    choose_column_name['value']=(temp.columns.tolist())
    choose_column_name.current(0) #直接選擇第一個，作為檔案成功讀取的象徵
    
def do_leave_not_duplicate(new_file,old_file):
    old_file.drop_duplicates(subset=[choose_column_name.get()],keep='first') #二個檔案都先分別排除重複
    new_file.drop_duplicates(subset=[choose_column_name.get()],keep='first')
    append_file=old_file.append(new_file) #新舊檔案合併
    append_file=append_file.drop_duplicates(subset=[choose_column_name.get()],keep=False) #以下拉選單選擇欄位，排除重複，留下完全不重複檔案
    output_xlsx(append_file) #重複使用全域輸出excel功能
    
leave_not_duplicate_toplevel=tk.Tk()
leave_not_duplicate_toplevel.option_add('*Font', '微軟正黑體 20') #更改字體
leave_not_duplicate_toplevel.title('移除重複，留下完全不重複')
tk.Label(leave_not_duplicate_toplevel,text='注意：二個檔案的欄位必須完全相同').grid(column=0,row=0,columnspan=2)
old_path_l=tk.Label(leave_not_duplicate_toplevel,text='舊檔案路徑：',bg='indian red')
old_path_l.grid(column=0,row=1)
old_path_entry=tk.Entry(leave_not_duplicate_toplevel)
old_path_entry.grid(column=1,row=1)
old_path_btn=tk.Button(leave_not_duplicate_toplevel,text='選擇檔案',command=lambda:[read_excel_path(old_path_entry),read_columns_name_tolist()])
old_path_btn.grid(column=2,row=1)
new_path_l=tk.Label(leave_not_duplicate_toplevel,text='新檔案路徑：',bg='cyan2')
new_path_l.grid(column=0,row=2)
new_path_entry=tk.Entry(leave_not_duplicate_toplevel)
new_path_entry.grid(column=1,row=2)
new_path_btn=tk.Button(leave_not_duplicate_toplevel,text='選擇檔案',command=lambda:read_excel_path(new_path_entry))
new_path_btn.grid(column=2,row=2)
choose_column_name_l=tk.Label(leave_not_duplicate_toplevel,text='請選擇篩選欄名：')
choose_column_name_l.grid(column=0,row=4)
choose_column_name_str=tk.StringVar()
choose_column_name=ttk.Combobox(leave_not_duplicate_toplevel,textvariable=choose_column_name_str)
choose_column_name.grid(column=1,row=4)
do_leave_not_duplicate_btn=tk.Button(leave_not_duplicate_toplevel,text='執行',command=check_column_name)
do_leave_not_duplicate_btn.grid(column=2,row=4)
leave_not_duplicate_toplevel.mainloop()

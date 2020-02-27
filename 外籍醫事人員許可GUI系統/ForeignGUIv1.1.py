# -*- coding:utf8 -*-
#Foreign Premit System

#import library
import os
import datetime,time
import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog
import pandas as pd

#get format def
def get_format_name():
    global get_fname,permit_names
    get_fname = fnames.get()

    #creat permit menu
    permit_path = fpath + "\{}".format(get_fname)
    permit_dir = os.listdir(permit_path)
    permit_names.set("選擇範本")
    # permit_button_frame = tk.Frame(window)
    # permit_button_frame.pack(anchor="nw",side=tk.TOP)
    permit_menu = tk.OptionMenu(window,permit_names,*permit_dir)
    permit_menu.config(width=14,height=1)
    permit_menu.place(x=0,y=250)
    permit_button = tk.Button(window,text="確認",command=get_permit_name,
                                  width=10,height=1)
    permit_button.place(x=142.5,y=252)


#get permit def
def get_permit_name():
    global get_fname,get_pname,permit_names,pfile_path
    get_pname = permit_names.get()
    pfile_path = fpath + "\{}\{}".format(get_fname,get_pname)
    load_success_frame = tk.Frame(window)
    load_success_frame.pack(anchor="nw",side=tk.TOP)
    load_success_label = tk.Label(load_success_frame,text="選取檔案位置{}".format(pfile_path))
    return load_success_label.pack()

#write permit def
def permit_write():
    global pfile_path,get_pname,body_y
    try:
        #get permit template
        open_file = open(pfile_path,encoding="utf8")
        permit_data = open_file.read()
        open_file.close()
        #get permit Entry
        health_name = health_name_entry.get()
        health_date = health_date_entry.get()
        health_word = health_word_entry.get()
        health_number = health_number_entry.get()
        person_name = person_name_entry.get()
        stay_date = stay_date_entry.get()
        work_location = work_location_entry.get()
        worklocation_bossname = worklocation_bossname_entry.get()
        work_permitnumber = work_permitnumber_entry.get()
        work_permitdate = work_permitdate_entry.get()

        #replace entry to template
        permit_data = permit_data.replace("{gov_name}",health_name)\
            .replace("{day}",health_date).replace("{gov_code}",health_word)\
            .replace("{gov_code2}",health_number).replace("{name}",person_name)\
            .replace("{stay_day}",stay_date).replace("{orang_name}",work_location)\
            .replace("{mec_boss}",worklocation_bossname).replace("{work_code}",work_permitnumber)\
            .replace("{work_day}",work_permitdate)

        #write target
        target_path = path+"\Permit\{}{}{}".format(person_name,get_fname,get_pname)
        permit_target = open(target_path,"w",encoding="utf8")
        permit_target.write(permit_data)
        permit_target.close()
        write_success_frame = tk.Frame(window)
        write_success_frame.pack(anchor="nw",side=tk.TOP)
        write_success_label = tk.Label(write_success_frame,
                                   text="{}{}的{}許可公文已經存至\n{}".format(person_name,get_fname,get_pname[:-4:],target_path))
        write_success_label.pack()

        #write access to body
        body_y+=40
        body_label = tk.Label(window,text="{}{}-{}".format(person_name,get_fname,get_pname[:-4:]))
        body_label.place(x=400,y=body_y)

    except:
        tk.messagebox.showinfo(title="範本編碼錯誤",message="請記得將該範本編碼轉為UTF-8喔喔喔")

#Aboutme def
def aboutme():
    about_word="這是一個隨便寫的公文範本圖形介面程式ver1.1\n" \
               "作者:Rbbb\n" \
               "更新項目:\n" \
               "1.年月日預設\n" \
               "2.增加已完成許可(今天日期)\n" \
               "3.用爛方法改善PermitButton位置(Tk.place)"
    tk.messagebox.showinfo(title="關於本程式",message=about_word)

#Choose data def
def choose_data():
    choose_path = filedialog.askopenfilename(title="選取輸入文件",initialdir=path)
    health_date_entry.delete(0,"end")
    stay_date_entry.delete(0,"end")
    work_permitdate_entry.delete(0,"end")
    try:
        open_keydata = pd.read_excel(choose_path,index_col=0)
        keydata = open_keydata.T
        keycolumn = keydata.columns
        keydict = {}
        for col_name in keycolumn:
            for value in keydata[col_name]:
                keydict[col_name] = value
    #put data in entry
        health_name_entry.insert("end", keydict[keycolumn[0]])
        health_date_entry.insert("end", keydict[keycolumn[1]])
        health_word_entry.insert("end", keydict[keycolumn[2]])
        health_number_entry.insert("end", keydict[keycolumn[3]])
        person_name_entry.insert("end", keydict[keycolumn[4]])
        stay_date_entry.insert("end", keydict[keycolumn[5]])
        work_location_entry.insert("end", keydict[keycolumn[6]])
        worklocation_bossname_entry.insert("end", keydict[keycolumn[7]])
        work_permitnumber_entry.insert("end", keydict[keycolumn[8]])
        work_permitdate_entry.insert("end", keydict[keycolumn[9]])
    except:
        tk.messagebox.showinfo(title="檔案選取類型錯誤",message="目前僅支援讀取Excel\n請使用.xlsx檔案")

##################GUI Body#####################
if __name__=='__main__':
#window setting
    window = tk.Tk()
    window.title("外籍醫事人員許可範本套用系統")
    window.geometry("697x500")
    window.config(background="#BBFFEE")
    top_frame = tk.Frame(window)
    top_frame.pack(anchor="ne")
    bottom_frame = tk.Frame(window)
    bottom_frame.pack(side=tk.BOTTOM)

#global setting
# permit_function = False
    get_fname,get_pname = "",""
    permit_names = tk.StringVar()
    pfile_path=""
    body_y=-3 #body write flow

#Menu
    top_menu = tk.Menu(window)
    window_menu = tk.Menu(top_menu,tearoff=0)
    window_menu.add_command(label="關於",command=aboutme)
    window_menu.add_command(label="離開",command=window.quit)
    file_menu = tk.Menu(top_menu,tearoff=0)
    file_menu.add_command(label="選取輸入文件",command=choose_data)
    top_menu.add_cascade(label="選單",menu=window_menu)
    top_menu.add_cascade(label="檔案",menu=file_menu)
    window.config(menu=top_menu)

#Replace Label and Entry
    health_name_frame = tk.Frame(window)
    health_name_frame.pack(anchor="nw",side=tk.TOP)
    health_name_label = tk.Label(health_name_frame,text="衛生局名稱",width=15,height=1).pack(side=tk.LEFT)
    health_name_entry = tk.Entry(health_name_frame,width=25)
    health_name_entry.pack(side=tk.LEFT)

    health_date_frame = tk.Frame(window)
    health_date_frame.pack(anchor="nw",side=tk.TOP)
    health_date_label = tk.Label(health_date_frame,text="衛生局發文日期",width=15,height=1).pack(side=tk.LEFT)
    health_date_entry = tk.Entry(health_date_frame,width=25)
    health_date_entry.insert("insert","年月日")
    health_date_entry.pack(side=tk.LEFT)

    health_word_frame = tk.Frame(window)
    health_word_frame.pack(anchor="nw",side=tk.TOP)
    health_word_label = tk.Label(health_word_frame,text="衛生局發文字號",width=15,height=1).pack(side=tk.LEFT)
    health_word_entry = tk.Entry(health_word_frame,width=25)
    health_word_entry.pack(side=tk.LEFT)

    health_number_frame = tk.Frame(window)
    health_number_frame.pack(anchor="nw",side=tk.TOP)
    health_number_label = tk.Label(health_number_frame,text="衛生局發文函號",width=15,height=1).pack(side=tk.LEFT)
    health_number_entry = tk.Entry(health_number_frame,width=25)
    health_number_entry.pack(side=tk.LEFT)

    person_name_frame = tk.Frame(window)
    person_name_frame.pack(anchor="nw",side=tk.TOP)
    person_name_label = tk.Label(person_name_frame,text="醫事人員名子",width=15,height=1).pack(side=tk.LEFT)
    person_name_entry = tk.Entry(person_name_frame,width=25)
    person_name_entry.pack(side=tk.LEFT)

    stay_date_frame = tk.Frame(window)
    stay_date_frame.pack(anchor="nw",side=tk.TOP)
    stay_date_label = tk.Label(stay_date_frame,text="居留期限",width=15,height=1).pack(side=tk.LEFT)
    stay_date_entry = tk.Entry(stay_date_frame,width=25)
    stay_date_entry.insert("insert","年月日")
    stay_date_entry.pack(side=tk.LEFT)

    work_location_frame = tk.Frame(window)
    work_location_frame.pack(anchor="nw",side=tk.TOP)
    work_location_label = tk.Label(work_location_frame,text="執業登記地點",width=15,height=1).pack(side=tk.LEFT)
    work_location_entry = tk.Entry(work_location_frame,width=25)
    work_location_entry.pack(side=tk.LEFT)

    worklocation_bossname_frame = tk.Frame(window)
    worklocation_bossname_frame.pack(anchor="nw",side=tk.TOP)
    worklocation_bossname_label = tk.Label(worklocation_bossname_frame,text="執登機構負責人",width=15,height=1).pack(side=tk.LEFT)
    worklocation_bossname_entry = tk.Entry(worklocation_bossname_frame,width=25)
    worklocation_bossname_entry.pack(side=tk.LEFT)

    work_permitnumber_frame = tk.Frame(window)
    work_permitnumber_frame.pack(anchor="nw",side=tk.TOP)
    work_permitnumber_label = tk.Label(work_permitnumber_frame,text="勞動部許可函號",width=15,height=1).pack(side=tk.LEFT)
    work_permitnumber_entry = tk.Entry(work_permitnumber_frame,width=25)
    work_permitnumber_entry.pack(side=tk.LEFT)

    work_permitdate_frame = tk.Frame(window)
    work_permitdate_frame.pack(anchor="nw",side=tk.TOP)
    work_permitdate_label = tk.Label(work_permitdate_frame,text="勞動部許可工作期限",width=15,height=1).pack(side=tk.LEFT)
    work_permitdate_entry = tk.Entry(work_permitdate_frame,width=25)
    work_permitdate_entry.insert("insert","年月日")
    work_permitdate_entry.pack(side=tk.LEFT)

# Format button
    format_button_frame = tk.Frame(window)
    format_button_frame.pack(anchor="nw",side=tk.TOP)
    path = os.getcwd()
    fpath = path+"\Format"
    fdir = os.listdir(fpath)
    fnames = tk.StringVar()
    fnames.set("選擇人員")
    format_menu = tk.OptionMenu(format_button_frame,fnames,*fdir)
    format_menu.config(width=14,height=1)
    format_menu.pack(side=tk.LEFT)
    format_choice_button = tk.Button(format_button_frame,text="確認",command=get_format_name,width=10,height=1)
    format_choice_button.pack(side=tk.LEFT)

#enpty one
    enpty_frame = tk.Frame(window)
    enpty_frame.pack(anchor="nw",side=tk.TOP)
    enpty_label = tk.Label(enpty_frame,text="",height=2).pack(side=tk.LEFT)

# Write Permit Button
    permit_write_button = tk.Button(bottom_frame,text="產生許可公文", command=permit_write)
    permit_write_button.pack()

# Date Time
    now_time = datetime.datetime.now().strftime("%Y-%m-%d")
    time_text = tk.Label(window,text="{}.已完成許可:".format(now_time)).place(x=400, y=body_y)


# Finish Permit
    window.mainloop()


import sqlite3
import datetime
import pandas as pd
import os
import time
import threading
import sys
from fpdf import FPDF
import winshell
from win32com.client import Dispatch

# get folder where exe or script is
if getattr(sys, 'frozen', False): 
    script_folder = os.path.dirname(sys.executable)

db_path = os.path.join(script_folder, 'main.db')

# db 
def make_db():
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY, event TEXT, time TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS session (id INTEGER PRIMARY KEY, start TEXT, duration REAL)")
    c.execute("INSERT OR IGNORE INTO session (id, start, duration) VALUES (1, '', 0)")
    conn.commit()
    conn.close()

# add event
def do_event(event, event_time):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("INSERT INTO events (event, time) VALUES (?, ?)", (event, event_time))
    conn.commit()
    conn.close()

# excel
def make_excel():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM events", conn)
    conn.close()
    excel_file = os.path.join(script_folder, 'main.xlsx')
    df.to_excel(excel_file, index=False)
    print(f"Saved to {excel_file}")

# pdf
def make_pdf():
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT * FROM events")
    rows = c.fetchall()
    conn.close()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, "PC Journal", 0, 1, 'C')
    pdf.cell(0, 10, "", 0, 1)
    pdf.cell(20, 10, "ID", 1)
    pdf.cell(40, 10, "Event", 1)
    pdf.cell(60, 10, "Time", 1)
    pdf.ln()
    for row in rows:
        pdf.cell(20, 10, str(row[0]), 1)
        pdf.cell(40, 10, row[1], 1)
        pdf.cell(60, 10, row[2], 1)
        pdf.ln()
    pdf_file = os.path.join(script_folder, 'main.pdf')
    pdf.output(pdf_file)
    print(f"Saved to {pdf_file}")

# log
def show_log():
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT * FROM events ORDER BY id DESC")  # newest first
    rows = c.fetchall()
    conn.close()
    if not rows:
        print("No events found.")
        return
    print("\nID | Event | Time")
    print("-" * 30)
    for row in rows:
        print(f"{row[0]} | {row[1]} | {row[2]}")

# autoboot
def set_boot():
    exe_path = os.path.join(script_folder, "PC_Journal.exe")
    startup_folder = winshell.startup()
    shortcut_name = "PC Journal.lnk"
    shortcut_path = os.path.join(startup_folder, shortcut_name)
    
    print("\nAuto-Boot Options:")
    print("1. Add to auto boot")
    print("2. Remove from auto boot")
    print("3. Back to main menu")
    choice = input("Pick a number (1-3): ")
    
    if choice == "1":
        if not os.path.exists(exe_path):
            print("No PC_Journal.exe found!")
            input("Press any key to continue...")
            return
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = script_folder
        shortcut.save()
        print("Added to auto-boot")
        input("Press any key to continue...")
    elif choice == "2":
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            print("Removed from auto-boot")
        else:
            print("No auto-boot shortcut found")
        input("Press any key to continue...")
    elif choice == "3":
        return
    else:
        print("Wrong number!")
        input("Press any key to continue...")

# timer
def save_session_loop(start_time_str, start_secs):
    while True:
        now_secs = time.time()
        how_long = now_secs - start_secs
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute("UPDATE session SET start = ?, duration = ? WHERE id = 1", (start_time_str, how_long))
        conn.commit()
        conn.close()
        time.sleep(1)

# check last
def get_last_session():
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT start, duration FROM session WHERE id = 1")
    last_session = c.fetchone()
    conn.close()
    if last_session:
        return last_session[0], last_session[1]
    return "", 0

# menu 
def menu():
    make_db()
    last_start, last_duration = get_last_session()
    if last_start != "" and last_duration > 0:
        start_time = datetime.datetime.strptime(last_start, '%Y-%m-%d %H:%M:%S')
        seconds = int(last_duration)
        off_time = start_time + datetime.timedelta(seconds=seconds)
        off_time_str = off_time.strftime('%Y-%m-%d %H:%M:%S')
        do_event("PC_OFF", off_time_str)

    start_secs = time.time()
    now = datetime.datetime.now()
    now_str = now.strftime('%Y-%m-%d %H:%M:%S')
    do_event("PC_ON", now_str)

    t = threading.Thread(target=save_session_loop, args=(now_str, start_secs))
    t.daemon = True
    t.start()

    if getattr(sys, 'frozen', False):
        while True:
            print("\nPC Journal Menu:")
            print("1. Save to Excel")
            print("2. Save to PDF")
            print("3. Show PC On/Off Log")
            print("4. Manage Auto-Boot")
            print("5. Exit")
            choice = input("Pick a number (1-5): ")
            if choice == "1":
                make_excel()
            elif choice == "2":
                make_pdf()
            elif choice == "3":
                show_log()
            elif choice == "4":
                set_boot()
            elif choice == "5":
                print("Goodbye!")
                sys.exit()
            else:
                print("Wrong number! Try again.")
    else:
        last_start, last_duration = get_last_session()  # again for script
        if last_start != "" and last_duration > 0:
            start_time = datetime.datetime.strptime(last_start, '%Y-%m-%d %H:%M:%S')
            seconds = int(last_duration)
            off_time = start_time + datetime.timedelta(seconds=seconds)
            off_time_str = off_time.strftime('%Y-%m-%d %H:%M:%S')
            do_event("PC_OFF", off_time_str)
        start_secs = time.time()
        now = datetime.datetime.now()
        now_str = now.strftime('%Y-%m-%d %H:%M:%S')
        do_event("PC_ON", now_str)
        t = threading.Thread(target=save_session_loop, args=(now_str, start_secs))
        t.daemon = True
        t.start()
        while True:
            print("\nMenu:")
            print("1. Excel")
            print("2. PDF")
            print("3. Log")
            print("4. Auto Boot")
            print("5. Quit")
            choice = input("Number (1-5): ")
            if choice == "1":
                make_excel()
            elif choice == "2":
                make_pdf()
            elif choice == "3":
                show_log()
            elif choice == "4":
                set_boot()
            elif choice == "5":
                print("Bye!")
                sys.exit()
            else:
                print("Wrong number!")

# go
def start():
    menu()

start()
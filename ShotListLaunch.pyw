import ttkbootstrap as ttk
import platform
import os
import re
import glob

import sys
def get_root_dir():
    PCName = platform.node()
    if(PCName == 'GLADOS'):
        return ""
    else:
        return r"C:\Users\danys\OneDrive\Documents\טוטל_פרינט\בתי ספר"

def get_years(root_dir):
    dirs = os.listdir(root_dir)
    years = []

    for dir in dirs:
        if(re.match(r'^\d\d\d\d$',str(dir)) and 
           os.path.isdir(fr'{root_dir}\{dir}')):
            years.append(dir)
    years.sort(reverse=True)
    return years
    
def get_shcools_with_accdb(root_dir,year):
    root = fr'{root_dir}\{year}'
    dirs = os.listdir(root)
    schools = []
    for dir in dirs:
        full_path = fr'{root}\{dir}'
        if( os.path.isdir(full_path) and 
           len(glob.glob(fr'{full_path}\*.accdb')) > 0):
            schools.append(dir)
    schools.sort()
    return schools

def update_school_listbox(root_dir,year):
    school_listbox.delete(*school_listbox.get_children())
    for school in get_shcools_with_accdb(root_dir,year):
        school_listbox.insert("", "end", text=school)
    first_item = school_listbox.get_children()[0]    
    school_listbox.selection_set(first_item)
    school_listbox.focus(first_item)

def launch_access():
    year = year_dropdown.get()
    school = school_listbox.item(school_listbox.focus())["text"]
    root = get_root_dir()
    dir = fr'{root}\{year}\{school}'
    access_file = glob.glob(fr'{dir}\*.accdb')
    os.startfile(access_file[0])

def open_location():
    year = year_dropdown.get()
    school = school_listbox.item(school_listbox.focus())["text"]
    root = get_root_dir()
    dir = fr'{root}\{year}\{school}'
    print(dir)
    os.startfile(dir)

def year_listbox_event(event):
    update_school_listbox(get_root_dir(),year_dropdown.get())

root_dir = get_root_dir()

window = ttk.Window(title='ShotList launcher',themename='cyborg')

years = get_years(root_dir)
year_dropdown = ttk.Combobox(values=years)
year_dropdown.current(0)
year_dropdown.pack(pady=10)
year_dropdown.bind("<<ComboboxSelected>>",year_listbox_event)

#School_list = ttk.Listbox()
school_frame = ttk.Frame()
scrollbar = ttk.Scrollbar(school_frame)
school_listbox = ttk.Treeview(school_frame, yscrollcommand=scrollbar.set, show="tree")
scrollbar.configure(command=school_listbox.yview)
update_school_listbox(root_dir,year_dropdown.get())


scrollbar.pack(side="right", fill="y")
school_listbox.pack(side="left",  expand=True)
school_frame.pack()

button_frame = ttk.Frame()
launch_button = ttk.Button(button_frame,text="Launch",command=launch_access)
launch_button.pack(side='left',padx=5)

folder_button = ttk.Button(button_frame,text="Open location",command=open_location)
folder_button.pack(side='left',padx=5)

button_frame.pack(pady=10)

window.mainloop()
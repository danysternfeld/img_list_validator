# A little gui tool to launch school access files.
# Assumes a dir structure of
#  root/
#   {year}/
#       {school1}
#       {school2}
#      ...
###################################################        
import ttkbootstrap as ttk
import platform
import os
import re
import glob

# get roo dir per PC's I use.
def get_root_dir():
    PCName = platform.node()
    if(PCName == 'glados'):
        return r"D:\users\dany\onedrive\Documents\טוטל_פרינט\בתי ספר"
    else:
        return r"C:\Users\danys\OneDrive\Documents\טוטל_פרינט\בתי ספר"

# List year directories
def get_years(root_dir):
    dirs = os.listdir(root_dir)
    years = []
    for dir in dirs:
        if(re.match(r'^\d\d\d\d$',str(dir)) and 
           os.path.isdir(fr'{root_dir}\{dir}')):
            years.append(dir)
    years.sort(reverse=True)
    return years
    
# list school dirs in year directory    
def get_shcools_with_accdb(root_dir,year):
    root = fr'{root_dir}\{year}'
    dirs = os.listdir(root)
    schools = []
    for dir in dirs:
        full_path = fr'{root}\{dir}'
        # filter for directories that contain access files.
        if( os.path.isdir(full_path) and 
           len(glob.glob(fr'{full_path}\*.accdb')) > 0):
            schools.append(dir)
    schools.sort()
    return schools


def update_school_listbox(root_dir,year):
    school_listbox.delete(*school_listbox.get_children())
    for school in get_shcools_with_accdb(root_dir,year):
        school_listbox.insert("", "end", text=school)
    # select first item as default    
    if school_listbox.get_children():
        first_item = school_listbox.get_children()[0]    
        school_listbox.selection_set(first_item)
        school_listbox.focus(first_item)

# launches the first access file found in a school dir selected in the school listbox
def launch_access():
    year = year_dropdown.get()
    school = school_listbox.item(school_listbox.focus())["text"]
    root = get_root_dir()
    dir = fr'{root}\{year}\{school}'
    access_file = glob.glob(fr'{dir}\*.accdb')
    if access_file:
        os.startfile(access_file[0])

# opens a school dir in explorer
def open_location():
    year = year_dropdown.get()
    school = school_listbox.item(school_listbox.focus())["text"]
    root = get_root_dir()
    dir = fr'{root}\{year}\{school}'
    print(dir)
    os.startfile(dir)
    return(dir)

def validate():
    dir = open_location()
    os.chdir(dir)
    os.startfile(os.path.dirname(os.path.realpath(__file__))+r'\img_list_validator.py')

# event handlers
def year_listbox_event(event):
    update_school_listbox(get_root_dir(),year_dropdown.get())

def launch_access_event(event):
    launch_access()    

root_dir = get_root_dir()

# GUI defs
window = ttk.Window(title='ShotList launcher',themename='cyborg')

years = get_years(root_dir)
year_dropdown = ttk.Combobox(values=years, justify='right')
year_dropdown.current(0)
year_dropdown.pack(pady=10)
year_dropdown.bind("<<ComboboxSelected>>",year_listbox_event)

#School_list = ttk.Listbox()
school_frame = ttk.Frame()
scrollbar = ttk.Scrollbar(school_frame)
school_listbox = ttk.Treeview(school_frame, yscrollcommand=scrollbar.set, show="tree")
scrollbar.configure(command=school_listbox.yview)
school_listbox.bind("<Double-1>", launch_access_event)
school_listbox.column('# 0', anchor='e')
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

validate_button = ttk.Button(window,text="validate",command=validate)
validate_button.pack(padx=5,pady=5)

window.mainloop()
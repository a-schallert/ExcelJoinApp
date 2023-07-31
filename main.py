import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu, Label, Text, Listbox
from tkinter.filedialog import asksaveasfile, asksaveasfilename, askdirectory
import os
import xlsxwriter
from excel_join import ExcelJoin

root = tk.Tk()
root.geometry('1000x950')
root.title('Excel Join')
root.pack_propagate(False)
root.resizable(0, 0)

data_frame = tk.LabelFrame(root, text='DataFrame')
data_frame.place(height=500, width=1000)
 
tv1 = ttk.Treeview(data_frame)
tv1.place(relheight=1, relwidth=1)

treescrollx = tk.Scrollbar(data_frame, orient='horizontal', command=tv1.xview)
treescrolly = tk.Scrollbar(data_frame, orient='vertical', command=tv1.yview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side='bottom', fill='x')
treescrolly.pack(side='right', fill='y')

# Outer frame that houses Load File funtions, Merge File functions, Save function, and Cancel button.
file_frame = tk.LabelFrame(root)
file_frame.place(height=445, width=1000, relx=0, rely=0.53)
 
save_button = tk.Button(file_frame, text='Save File', command=lambda: save_file())
save_button.place(height=45, width=150, relx=0.35, rely=0.87)

cancel_button = tk.Button(file_frame, text='Cancel', command=root.quit)
cancel_button.place(height=45, width=150, relx=0.51, rely=0.87)

# First file frame to load in first file.
label_frame_1 = tk.LabelFrame(file_frame, text='Load First File')
label_frame_1.place(height=75, width=975, relx=0.01, rely=0)
 
label_file_1 = ttk.Label(label_frame_1, text='No File Selected')
label_file_1.place(relx=0, rely=0.25)

browse_button_1 = tk.Button(label_frame_1, text='Browse File 1', command=lambda: file_dialog('first_file'))
browse_button_1.place(height=45, width=150, relx=0.84, rely=0)

# Second file frame to load in second file.
label_frame_2 = tk.LabelFrame(file_frame, text='Load Second File')
label_frame_2.place(height=75, width=975, relx=0.01, rely=0.19)

label_file_2 = ttk.Label(label_frame_2, text='No File Selected')
label_file_2.place(relx=0, rely=0.25)

browse_button_2 = tk.Button(label_frame_2, text='Browse File 2', command=lambda: file_dialog('second_file'))
browse_button_2.place(height=45, width=150, relx=0.84, rely=0)

# Third file frame to add Radiobuttons for each join type and Dropdown index for the join method.
label_frame_3 = tk.LabelFrame(file_frame, text='Merge Files')
label_frame_3.place(height=200, width=975, relx=0.01, rely=0.40)

listbox_frame = tk.LabelFrame(label_frame_3)
listbox_frame.place(height=175, width=485, relx=0, rely=0)

listbox_label = tk.Label(listbox_frame, text='Choose an index to join on:')
listbox_label.place(relx=0, rely=0)

listbox = tk.Listbox(listbox_frame)
listbox.place(height=150, width=390, relx=0, rely=0.10)

action_button = tk.Button(listbox_frame, text='Get Index', command=lambda: get_index())
action_button.place(height=150, width=75, relx=0.83, rely=0.10)

radio_frame = tk.LabelFrame(label_frame_3)
radio_frame.place(height=175, width=485, relx=0.50, rely=0)

v = tk.IntVar()

radio_label = tk.Label(radio_frame, text='Choose a join method:')
radio_label.place(relx=0, rely=0)

rb1 = tk.Radiobutton(radio_frame, text='inner', variable=v, value=1, command=lambda: merge_files('inner')) # rb_selection('inner')
rb1.place(relx=0, rely=0.15)

rb2 = tk.Radiobutton(radio_frame, text='left', variable=v, value=2, command=lambda: merge_files('left'))
rb2.place(relx=0.14, rely=0.15)

rb3 = tk.Radiobutton(radio_frame, text='right', variable=v, value=3, command=lambda: merge_files('right'))
rb3.place(relx=0.26, rely=0.15)

rb4 = tk.Radiobutton(radio_frame, text='outer', variable=v, value=4, command=lambda: merge_files('outer'))
rb4.place(relx=0.40, rely=0.15)

method_label = ttk.Label(radio_frame, text='No Method Selected')
method_label.place(relx=0, rely=0.45)

# Function to open file browser and display file path on application for both files.
def file_dialog(filename=['first_file', 'second_file']):
    if filename == 'first_file':
        file_one = filedialog.askopenfilename(initialdir='/', title='Select A File', filetypes=(('Excel Files', '*.xlsx'), ('All Files', '*.*')))
        label_file_1['text'] = file_one
    elif filename == 'second_file':
        file_two = filedialog.askopenfilename(initialdir='/', title='Select A File', filetypes=(('Excel Files', '*.xlsx'), ('All Files', '*.*')))
        label_file_2['text'] = file_two
    else:
        return None

# Decoration function to clear listbox when new files are selected.
def clear_list(func):
    def wrapper():
        listbox.delete(0, 'end')
        func()
    return wrapper

# Function to extract file paths and the get the combined column headers for each Excel file in order to select the index to join on.
@clear_list
def get_index():
    path_one = label_file_1.cget('text')
    path_two = label_file_2.cget('text')
    path_list = []
    if path_one and path_two != 'No File Selected':
        path_one_string = r'{}'.format(path_one)
        path_two_string = r'{}'.format(path_two)
        path_list.append(path_one_string)
        path_list.append(path_two_string)
        joined_file = ExcelJoin(path_list[0], path_list[1])
        index_list = joined_file.intersection()
        for item in index_list:
            listbox.insert('end', item)
    if index_list == []:
        tk.messagebox.showerror('Information', 'No Index Match.')
    else:
        return None

# Function to display definition of each join type when a readio button is selected.
def merge_files(how=['inner', 'left', 'right', 'outer']):
    path_one = label_file_1.cget('text')
    path_two = label_file_2.cget('text')
    path_list = []
    path_one_string = r'{}'.format(path_one)
    path_two_string = r'{}'.format(path_two)
    path_list.append(path_one_string)
    path_list.append(path_two_string)
    joined_file = ExcelJoin(path_list[0], path_list[1])

    for item in listbox.curselection():
        selection = listbox.get(item)
         
        if how == 'inner':
            new_df = joined_file(on=selection, how=how)
            load_dataframe(new_df)
            definition = 'Inner join produces an output data frame of only those rows for which the condition is \nsatisfied in both rows.'
            method_label['text'] = definition
        elif how == 'left':
            new_df = joined_file(on=selection, how=how)
            load_dataframe(new_df)
            definition ='Left join operation provides all the rows from 1st dataframe and matching rows from \nthe 2nd dataframe. If the rows are not matched in the 2nd dataframe then they \nwill be replaced by NaN.'
            method_label['text'] = definition
        elif how == 'right':
            new_df = joined_file(on=selection, how=how)
            load_dataframe(new_df)
            definition = 'Right join is somewhat similar to left join in which the output dataframe will consist \nof all the rows from the 2nd dataframe and matching rows from the 1st dataframe. \nIf the rows are not matched in the 1st row then they will be replaced by NaN.'
            method_label['text'] = definition
        elif how == 'outer':
            new_df = joined_file(on=selection, how=how)
            load_dataframe(new_df)
            definition = 'Outer join provides the output dataframe consisting of rows from both the dataframes. \nValues will be shown if rows are matched otherwise NaN will be shown for rows that \ndo not match.'
            method_label['text'] = definition
        else:
            return None

def load_dataframe(df):
    clear_data()
    tv1['column'] = list(df.columns)
    tv1['show'] = 'headings'
    for column in tv1['columns']:
        tv1.heading(column, text=column)
 
    df_excel_rows = df.to_numpy().tolist()
    for row in df_excel_rows:
        tv1.insert('', 'end', values=row)

    return None

def clear_data():
    tv1.delete(*tv1.get_children())

# Function to save file to desktop.
def save_file():
    files = [('CSV File', '*.csv')]
    savefile = filedialog.asksaveasfile(initialdir='/', title='Save File', defaultextension='*.csv', filetypes=files, initialfile='new_joined_file', mode='w')
    headers = list(tv1['columns'])
    items_list = []
    data = {}

    for header in headers:
        data[header] = []

    for item in tv1.get_children():
        items = tv1.item(item)['values']
        items_list.append(items)

        count = 0

        for arg in items:
            key = headers[count]
            temp_list = list(data[key])
            temp_list.append(arg)
            data[key] = temp_list
            count += 1

    df = pd.DataFrame(data=data)
    df.to_csv(str(savefile.name))

root.mainloop()
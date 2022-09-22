from distutils.cmd import Command
from tkinter import ttk
import pandas as pd
import xlwings as xw
import openpyxl as openpyxl
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile

# mark_cap_file_path = 'March31st_Market_Caps.xlsx'
mark_cap_file_path = ''


# Intializing empty dataframe
main_df = pd.DataFrame()

lookup_df = pd.DataFrame()

# wb = xw.Book(mark_cap_file_path)
# sheet2 = wb.sheets('Sheet2')

# mark_cap_df = pd.read_excel(mark_cap_file_path, sheet_name='Sheet1')
# buy_qnty_df = pd.read_excel(mark_cap_file_path, sheet_name='Sheet2')


def init_dataframe(filename):
    excel_file = pd.ExcelFile(filename)
    global mark_cap_file_path
    mark_cap_file_path = filename
    print(excel_file.sheet_names)
    show_relevant_widgets(list(excel_file.sheet_names))

def show_relevant_widgets(sheet_names):
    padding = 0
    choose_working_sheet = Label(
        text='Choose working sheet'
    )
    choose_working_sheet.place(x=15, y = 45)
    for sheet_name in sheet_names:
        wcb = Checkbutton(
            root,
            text=sheet_name,
            variable=working_sheet,
            onvalue=sheet_name,
            offvalue='Sheet Not Selected',
            command=lambda: open_working_sheet()
        )
        wcb.place(x=(15 + padding), y=75)
        padding = padding + 70 + len(sheet_name)
        
    choose_lookup_sheet = Label(
        text='Choose lookup sheet'
    )
    choose_lookup_sheet.place(x=15,y=105)
    padding = 0
    for sheet_name in sheet_names:
        lcb = Checkbutton(
            root,
            text=sheet_name,
            variable=lookup_sheet,
            onvalue=sheet_name,
            offvalue='Sheet Not Selected',
            command=lambda: initialize_lookup_df()
        )
        lcb.place(x=(15 + padding), y=135)
        padding = padding + 70 + len(sheet_name)


def open_working_sheet():
    # print(selected_sheet_name.get())
    # print(mark_cap_file_path)
    global main_df
    main_df = pd.read_excel(mark_cap_file_path, sheet_name=working_sheet.get())
    print(list(main_df))
    show_working_sheet_columns()

def initialize_lookup_df():
    global lookup_df
    lookup_df = pd.read_excel(mark_cap_file_path, sheet_name=lookup_sheet.get())
    print(list(lookup_df))
    show_lookup_sheet_columns()


def show_working_sheet_columns():
    padding = 0
    choose_ws_column = Label(
        text='Choose Working sheet column'
    )
    choose_ws_column.place(x=15,y=165)
    for column_name in list(main_df):
        ws_column_cb = Checkbutton(
                root,
                text=column_name,
                variable=left_lookup_column,
                onvalue=column_name,
                offvalue='Column Not Selected',
            )
        ws_column_cb.place(x=15+padding, y = 195)
        padding = padding + 70 + len(column_name)
 


def show_lookup_sheet_columns():
    choose_ls_column = Label(
        text='Choose Lookup sheet column'
    )
    choose_ls_column.place(x=15,y=225)
    padding = 0
    for column_name in list(lookup_df):
        ls_column_cb = Checkbutton(
                root,
                text=column_name,
                variable=right_lookup_column,
                onvalue=column_name,
                offvalue='Column Not Selected',
                command=lambda: select_column_to_merge()
            )
        ls_column_cb.place(x=15+padding, y = 255)
        padding = padding + 70 + len(column_name)


def select_column_to_merge():
    choose_ls_column = Label(
        text='Choose column to merge'
    )
    choose_ls_column.place(x=15,y=285)
    padding = 0
    for column_name in list(lookup_df):
        ls_column_merge_cb = Checkbutton(
                root,
                text=column_name,
                variable=column_to_merge,
                onvalue=column_name,
                offvalue='Column Not Selected',
            )
        ls_column_merge_cb.place(x=15+padding, y = 315)
        padding = padding + 70 + len(column_name)
    
    perform_lookup_btn = Button(
        root,
        text='Perform Lookup',
        command=lambda: perform_lookup()
    )
    perform_lookup_btn.pack()


def perform_lookup_pass():
    print('Lookup to be performewd')
    pass



def perform_lookup():


    res = main_df.merge(lookup_df[[right_lookup_column.get(),column_to_merge.get()]],
     left_on=left_lookup_column.get(), right_on=right_lookup_column.get(), how='left')
    print(res)
    # save_wb(res, column_to_merge)
    with pd.ExcelWriter(mark_cap_file_path, mode='a') as writer:
        res.to_excel(writer, sheet_name='new_sheet1', index=False)


# def save_wb(merged_df, merge_column):
#     columns = [merge_column]
#     sheet2.range('C1').options(index=False).value = merged_df[columns]
#     sheet2.range('C1:C1').color = (253,233,217)

def is_valid_column(column_name, dataframe):
    if column_name not in dataframe.columns:
        print('Invalid column name')
        quit()

def open_file():
    file_path = askopenfile(mode='r', filetypes=[('Excel Files', '*xlsx')])
    if file_path is not None:
        init_dataframe(filename=file_path.name)


# lookup_input()


if __name__ == '__main__':
    # create a GUI window
    root = Tk()
 
    root.geometry("700x500")
    # set the background colour of GUI window
    root.title('PyLookupApp')
    
    working_sheet = StringVar()
    lookup_sheet = StringVar()
    left_lookup_column = StringVar()
    right_lookup_column = StringVar()
    column_to_merge = StringVar()

    excel_upload_text = Label(text='Upload File', background='yellow', foreground='black')

    excel_upload_text.place(x=15,y=25)
    
    excel_upload_button = Button(root, command=lambda: open_file(), text='Upload File', width='30')

    excel_upload_button.place(x=100,y=25)




    mainloop()

# upld = Button(
#     ws, 
#     text='Upload Files', 
#     command=uploadFiles
#     )
# upld.grid(row=3, columnspan=3, pady=10)



# ws.mainloop()



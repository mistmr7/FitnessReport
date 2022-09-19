from email import message
import pandas as pd
import numpy as np
import openpyxl
import os
import getpass
from tkinter import Tk, messagebox, Toplevel
from tkinter.filedialog import askopenfilename

def create_summary_entry(df, day, start_time, **kwargs):
    if kwargs:
        result_df = df[(df["Day"] == day) & (df["Start Time"] == start_time) & (df.Class == kwargs["Class"])]
    else:
        result_df = df[(df["Day"] == day) & (df["Start Time"] == start_time)]
    new_row = []
    row_1 = result_df.iloc[0]
    new_row.append(row_1.Day)
    new_row.append(row_1["Start Time"])
    new_row.append(row_1["End Time"])
    new_row.append(row_1["Class"])
    new_row.append(row_1.Studio)
    new_row.append(row_1.Category)
    new_row.append(row_1.Location)
    attendance = result_df.Attendance.sum()
    classes = len(result_df)
    ppl_per_class = f"{attendance/classes:.2f}"
    new_row.append(attendance)
    new_row.append(classes)
    new_row.append(ppl_per_class)
    return new_row


def create_table():
    resulting_df = pd.DataFrame(
        columns=[
            "Day",
            "Start Time",
            "End Time",
            "Class",
            "Studio",
            "Category",
            "Location",
            "Total Attendance",
            "Total Classes",
            "People per Class",
        ]
    )
    return resulting_df


def add_next_row(table, row):
    table.loc[len(table.index)] = row
    return table


def write_the_file(filename, new_table, sheet):
    try:
        with pd.ExcelWriter(filename, engine="openpyxl", mode="a") as writer:
            new_table.to_excel(writer, sheet_name=sheet, index=False)
        workbook = openpyxl.load_workbook(filename)
        ws = workbook[sheet]
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length
    except ValueError:
        workbook = openpyxl.load_workbook(filename)
        std = workbook[sheet]
        workbook.remove(std)
        workbook.save(filename)
        write_the_file(filename, new_table, sheet)
    else:
        try:
            workbook.remove(workbook['Sheet'])
        except KeyError:
            pass
        workbook.save(filename)

def create_am_pm_headers():
    resulting_df = pd.DataFrame(
        columns=[
            "Day",
            "Time",
            "Total Attendance",
            "Total Classes",
            "People per Class",
        ]
    )
    return resulting_df


def create_am_pm_summary_entry(df, day, start_time, **kwargs):
    result_df = df[(df["Day"] == day) & (df['Start Time'].str.contains(start_time))]
    new_row = []
    new_row.append(day)
    new_row.append(start_time)
    attendance = result_df['Total Attendance'].sum()
    classes = result_df['Total Classes'].sum()
    people_per_class = attendance/classes if classes > 0 else 0
    print(people_per_class)
    new_row.append(str(attendance))
    new_row.append(str(classes))
    new_row.append(str(round(people_per_class, 2)))
    return new_row


def run_program():
    root = Tk()
    texto = Toplevel(root)
    messagebox.showinfo('Fitness Report Summary Creator','Please select the excel fitness report that you would like to analyze.', parent=texto)
    filename = askopenfilename(initialdir=f'C:\\Users\\{getpass.getuser()}\\Downloads', title='Select the excel file you would like to use')
    df = pd.read_excel(filename, usecols="A:M", keep_default_na=False)
    start_index = df[df.Day == "Monday"].first_valid_index()
    end_index = df[50:][df.Day == "Monday"].first_valid_index()
    new_table = create_table()
    for i in range(start_index - 1, end_index):
        day = df.iloc[i].Day
        time = df.iloc[i]["Start Time"]
        fitness_class = df.iloc[i].Class
        add_next_row(new_table, create_summary_entry(df, day, time, Class=fitness_class))
    write_the_file(filename, new_table, "Results Summary")

    # Create Second Sheet
    df2 = pd.read_excel(filename, sheet_name='Results Summary')
    new_table2 = create_am_pm_headers()
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    for day in days:
        add_next_row(new_table2, create_am_pm_summary_entry(df2, day, 'am'))
        add_next_row(new_table2, create_am_pm_summary_entry(df2, day, 'pm'))
    new_table2['People per Class'] = new_table2['People per Class'].replace(np.nan, 0)
    write_the_file(filename, new_table2, 'AM PM Summary')
    os.startfile(filename)
    root.destroy()


if __name__ == '__main__':
    try:
        run_program()
    except FileNotFoundError:
        messagebox.showerror('Error', 'No file was selected. Please run the program again and select the fitness report you would like to analyze.')
    except AttributeError:
        messagebox.showerror('Error', 'Are you sure you selected the correct file? Run the program again and select the downloaded fitness report.')
    except PermissionError:
        messagebox.showerror('Error', 'Please save and close the excel file you are wanting to analyze and run the program again.')
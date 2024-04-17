## Author: Benedyct Liew
## Start Date: 12/03/2024
## Desc: Create an excel document that tracks the progress of improvement in typing speed.


# Imports
import openpyxl as xl
import os
import datetime
import tkinter as tk
#from PIL import Image
#import pytesseract

#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Create the document
# @wb_name: name of the file in directory
def create_excel(wb_name):
    wb = xl.Workbook()
    ws = wb.active
    ws.title = "Results"

    # Headers 
    ws.append(["Date", "Time",  "WPM", "Char", "Accuracy (%)"])

    # Set the size of the two larger columns
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['E'].width = 13

    wb.save(wb_name)


# If file exists then open the file
# @wb_name: name of the file in directory
def load_file(wb_name):
    wb = xl.load_workbook(wb_name)
    ws = wb.active

    # Appends data
    append_data(ws)

    wb.template = False
    wb.save(wb_name)


# Append the data into the spreadsheet
# @ws: the pointer to the worksheet
def append_data(ws):
    date = datetime.datetime.now().strftime("%d/%m/%Y")
    time = datetime.datetime.now().strftime("%X")

    input_window = tk.Tk()
    input_window.title('Results Input')
    input_window.geometry("300x100")

    # Submits the entry into the document
    def submit():
        try:
            ws.append([date, time, int(WPM.get()),int(char.get()),int(acc.get())])
        except ValueError:
            pass

    # Closes the pop up window
    def destroy():
        input_window.destroy()

    # Sets string for entry labels
    wpm_var = tk.StringVar()
    cha_var = tk.StringVar()
    acc_var = tk.StringVar()

    wpm_var.set("")
    cha_var.set("")
    acc_var.set("")

    tk.Label(input_window, text='Enter the WPM').grid(row=0)
    tk.Label(input_window, text='Enter the char/min').grid(row=1)
    tk.Label(input_window, text='Enter the Accuracy in Percent').grid(row=2)

    WPM = tk.Entry(input_window, textvariable  = wpm_var)
    char = tk.Entry(input_window, textvariable  = cha_var)
    acc = tk.Entry(input_window, textvariable  = acc_var)

    WPM.grid(row=0, column=1)
    char.grid(row=1, column=1)
    acc.grid(row=2, column=1)

    exit = tk.Button(input_window, text='Submit', command=lambda: [f for f in [submit(), destroy()]], width=10)
    exit.grid(row=3, column=1)

    input_window.mainloop()



# The main function
def main():
    filename = "Desktop\Typing speed\Typing.xlsx"
    if not os.path.isfile(filename):
        create_excel(filename)
        load_file(filename)
    else:
        load_file(filename)


# Call`` the main function 
main()
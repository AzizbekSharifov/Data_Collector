'''                    !!!!!!!! IMPORTANT !!!!!!!!!!!
TO RUN THE CODE YOU HAVE TO "PIP INSTALL TK", "PIP INSTALL OPENPYXL", "PIP INSTALL PANDAS"
AND CHANGE THE LOCATION IN LINE 28, 45, 66, 119, 145, 151, 159, 161, 171 TO LOCATION IN YOUR COMPUTER WHICH IS WHERE YOU PUT THIS FILE
'''


# ---------------- IMPORTS ----------------
from openpyxl import load_workbook
import random
import PySimpleGUI as sg
from datetime import datetime
from tkinter import *
import pandas as pd
print("______________________________")
print("Program is working now....")              #=> when program is start to working, these word come out in terminal
print("______________________________")

#DATA ENTRY WINDOW USING PySimpleGUI
sg.theme('DarkAmber')
entry_words_list = ['Welcome!!!', 'Greetings!!!', 'Nice to see you!🖐', 'Hi!👋', 'Hello!!👋', 'Wassup😎' ]

layout = [[sg.Text(f'{random.choice(entry_words_list)}')],
          [sg.Text('Please fill out the following fields:')],
          [sg.Text('Name'), sg.Push(), sg.InputText(key='Name')],
          [sg.Text('Student_ID'), sg.Push(), sg.InputText(key='Student_ID')],
          [sg.Text('Grade'), sg.Push(), sg.InputText(key='Grade')],
          [sg.Button('Save'), sg.Button("Clear"), sg.Button('Search'), sg.Button('Exit')]]

window = sg.Window('Data Entry', layout, element_justification='center', icon='F:/INHA University 2022/Software Programming/MID-Term project/Icons/Data-Entry.ico')

#--------------------------------------------------

#Clear button's function
def clear_input():
    for key in values:    #=> this helps to work clear button
        window[key]('')
    return None
#------------------------

#LOOP is starting (without stopping)
while True:

    # Exit button
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        option = sg.popup_yes_no("Are you Really want to Quit!!", title="Exit", icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/Exit.ico")
        if option == 'Yes':
            break
        else:                                   #=>  this helps to us give option when EXIT button clicked
            pass
    #-------------------------------------------------

    # Clear button
    if event == "Clear":
        clear_input()
    #___________________

    # Search button
    if event == "Search":   #=> from here started search window
        # USING TKINTER

        #SEARCH WINDOW
        search_window = Tk(className="Search window", )
        search_window.configure(bg="#2c2825")
        search_window.resizable(width=False, height=False)     #=> main window
        search_window.geometry("500x300")
        search_window.iconbitmap('F:/INHA University 2022/Software Programming/MID-Term project/Icons/search.ico')
        #---------------------------------------------

        # ALERT!!! SEARCHING BY ID
        mylabel = Label(search_window, text="Start Searching by ID...", font=("Helvetica", 14), fg="black",bg="#fdcb52")
        mylabel.pack(pady=20)
        #------------------------------

        # Create entry box
        entry_box = Entry(search_window, font=("Helvetica", 20), bg="#705e52", fg="#fdcb52")
        entry_box.get()
        entry_box.pack()        #=> here user can input Student ID
        #---------------------

        # answer_ box
        answer_box = Listbox(search_window, font=("Helvetica", 16), bg="#705e52", fg='black', width=50) #=> In search result is showed in answer box
        #----------------

        # Create Search button's function
        def search_func():
            try:                                                #=> Button function, search button don't worked without this.
                data = pd.read_excel('DATA_ Base.xlsx')
                data = data.to_dict()
        #-------------------------------------------------

                #If you write word or something wrong
                if not entry_box.get().isdigit():
                    answer_box.delete(0, 2)
                    answer_box.insert(0, f"You entered something wrong.")               #=> user always input a number in entry box, if not so error window come out.
                    answer_box.insert(1, f"You are only allowed to input numbers!!!")
                #--------------------------------------------------------

                # if user input wrong ID
                else:
                    if int(entry_box.get()) not in list(data['Student_ID'].values()):
                        answer_box.delete(0, 2)                                         #=> if student ID not in Excel file, these rows are worked
                        answer_box.insert(0, f"This data is not available!!!")
                #------------------------------------------------------

                    #if everything is good
                    else:
                        num = list(data['Student_ID'].values()).index(int(entry_box.get()))
                                                                                            #=> if everything is good the result shown in answer box
                        answer_box.delete(0, 3)
                        answer_box.insert(0, f"           №: {data['№'][num]}")
                        answer_box.insert(1, f"           Name: {data['Name'][num]}")
                        answer_box.insert(2, f"           Grade: {data['Grade'][num]}")
                        answer_box.insert(3, f"           Time_stamp: {data['Time_stamp'][num]}")
                    #___________________________________________________________

             # if Excel file opened during working a program
            except PermissionError:                                             #=> user should close Excel file during the working with the program
                 sg.popup('Please close the Excel file!!!', title="ALERT!",
                          icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/WARMING.ico")
                 #---------------------------------------------------

        #SEARCH button
        search_button = Button(search_window, text="search", font=('Times New Roman', 13, 'bold'), bg="#fdcb52", fg='black',
                           relief='ridge', pady=5, command=search_func)         #=> design of search button
        search_button.pack()

        answer_box.pack(pady=20)

        search_window.mainloop()
        #---------------------------------------------


    # Save button and worked with an Excel file where the dates are added
    if event == 'Save':
                                                            #=> If save button is clicked these rows are worked
        try:
            wb = load_workbook('DATA_ Base.xlsx')    #=> file which all datum are collected
            sheet = wb['Sheet1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S") # used the datetime package to indicate when the data was inserted.

            #if user does not enter any information in input box,
            if values['Name'] == '' or values['Student_ID'] == '' or values['Grade'] == '':
                sg.popup("Please fill out all fields", title="Error!",                          #=> if user does not enter anything in input boxes
                               icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/WARMING.ico")
            #-------------------------------------------------------------------

            #IF user input word instead of number in Student_ID box, the extra window is opened
            elif not values['Student_ID'].isdigit():
                sg.popup("For student ID input only numbers", title="Error!",               #=> if user input word in Student_ID section
                               icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/WARMING.ico")
                window['Student_ID'].update(value='')
            #---------------------------------------------------------------------------

             #IF user fill out all the input boxes, the datum are saved in Excel file.
            else:
                data = [ID, values['Name'], values['Student_ID'], values['Grade'], time_stamp]

                sheet.append(data)                                          #=> if everything is good all datum are saved
                sg.popup('DATA', 'Successfully Saved!',
                         icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/Save.ico")  # Extra small window opened when data saved successfully
                window['Name'].update(value='')
                window['Student_ID'].update(value='')
                window['Grade'].update(value='')
                window['Name'].set_focus()

                wb.save('DATA_ Base.xlsx')
            #-------------------------------------------------------------------

        except PermissionError:
            sg.popup('File in use','Close Excel file and try again.',title="ALERT!", icon="F:/INHA University 2022/Software Programming/MID-Term project/Icons/WARMING.ico")

    #----------------------------------------------------

window.close()

print("______________________________")
print("Program Stopped!!!")                     #=> when program is stopped, these word come out in terminal
print("______________________________")

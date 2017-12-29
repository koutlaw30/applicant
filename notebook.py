from tkinter import ttk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog as fd
from tkinter import *
from tkinter import messagebox
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import PyPDF2
import re
import pyperclip
import os



def file_content():
    """ this function extracts content from a PDF file """
    pdf_stream = open('resume.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdf_stream)
    first_page_content = pdfReader.getPage(0).extractText().replace('\n', '')
    pdf_stream.close()
    return first_page_content


def email_finder(string_content):
    """ returns emails in a list """
    email_list = []
    email_pattern = r"([\w\.-]+)@([\w\.-]+)(\.[\w\.]+)"
    emails = re.findall(pattern=email_pattern, string=string_content)
    for email_ in emails:
        email = email_[0] + '@' + email_[1] + email_[2]
        email_list.append(email)

    return email_list


pdf_content = file_content()
#print(email_finder(pdf_content))

def deleteselectedapplicants():
         selection = applicantlistbox.curselection()
         applicantlistbox.delete(selection[0])

def deleteallapplicants():
         applicantlistbox.delete(0,END)


def deleteallpanelmembers():
         panellistbox.delete(0,END)


def deleteselectedemails():
         selection = email_list_box.curselection()
         email_list_box.delete(selection[0])

def addemails():
         email_list_box.insert(0, myemail.get())

def addapplicants():
         applicantlistbox.insert(0, myapplicant.get())

         
def createworksheets():
        z = StringVar()
        wb = openpyxl.load_workbook('My Name Application.xlsx')
        #ws2 = wb.create_sheet(str((applicantlistbox.get(applicantlistbox.curselection()))))
        #wb[(str((applicantlistbox.get(applicantlistbox.curselection()))))]['A1'] = "INDIVIDUAL SCORE SHEET"
        for z, listbox_entry in enumerate(applicantlistbox.get(0, END)):
            for y, listbox_entry in enumerate(panellistbox.get(0, END)):
            #If NumberofInterviewQuestions.Value = 6 Then

                """Your old method
                # ws2 = wb.create_sheet(str((applicantlistbox.get(z) + " " + str(ffint.get()))))
                
                #Applying Font to all Worksheets
                a1 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A1']
                ft = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].font = Font(name='Arial', size=12, bold = True)
                ft1 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].font = Font(name='Arial', bold = True)
                a1.font = ft
                a9 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A9']
                a12 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A12']
                a15 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A15']
                a18 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A18']
                a21 = wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A21']
                a9.font = ft1
                a12.font = ft1
                a15.font = ft1
                a18.font = ft1
                a21.font = ft1
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A1'] = "INDIVIDUAL SCORE SHEET"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A2'] = "RPA: " + str(myrpa.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A4'] = "Applicant’s Name: " + str((applicantlistbox.get(z)))
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A6'] = "Panel Member’s Name: " + str(panel_member1.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A8'] = "Individual Application Scores"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A9'] = "Max- " + str(maxscore.get()) 
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A11'] = "A. Experience"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A12'] = "Max- " + str(experience.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A14'] = "B. Education"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A15'] = "Max- " + str(education.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A17'] = "C. Training"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A18'] = "Max- " + str(training.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A20'] = "D. Awards"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A21'] = "Max- " + str(awards.get())
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))]['A24'] = "Total Application Points"
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].merge_cells('A1:G1', 'A2:F1')
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].row_dimensions[1].height = 20
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].row_dimensions[2].height = 18
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].row_dimensions[3].height = 18
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].row_dimensions[4].height = 15
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].row_dimensions[5-57].height = 15.75
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['A'].width = 30.86
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['B'].width = 6.57
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['C'].width = 8.43
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['D'].width = 8.57
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['E'].width = 15.29
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['F'].width = 19
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].column_dimensions['G'].width = 21.7
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.PrintArea = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.LeftHeader = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.CenterHeader = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.RightHeader = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.LeftFooter = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.CenterFooter = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.RightFooter = ""
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.FitToPagesWide = 1
                wb[(str((applicantlistbox.get(z)+ " " + str(ffint.get()))))].page_setup.FitToPagesTall = 1
                """



                # new method
                formatted_name = f'''{applicantlistbox.get(z).split(', ')[0]} {applicantlistbox.get(z).split(', ')[1][0]} {panellistbox.get(y).split(' ')[0][0]} {panellistbox.get(y).split(' ')[1][0]}'''
                ws2 = wb.create_sheet(formatted_name)

                a1 = wb[formatted_name]['A1']
                ft = wb[formatted_name].font = Font(name='Arial', size=12, bold=True)
                ft1 = wb[formatted_name].font = Font(name='Arial', bold=True)
                a1.font = ft
                a9 = wb[formatted_name]['A9']
                a12 = wb[formatted_name]['A12']
                a15 = wb[formatted_name]['A15']
                a18 = wb[formatted_name]['A18']
                a21 = wb[formatted_name]['A21']
                a9.font = ft1
                a12.font = ft1
                a15.font = ft1
                a18.font = ft1
                a21.font = ft1
                wb[formatted_name]['A1'] = "INDIVIDUAL SCORE SHEET"
                wb[formatted_name]['A2'] = "RPA: " + str(myrpa.get())
                wb[formatted_name]['A4'] = "Applicant’s Name: " + str(
                    (applicantlistbox.get(z)))
                wb[formatted_name]['A6'] = "Panel Member’s Name: " + str(
                    (panellistbox.get(y)))
                wb[formatted_name]['A8'] = "Individual Application Scores"
                wb[formatted_name]['A9'] = "Max- " + str(maxscore.get())
                wb[formatted_name]['A11'] = "A. Experience"
                wb[formatted_name]['A12'] = "Max- " + str(experience.get())
                wb[formatted_name]['A14'] = "B. Education"
                wb[formatted_name]['A15'] = "Max- " + str(education.get())
                wb[formatted_name]['A17'] = "C. Training"
                wb[formatted_name]['A18'] = "Max- " + str(training.get())
                wb[formatted_name]['A20'] = "D. Awards"
                wb[formatted_name]['A21'] = "Max- " + str(awards.get())
                wb[formatted_name]['A24'] = "Total Application Points"
                wb[formatted_name].merge_cells('A1:G1', 'A2:F1')
                wb[formatted_name].row_dimensions[1].height = 20
                wb[formatted_name].row_dimensions[2].height = 18
                wb[formatted_name].row_dimensions[3].height = 18
                wb[formatted_name].row_dimensions[4].height = 15
                wb[formatted_name].row_dimensions[5 - 57].height = 15.75
                wb[formatted_name].column_dimensions['A'].width = 30.86
                wb[formatted_name].column_dimensions['B'].width = 6.57
                wb[formatted_name].column_dimensions['C'].width = 8.43
                wb[formatted_name].column_dimensions['D'].width = 8.57
                wb[formatted_name].column_dimensions['E'].width = 15.29
                wb[formatted_name].column_dimensions['F'].width = 19
                wb[formatted_name].column_dimensions['G'].width = 21.7
                wb[formatted_name].page_setup.PrintArea = ""
                wb[formatted_name].page_setup.LeftHeader = ""
                wb[formatted_name].page_setup.CenterHeader = ""
                wb[formatted_name].page_setup.RightHeader = ""
                wb[formatted_name].page_setup.LeftFooter = ""
                wb[formatted_name].page_setup.CenterFooter = ""
                wb[formatted_name].page_setup.RightFooter = ""
                wb[formatted_name].page_setup.FitToPagesWide = 1
                wb[formatted_name].page_setup.FitToPagesTall = 1

        #ws2 = wb.create_sheet(str((applicantlistbox.get(z))) + " 1 Score Sheet PM " + ffint.get() + " " & flint.get())
        wb.save('My Name Application.xlsx')
        messagebox.showinfo("Update Spreadsheet", "The spreadsheet has been updated")


def importpaneltextfile():
        print ("Opening and closing the file.")
        filename = fd.askopenfilename(initialdir="/", title="Select file",
                                      filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
        with open(filename, 'r') as file:
            [panellistbox.insert(1, item) for item in file.readlines()]

def panel1update():
            PanelCombo1['values'] = ['jjjj']


def importapptextfile():
        applicantlistbox.delete(0,END)
        print ("Opening and closing the file.")
        filename = fd.askopenfilename(initialdir="/", title="Select file",
                                      filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
        with open(filename, 'r') as file:
            [applicantlistbox.insert(1, item) for item in file.readlines()]
            Button(page2, text='Import Applicants', activebackground="white", bd=3,bg="white",width=13, command=importapptextfile).grid(row=2, column=2, sticky=W, pady=4)
            

def PDF():
        email_list_box.delete(0,END)
        print ("Opening and closing the pdf file.")
        filename = fd.askopenfilename(initialdir="/", title="Select file",
                                      filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
        with open(filename, 'r') as file:
            [email_list_box.insert(1, item) for item in file.readlines()]
            Button(page1, text='Import PDF', activebackground="white", bd=3,bg="white",width=13, command=PDF).grid(row=2, column=3, sticky=W, pady=4)


def printpanel(evt):
        value=str((panellistbox.get(panellistbox.curselection())))
        print (value)         

def printapplicants(evt):
        value=str((applicantlistbox.get(applicantlistbox.curselection())))
        print (value)



def UpdatePage2():
        wb = openpyxl.load_workbook('Applicant Application.xlsx')
        wb["Page 2"]['A5'] = "Position: " + str(myjobtitle.get())
        wb["Page 2"]['A6'] = "Cert# " + str(mycert.get())
        wb["Page 2"]['B8'] = "EXPERIENCE-" + str(experience.get())
        wb["Page 2"]['F8'] = "EDUCATION-" + str(education.get())
        wb["Page 2"]['J8'] = "TRAINING-" + str(training.get())
        wb["Page 2"]['N8'] = "AWARDS-" + str(awards.get())
        wb["Page 2"]['R8'] = "INTERVIEW QUESTIONS-" + str(maxinterview.get())
        wb["Page 2"]['A9'] = "Panel Initials"
        row_count = wb["Page 2"].max_row
        print (row_count)
        wb["Page 2"]['A' + row_count] = "SIGNATURE"
        


        

#ws.max_row
#IncludeZeros

##Sheets("Page 2").Range("A" & applicantlastrow + 4).Value = "SIGNATURE"
##Sheets("Page 2").Range("F" & applicantlastrow + 4).Value = "SIGNATURE"
##Sheets("Page 2").Range("P" & applicantlastrow + 4).Value = "SIGNATURE"
##
##Sheets("Page 2").Range("A" & applicantlastrow + 4).Font.Bold = True
##Sheets("Page 2").Range("F" & applicantlastrow + 4).Font.Bold = True
##Sheets("Page 2").Range("P" & applicantlastrow + 4).Font.Bold = True
##Sheets("Page 2").Range("A" & applicantlastrow + 4).Font.Size = 12
##Sheets("Page 2").Range("F" & applicantlastrow + 4).Font.Size = 12
##Sheets("Page 2").Range("P" & applicantlastrow + 4).Font.Size = 12

        wb.save('Applicant Application.xlsx')
        messagebox.showinfo("Update Spreadsheet", "The spreadsheet has been updated")

def UpdateSpreadsheet():
        wb = openpyxl.load_workbook('My Name Application.xlsx')
        wb["Sheet1"]['A1'] = "Your Name: " + str(myjobtitle.get())
        wb.save('My Name Application.xlsx')
        messagebox.showinfo("Update Spreadsheet", "The spreadsheet has been updated")


def chairperson():
    if chairperson_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="ChairPerson", fg="red").grid(row=2, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=10).grid(row=2, column=5, sticky=W)
        
def chairperson_eeo():
    if chairperson_eeo_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="ChairPerson/EEO", fg="red").grid(row=2, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=13).grid(row=2, column=5, sticky=W)

def sme_eeo():
    if sme_eeo_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="SME/EEO", fg="red").grid(row=4, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=7).grid(row=5, column=5, sticky=W)
    


def sme_one():
    if sme_one_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="SME", fg="red").grid(row=4, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=4).grid(row=4, column=5, sticky=W)
    

def sme_two():
    if sme_two_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="SME", fg="red").grid(row=6, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=4).grid(row=6, column=5, sticky=W)
    

def sme_three():
    if sme_three_checkbox.get():
        Label(page3, text="Select the Panel Member Titles", bg="lightgray", width=25).grid(row=1, column=5, sticky=W)
        Label(page3, text="SME", fg="red").grid(row=8, column=5, sticky=W)
    else:
        Label(page3, text="", bg="lightgray", width=4).grid(row=8, column=5, sticky=W)

def fmiddle():
    if fmiddle_checkbox.get():
        messagebox.showinfo("Initials", "Type First and Middle Initial")
        first_first_init_entry_box = Entry(page3, textvariable=ffint, bd=3,bg="red",width=5).grid(row=3, column=8, sticky=W, pady=4)


def smiddle():
    if smiddle_checkbox.get():
        messagebox.showinfo("Initials", "Type First and Middle Initial")
        second_first_init_entry_box = Entry(page3, textvariable=sfint, bd=3,bg="red",width=5).grid(row=5, column=8, sticky=W, pady=4)

def tmiddle():
    if tmiddle_checkbox.get():
        messagebox.showinfo("Initials", "Type First and Middle Initial")
        third_first_init_entry_box = Entry(page3, textvariable=tfint, bd=3,bg="red",width=5).grid(row=7, column=8, sticky=W, pady=4)

def fomiddle():
    if fomiddle_checkbox.get():
        messagebox.showinfo("Initials", "Type First and Middle Initial")
        forth_first_init_entry_box = Entry(page3, textvariable=fofint, bd=3,bg="red",width=5).grid(row=9, column=8, sticky=W, pady=4)


def show():
        numofquest = numofintvquestions.get()

        if numofquest == "6":
            Label(page4, text="Enter the Interview Question Six Points", fg="red").grid(row=14, column=4, sticky=W)
            quest_six=Entry(page4, textvariable=question_six, bd=3,bg="white",width=7).grid(row=15, column=4, sticky=W, pady=4)

        if InterviewYes_checkbox.get():
            Label(page4, text="Enter the Interview Question Two Points", fg="red").grid(row=6, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Three Points", fg="red").grid(row=8, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Four Points", fg="red").grid(row=10, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Five Points", fg="red").grid(row=12, column=4, sticky=W)
            quest_two=Entry(page4, textvariable=question_one, bd=3,bg="white",width=7).grid(row=7, column=4, sticky=W, pady=4)
            quest_three=Entry(page4, textvariable=question_one, bd=3,bg="white",width=7).grid(row=9, column=4, sticky=W, pady=4)
            quest_four=Entry(page4, textvariable=question_one, bd=3,bg="white",width=7).grid(row=11, column=4, sticky=W, pady=4)
            quest_five=Entry(page4, textvariable=question_one, bd=3,bg="white",width=7).grid(row=13, column=4, sticky=W, pady=4)
            

        if InterviewNo_checkbox.get():
            Label(page4, text="Enter the Interview Question Two Points", fg="red").grid(row=6, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Three Points", fg="red").grid(row=8, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Four Points", fg="red").grid(row=10, column=4, sticky=W)
            Label(page4, text="Enter the Interview Question Five Points", fg="red").grid(row=12, column=4, sticky=W)
            quest_two=Entry(page4, textvariable=question_two, bd=3,bg="white",width=7).grid(row=7, column=4, sticky=W, pady=4)
            quest_three=Entry(page4, textvariable=question_three, bd=3,bg="white",width=7).grid(row=9, column=4, sticky=W, pady=4)
            quest_four=Entry(page4, textvariable=question_four, bd=3,bg="white",width=7).grid(row=11, column=4, sticky=W, pady=4)
            quest_five=Entry(page4, textvariable=question_five, bd=3,bg="white",width=7).grid(row=13, column=4, sticky=W, pady=4)
            
def printworksheets():
    if survey_checkbox.get():
        messagebox.showinfo("Survey Spreadsheet", "The survey spreadsheets are being printed")
    if score_checkbox.get():
        messagebox.showinfo("Individual Score Spreadsheet", "The individual score spreadsheets are being printed")
    if indnotes_checkbox.get():
        messagebox.showinfo("Individual Notes Spreadsheet", "The individual notes spreadsheet are being printed")
    if intnotes_checkbox.get():
        messagebox.showinfo("Interview Notes Spreadsheet", "The interview notes spreadsheet are being printed")




root = tk.Tk()
root.title("Survey of Interest Application")
root.geometry("800x520+0+0")
root.configure(background='red')
nb = ttk.Notebook(root)

myjobtitle = StringVar()
myemail = StringVar()
myseries = StringVar()
myrpa = StringVar()
mycert = StringVar()
myduedate = StringVar()
mytime = StringVar()
myapplicant = StringVar()
survey_checkbox = IntVar()
score_checkbox = IntVar()
indnotes_checkbox = IntVar()
indnotes_checkbox = IntVar()
intnotes_checkbox = IntVar()
intnotes_checkbox = IntVar()
panel_member1 = StringVar()
panel_member2 = StringVar()
panel_member3 = StringVar()
panel_member4 = StringVar()
ffint = StringVar()
flint = StringVar()
sfint = StringVar()
slint = StringVar()
tfint = StringVar()
tlint = StringVar()
fofint = StringVar()
folint = StringVar()
maxscore = StringVar()
experience = StringVar()
education = StringVar()
training = StringVar()
awards = StringVar()
maxinterview = StringVar()
question_one = StringVar()
question_two = StringVar()
question_three = StringVar()
question_four = StringVar()
question_five = StringVar()
question_six = StringVar()
applicant = StringVar()
numofintvquestions = StringVar()





page1 = ttk.Frame(nb)
Label(page1, text="Enter the Job Title", fg="red").grid(row=1, column=3, sticky=W)
job_title_entry_box = Entry(page1, textvariable=myjobtitle, bd=3,bg="white",width=45).grid(row=2, column=3, sticky=W, pady=4)
Label(page1, text="Enter the Series", fg="red").grid(row=3, column=3, sticky=W)
series_entry_box = Entry(page1, textvariable=myseries, bd=3,bg="white",width=15).grid(row=4, column=3, sticky=W, pady=4)
Label(page1, text="Enter the RPA", fg="red").grid(row=5, column=3, sticky=W)
rpa_entry_box = Entry(page1, textvariable=myrpa, bd=3,bg="white",width=15).grid(row=6, column=3, sticky=W, pady=4)
Label(page1, text="Enter the Cert", fg="red").grid(row=7, column=3, sticky=W)
cert_entry_box = Entry(page1, textvariable=mycert, bd=3,bg="white",width=15).grid(row=8, column=3, sticky=W, pady=4)
Label(page1, text="Enter the Due Date", fg="red").grid(row=9, column=3, sticky=W)
due_date_entry_box = Entry(page1, textvariable=myduedate, bd=3,bg="white",width=15).grid(row=10, column=3, sticky=W, pady=4)
Label(page1, text="Enter the Time", fg="red").grid(row=11, column=3, sticky=W)
time_entry_box = Entry(page1, textvariable=mytime, bd=3,bg="white",width=15).grid(row=12, column=3, sticky=W, pady=4)

Label(page1, text="Email Addresses from PDF", fg="red").grid(row=1, column=1, sticky=W)
email_list_box=Listbox(page1, height=9, width=40)
email_list_box.grid(row=2, column=1, sticky=W, pady=4)
[email_list_box.insert(1, item) for item in (email_finder(pdf_content))]

Button(page1, text='Send Email', activebackground="white", bd=3,bg="white",width=28, command=createworksheets).grid(row=3, column=4, sticky=W, pady=4)
Button(page1, text='Delete Selected Email Addresses', activebackground="white", bd=3,bg="white",width=30, command=deleteselectedemails).grid(row=5, column=1, sticky=W, pady=4)
email_entry_box = Entry(page1, textvariable=myemail, bd=3,bg="white",width=35).grid(row=6, column=1, sticky=W, pady=4)
Button(page1, text='Add Email Addresses', activebackground="white", bd=3,bg="white",width=30, command=addemails).grid(row=7, column=1, sticky=W, pady=4)


page2 = ttk.Frame(nb)
Label(page2, text="Applicants", fg="black").grid(row=1, column=1, sticky=W)
Label(page2, text="Enter the Job Title", fg="red").grid(row=3, column=1, sticky=W)
job_title_entry_box = Entry(page2, textvariable=myjobtitle, bd=3,bg="white",width=45).grid(row=4, column=1, sticky=W, pady=4)
Label(page2, text="Enter the Series", fg="red").grid(row=5, column=1, sticky=W)
series_entry_box = Entry(page2, textvariable=myseries, bd=3,bg="white",width=15).grid(row=6, column=1, sticky=W, pady=4)
Label(page2, text="Enter the RPA", fg="red").grid(row=7, column=1, sticky=W)
rpa_entry_box = Entry(page2, textvariable=myrpa, bd=3,bg="white",width=15).grid(row=8, column=1, sticky=W, pady=4)
Label(page2, text="Enter the Cert", fg="red").grid(row=9, column=1, sticky=W)
cert_entry_box = Entry(page2, textvariable=mycert, bd=3,bg="white",width=15).grid(row=10, column=1, sticky=W, pady=4)
Checkbutton(page2, text="Further Considering (Declined)").grid(row=3, column=2,sticky=W)
Checkbutton(page2, text="Declined Interview").grid(row=3, column=3,sticky=W)
applicantlistbox=Listbox(page2,width=30,height=8,font=('times',13))
applicantlistbox.bind('<<ListboxSelect>>',printapplicants)
applicantlistbox.place(x=32,y=90)
#Lb1=Listbox(page2, height=8, width=40)
#yscroll = Scrollbar(page2, orient=VERTICAL)
#Lb1['yscrollcommand'] = yscroll.set
#yscroll['command'] = Lb1.yview
applicantlistbox.grid(row=2, column=1, sticky=W, pady=4)
#yscroll.grid(row=2, column=1, rowspan=2, sticky=N+S+E)
Button(page2, text='Import Applicants', activebackground="white", bd=3,fg="red",width=13, command=importapptextfile).grid(row=2, column=2, sticky=W, pady=4)
Button(page2, text='Delete All Applicants', activebackground="white", bd=3,bg="white",width=20, command=deleteallapplicants).grid(row=2, column=3, sticky=W, pady=4)
Button(page2, text='Delete Selected Applicants', activebackground="white", bd=3,bg="white",width=20, command=deleteselectedapplicants).grid(row=2, column=4, sticky=W, pady=4)
applicant_entry_box = Entry(page2, textvariable=myapplicant, bd=3,bg="white",width=35).grid(row=1, column=2, sticky=W, pady=4)
Button(page2, text='Add Applicant', activebackground="white", bd=3,bg="white",width=12, command=addapplicants).grid(row=1, column=3, sticky=W, pady=4)
Button(page2, text='Update', activebackground="white", bd=3,bg="white",width=12, command=createworksheets).grid(row=4, column=4, sticky=W, pady=4)

page3 = ttk.Frame(nb)

Label(page3, text="**Complete the Items in Red**", fg="red").grid(row=1, column=1, sticky=W)


#Lb13=Listbox(page3,height=6, width=40)
#Lb13.grid(row=3, column=1, sticky=W, pady=4)

Label(page3, text="Enter the # of Panel Members", fg="red").grid(row=9, column=1, sticky=W)
Spinbox1=Spinbox(page3, from_=3,to=4, width=5).grid(row=10, column=1, sticky=W, pady=4)
Label(page3, text="Select the Panel Member Titles", fg="red").grid(row=11, column=1, sticky=W)

panellistbox=Listbox(page3,width=21,height=5, font=('times',13))
panellistbox.bind('<<ListboxSelect>>',printpanel)
panellistbox.place(x=12,y=15)


    
chairperson_checkbox = IntVar()
Checkbutton(page3, text="ChairPerson", variable=chairperson_checkbox, command=chairperson).grid(row=12, column=1,sticky=W)
chairperson_eeo_checkbox = IntVar()
Checkbutton(page3, text="ChairPerson/EEO", variable=chairperson_eeo_checkbox, command=chairperson_eeo).grid(row=13, column=1,sticky=W)

sme_eeo_checkbox = IntVar()
Checkbutton(page3, text="SME/EEO", variable=sme_eeo_checkbox, command=sme_eeo).grid(row=14, column=1,sticky=W)

sme_one_checkbox = IntVar()
Checkbutton(page3, text="SME", variable=sme_one_checkbox, command=sme_one).grid(row=15, column=1,sticky=W)    

sme_two_checkbox = IntVar()
Checkbutton(page3, text="SME", variable=sme_two_checkbox, command=sme_two).grid(row=16, column=1,sticky=W)

sme_three_checkbox = IntVar()
Checkbutton(page3, text="SME", variable=sme_three_checkbox, command=sme_three).grid(row=17, column=1,sticky=W)

a = panel_member1
#fmiddle_checkbox = a[1:]
b = panel_member2
#smiddle_checkbox = b[1:]
c = panel_member3
#tmiddle_checkbox = c[1:]
d = panel_member4
#fomiddle_checkbox = d[1:]

    
fmiddle_checkbox = IntVar()
Checkbutton(page3, text="Include Middle Initial", variable=fmiddle_checkbox, command=fmiddle).grid(row=3, column=6,sticky=W)
smiddle_checkbox = IntVar()
Checkbutton(page3, text="Include Middle Initial", variable=smiddle_checkbox, command=smiddle).grid(row=5, column=6,sticky=W)
tmiddle_checkbox = IntVar()
Checkbutton(page3, text="Include Middle Initial", variable=tmiddle_checkbox, command=tmiddle).grid(row=7, column=6,sticky=W)
fomiddle_checkbox = IntVar()
Checkbutton(page3, text="Include Middle Initial", variable=fomiddle_checkbox, command=fomiddle).grid(row=9, column=6,sticky=W)


PanelCombo1=ttk.Combobox(page3, textvariable=panel_member1, width=25, postcommand=panel1update).grid(row=3, column=5, sticky=W, pady=4)
PanelCombo2=ttk.Combobox(page3, textvariable=panel_member2, width=25).grid(row=5, column=5, sticky=W, pady=4)
PanelCombo3=ttk.Combobox(page3, textvariable=panel_member3, width=25).grid(row=7, column=5, sticky=W, pady=4)
PanelCombo4=ttk.Combobox(page3, textvariable=panel_member4, width=25).grid(row=9, column=5, sticky=W, pady=4)




first_last_init_entry_box = Entry(page3, textvariable=flint, bd=3,bg="white",width=5).grid(row=3, column=11, sticky=W, pady=4)
Label(page3, text="+", bg="white",width=3).grid(row=3, column=10, sticky=W)


second_last_init_entry_box = Entry(page3, textvariable=slint, bd=3,bg="white",width=5).grid(row=5, column=11, sticky=W, pady=4)
Label(page3, text="+", bg="white",width=3).grid(row=5, column=10, sticky=W)

third_last_init_entry_box = Entry(page3, textvariable=tlint, bd=3,bg="white",width=5).grid(row=7, column=11, sticky=W, pady=4)
Label(page3, text="+", bg="white",width=3).grid(row=7, column=10, sticky=W)

forth_first_init_entry_box = Entry(page3, textvariable=fofint, bd=3,bg="white",width=5).grid(row=9, column=11, sticky=W, pady=4)
Label(page3, text="+", bg="white",width=3).grid(row=9, column=10, sticky=W)
forth_last_init_entry_box = Entry(page3, textvariable=folint, bd=3,bg="white",width=5).grid(row=9, column=11, sticky=W, pady=4)


Button(page3, text='Import Panel Members', activebackground="white", bd=3,fg="red",width=20, command=importpaneltextfile).grid(row=8, column=1, sticky=W, pady=4)
Button(page3, text='Delete Panel Members', activebackground="white", bd=3,fg="red",width=20, command=deleteallpanelmembers).grid(row=9, column=2, sticky=W, pady=4)




page4 = ttk.Frame(nb)
Label(page4, text="**Complete the Items in Red**", fg="red").grid(row=1, column=1, sticky=W)
Label(page4, text="Enter the Experience Points", fg="red").grid(row=2, column=1, sticky=W)
Entry(page4, textvariable=experience, bd=3,bg="white",width=7).grid(row=3, column=1, sticky=W, pady=4)
Label(page4, text="Enter the Education Points", fg="red").grid(row=4, column=1, sticky=W)
Entry(page4, textvariable=education, bd=3,bg="white",width=7).grid(row=5, column=1, sticky=W, pady=4)
Label(page4, text="Enter the Training Points", fg="red").grid(row=6, column=1, sticky=W)
Entry(page4, textvariable=training, bd=3,bg="white",width=7).grid(row=7, column=1, sticky=W, pady=4)
Label(page4, text="Enter the Awards Points", fg="red").grid(row=8, column=1, sticky=W)
Entry(page4, textvariable=awards, bd=3,bg="white",width=7).grid(row=9, column=1, sticky=W, pady=4)
Label(page4, text="Individual Max Points", bg="white").grid(row=10, column=1, sticky=W)
Entry(page4, textvariable=maxscore, bd=3,bg="white",width=7).grid(row=11, column=1, sticky=W, pady=4)

Label(page4, text="Enter the # of Interview Questions", fg="red").grid(row=2, column=4, sticky=W)
Spinbox2=Spinbox(page4, from_=5,to=6, textvariable=numofintvquestions, width=8).grid(row=3, column=4, sticky=W, pady=4)

Label(page4, text="Enter the Interview Question One Points", fg="red").grid(row=4, column=4, sticky=W)
quest_one=Entry(page4, textvariable=question_one, bd=3,bg="white",width=7).grid(row=5, column=4, sticky=W, pady=4)



Button(page4, text='Press Update', activebackground="white", bd=3,bg="white",width=10, command=show).grid(row=3, column=4, sticky=E, pady=1)

Label(page4, text="Are the Interview Points all the same?", fg="red").grid(row=2, column=5, sticky=W)   
InterviewYes_checkbox = IntVar()
Checkbutton(page4, text="Yes", variable=InterviewYes_checkbox, command=show).grid(row=3, column=5,sticky=W)
InterviewNo_checkbox = IntVar()
Checkbutton(page4, text="No", variable=InterviewNo_checkbox, command=show).grid(row=4, column=5,sticky=W)

Label(page4, text="Interview Max Points", fg="red").grid(row=18, column=4, sticky=W)
Entry1=Entry(page4, textvariable=maxinterview, bd=3,bg="white",width=7).grid(row=19, column=4, sticky=W, pady=4)


page5 = ttk.Frame(nb)
Button(page5, text='Create Survey Pages', activebackground="white", bd=3,bg="white",width=28).grid(row=3, column=1, sticky=W, pady=4)
Button(page5, text='Create Individual Worksheet Notes', activebackground="white", bd=3,bg="white",width=28).grid(row=4, column=1, sticky=W, pady=4)
Button(page5, text='Create Individual Score Worksheets', activebackground="white", bd=3,bg="white",width=28).grid(row=5, column=1, sticky=W, pady=4)
Button(page5, text='Create Individual Interview Pages', activebackground="white", bd=3,bg="white",width=28).grid(row=6, column=1, sticky=W, pady=4)
Button(page5, text='Update Page 2', activebackground="white", bd=3,bg="white",width=28, command=UpdatePage2).grid(row=7, column=1, sticky=W, pady=4)

survey_checkbox = IntVar()
Checkbutton(page5, text="Survey_Pages", variable=survey_checkbox).grid(row=9, column=1,sticky=W)
score_checkbox = IntVar()
Checkbutton(page5, text="Individual Score Sheets", variable=score_checkbox).grid(row=10, column=1,sticky=W)
indnotes_checkbox = IntVar()
Checkbutton(page5, text="Individual Notes Sheets", variable=indnotes_checkbox).grid(row=11, column=1,sticky=W)
intnotes_checkbox = IntVar()
Checkbutton(page5, text="Interview Notes Sheets", variable=intnotes_checkbox).grid(row=12, column=1,sticky=W)
Button(page5, text='Page Setup', activebackground="white", bd=3,bg="white",width=28).grid(row=13, column=1, sticky=W, pady=4)
Button(page5, text='Print', activebackground="white", bd=3,bg="white",width=28, command=printworksheets).grid(row=14, column=1, sticky=W, pady=4)




nb.add(page1, text='Email Addresses')
nb.add(page2, text='Job Title')
nb.add(page3, text='Panel Members')
nb.add(page4, text='Applicant/Interview Points')
nb.add(page5, text='Produce/Print Worksheets')

nb.pack(expand=1, fill="both")

root.mainloop()

from openpyxl import load_workbook
#import xlwings
import sys
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
import time

sys.path.append(".")

from read_document import read_document 

ws = Tk()
ws.title('Excel Parser')
ws.geometry('700x700') 

label_0=Label(ws,text="Insert the column value", width=20,font=("bold",20))
label_0.place(x=200,y=60)

coverage_label=Label(ws,text="Coverage", width=20,font=("bold",10))
coverage_label.place(x=180,y=130)

coverage_entry=Entry(ws)
coverage_entry.insert(END, 'Copertura')
coverage_entry.place(x=340,y=130)

impression_label=Label(ws,text="Impression", width=20,font=("bold",10))
impression_label.place(x=180,y=170)

impression_entry=Entry(ws)
impression_entry.insert(END, 'Impression')
impression_entry.place(x=340,y=170)

likes_label=Label(ws,text="Likes", width=20,font=("bold",10))
likes_label.place(x=180,y=210)

likes_entry=Entry(ws)
likes_entry.insert(END, 'Likes')
likes_entry.place(x=340,y=210)

hashtag_label=Label(ws,text="Hashtag", width=20,font=("bold",10))
hashtag_label.place(x=180,y=250)

hashtag_entry=Entry(ws)
hashtag_entry.insert(END, 'Hashtag')
hashtag_entry.place(x=340,y=250)

first_post_label=Label(ws,text="Row of the first post", width=20,font=("bold",10))
first_post_label.place(x=180,y=290)

first_post_entry=Entry(ws)
first_post_entry.insert(END, 10)
first_post_entry.place(x=340,y=290)

header_row_label=Label(ws,text="Header row", width=20,font=("bold",10))
header_row_label.place(x=180,y=330)

header_row_entry=Entry(ws)
header_row_entry.insert(END, 8)
header_row_entry.place(x=340,y=330)

excel_label = Label(ws, text='Upload Excel file', width=20,font=("bold",10))
excel_label.place(x=180, y=410)

excel_button = Button(ws, text ='Choose File and get column values', command = lambda:open_file()) 
excel_button.place(x=340, y=410)

common_hashtag_label = Label(ws, text = 'The commons hashtag are:', font=("bold",10))
common_hashtag_label.place(x=180, y=450)

highest_coverage_label = Label(ws, text = 'The post with the highest coverage is at:' ,font=("bold",10))
highest_coverage_label.place(x=180, y=490)

highest_impression_label = Label(ws, text = 'The post with the highest value of impression is at:', font=("bold",10))
highest_impression_label.place(x=180, y=530)

highest_likes_label = Label(ws, text = 'The post with the highest number of likes is at:', font=("bold",10))
highest_likes_label.place(x=180, y=570)

def get_args():
    args = {}
    args["hashtag_header"] = hashtag_entry.get()
    args["coverage_header"] = coverage_entry.get()
    args["like_header"] = likes_entry.get()
    args["impression_header"] = impression_entry.get()
    args["row_header"] = header_row_entry.get()
    args["first_post"] = first_post_entry.get() 
    return args

def open_file():
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")]) 
    if file_path != "":
        wb = load_workbook(file_path)
        sheets = wb.sheetnames
        sheet = wb[sheets[0]]
        args = get_args()
        s = read_document(sheet)
        s.check_row_used(args["first_post"])
        headers = s.find_header(args["coverage_header"], args["like_header"], args["row_header"], args["impression_header"], args["hashtag_header"])
        if "" in headers:
            highest_coverage_label["text"] = 'The post with the highest coverage is at: "Error, please review your headers!"'
            highest_impression_label["text"] = 'The post with the highest value of impression is at: "Error, please review your headers!"' 
            highest_likes_label["text"] = 'The post with the highest number of likes is at: "Error, please review your headers!"'
        else:
            rows = s.find_top_post()
            highest_coverage_label["text"] = "".join(['The post with the highest coverage is at: A', str(rows[0])])
            highest_impression_label["text"] = "".join(['The post with the highest value of impression is at: A', str(rows[1])])
            highest_likes_label["text"] = "".join(['The post with the highest number of likes is at: A', str(rows[2])])   
            common_hashtag_label["text"] = "".join(["The commons hashtag are: ", str(s.compare_hashtag())])
    else:
        pass


ws.mainloop()
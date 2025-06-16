import customtkinter as tk
from tkinter import filedialog 
#import tkinter
from PyPDF2 import PdfReader, PdfWriter,PdfMerger
from customtkinter import filedialog
from pdf2docx import Converter
#import fitz 
from PIL import Image, ImageTk
from tkinter import messagebox 
#import mysql.connector


tk.set_appearance_mode("dark")  # Modes: system (default), light, dark
tk.set_default_color_theme("green")  # Themes: blue (default), dark-blue, green


# app = tk.CTk()  #creating cutstom tkinter window
# app.geometry("750x500")
# app.title('Login')


# def button_function():
#     if username.get()=="" or pass_word.get()=="":
#         messagebox.showerror("Error","All feilds are required!")
#     else:
#         conn = mysql.connector.connect(host='localhost',user = 'root',password='Ajit@12345',database='mydata',port=3306)
#         cur=conn.cursor()
#         query=("select * from login where password=%s")
#         value = (pass_word.get(),)
#         cur.execute(query , value)
#         row = cur.fetchone()
#         if row!=None:
#             messagebox.showerror("Error","password already exist")
#         else :
#             cur.execute("insert into login values(%s,%s)",(var_username.get(),var_password.get()))
#         conn.commit()
#         conn.close()
#         messagebox.showinfo("info","Login seccesfully")
            
    
#     app.destroy()            # destroy current window and creating new one 
    # w = tk.CTk()  
    # w.geometry("1280x720")
    # w.title('Welcome')
    # l1=tk.CTkLabel(master=w, text="Home Page",font=('Century Gothic',60))
    # l1.place(relx=0.5, rely=0.5,  anchor=tkinter.CENTER)
    # w.mainloop()
    
    # tk.set_appearance_mode("light")  # Modes: system (default), light, dark
    # tk.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green
    
def select_file():
 file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
 entry.delete(0, tk.END)
 entry.insert(0, file_path)

# def compress_pdf():
    
#     file_path = entry.get()
#     output_path = file_path.replace('.pdf', '_compressed.pdf')
#     pdf_document = fitz.open(file_path)
    
#     # Create a new PDF with compressed images
#     new_pdf = fitz.open()
#     for page_number in range(len(pdf_document)):
#         page = pdf_document[page_number]
#         pix = page.get_pixmap()
#         new_pdf_page = new_pdf.new_page(width = page.rect.width, height = page.rect.height)
#         new_pdf_page.insert_image(page.rect, pixmap = pix)
    
#     # Save the new PDF
#     new_pdf.save(output_path)
#     pdf_document.close()
#     status_label.config(text='PDF compressed successfully!')

def cwindow():
    c = tk.CTk()
    tk.set_appearance_mode("dark")
    c.geometry("900x600")
    c.resizable(width=False , height=False)
    pdf = entry.get()
    if pdf == "":
        messagebox.showerror("error","Please select pdf")
    else: 
     convert_to_wordb = tk.CTkButton(c, text='convert to word',width=300 , height=70,fg_color="#f73d3d",hover_color='#c44848',font=("Century Gothic",24),bg_color="black",command=convert_to_word)
     convert_to_wordb.place(x=300,y=200)
    
    #  convert_to_txtb= tk.CTkButton(c, text='convert to txt',width=300 , height=70,fg_color="#f73d3d",hover_color='#c44848',font=("Century Gothic",24),bg_color="black",command=convert_to_txt)
    #  convert_to_txtb.place(x=500,y=100)
    
     c.mainloop()
     c.destroy()




def splitwindow():
    file_path = entry.get()
    if file_path=="" :
        messagebox.showerror("error","Please select pdf")
    else:
        
        s=tk.CTk()
        tk.set_appearance_mode("dark")
        s.geometry("900x600")
        s.resizable(width=False , height=False)
    
   #     Start Page
        def add_range_entry():
            range_entries.append(( tk.CTkEntry(s, width=35 , height=15),  tk.CTkEntry(s, width=35,height=15)))
            row = len(range_entries) + 3
            range_entries[-1][0].place(x=220, y=80+row*25 )
            range_entries[-1][1].place(x=280, y=80+row*25)
            add_button.place(x=350, y=80+row*25)
            remove_button.place(x=400, y=80+row*25)

        def remove_range_entry():
            if range_entries:
                entry1, entry2 = range_entries.pop()
                entry1.destroy()
                entry2.destroy()
                add_button.place_forget()
                remove_button.place_forget()
                for i, (e1, e2) in enumerate(range_entries):
                    row = i + 4
                    e1.place(x=220, y=80+row*25)
                    e2.place(x=280, y=80+row*25)
                    add_button.place(x=350, y=80+row*25)
                    remove_button.place(x=400, y=80+row*25)
        
        def split_pdf(input_path, output_path, ranges):
            try:
                pdf_file = open(input_path, 'rb')
                pdf_reader = PdfReader(pdf_file)
                if pdf_reader.is_encrypted:
                    messagebox.showerror("Error", "PDF is encrypted")
                else:
                    
                 for i, (start, end) in enumerate(ranges):
                        start, end = int(start), int(end)
                        if start < 1 or end > len(pdf_reader.pages):
                            messagebox.showerror("Error", f"Invalid range {start}-{end} for the PDF file")
                            return
                        
                        pdf_writer = PdfWriter()
                        for page_num in range(start - 1, end):
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                            
                        with open(f"{output_path}_part{i+1}.pdf", 'wb') as output_file:
                            pdf_writer.write(output_file)
                            
                 messagebox.showinfo("Success", "PDF has been successfully splitted")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
                
        def split():
            input_path = entry.get()
            output_path = entry.get()
            
            ranges = []
            for entry1, entry2 in range_entries:
                start = entry1.get()
                end = entry2.get()
                if start and end:
                    ranges.append((start, end))
            
            if not input_path or not output_path or not ranges:
                messagebox.showerror("Error", "Please fill all the fields")
                return
            
            split_pdf(input_path, output_path, ranges)
            
        
        # Create the main window
        ranges_label =  tk.CTkLabel(s, text="Page Ranges (start-end):" ,font=("Century Gothic",18))
        ranges_label.place(x=200, y=40)
        
        range_entries = []
        
        add_button =  tk.CTkButton(s, text="Add", command=add_range_entry,width=15 )
        add_button.place(x=350, y=80)
        
        remove_button =  tk.CTkButton(s, text="Remove", command=remove_range_entry , width=15)
        remove_button.place(x=400, y=80)
        
        sbutton = tk.CTkButton(s,text="split",command=split ,width=300 , height=70 , font=("Century Gothic",24))
        sbutton.place(x=300, y =350)
    
        s.mainloop()
    
    

    
def pdfencrypt():
    pdf_file = entry.get()
    if pdf_file=="" :
        messagebox.showerror("error","Please select pdf")
    
    else:
       def encrypt_pdf():
          
           password = password_entry.get()
           output_file = entry.get()
       
           writer = PdfWriter()
       
           try:
               reader = PdfReader(pdf_file)
               if reader.is_encrypted:
                    messagebox.showerror("Error", "PDF is already encrypted")
               else:
                for page in reader.pages:
                    writer.add_page(page)
        
                writer.encrypt(password)
                with open(output_file, 'wb') as f:
                    writer.write(f)
                
                messagebox.showinfo("Success", "PDF encrypted successfully!")
           except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
       e = tk.CTk()
       tk.set_appearance_mode("dark")
       e.geometry("900x600")
       e.resizable(width=False , height=False)
       tk.CTkLabel(e, text="Password Length:" , font=("Arial" , 19)).place(x = 250 , y = 150 )
       tk.CTkLabel(e, text="Enter Password:", font=("Arial" , 19)).place(x = 250 , y = 190)
       
       # Entries
       length_entry = tk.CTkEntry(e, width=30)
       length_entry.place(x = 410 , y = 150)
       length_entry.insert(0, "12")
       password_entry = tk.CTkEntry(e, width=100)
       password_entry.place(x = 400 , y = 190)
       
       # Buttons
       pass_button = tk.CTkButton(e, text='Encrypt',width=300 , height=70,fg_color="#f73d3d",hover_color='#c44848',font=("Century Gothic",24),bg_color="black",command=encrypt_pdf)
       pass_button.place(x=250,y=250)
       e.mainloop()

    
def convert_to_word():
    # Open PDF file
    # pdf_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_file_path = entry.get()
    if pdf_file_path:
        reader = PdfReader(pdf_file_path)
        if reader.is_encrypted:
                    messagebox.showerror("Error", "PDF is encrypted")
        else:
        # Specify output Word file path
            word_file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if word_file_path:
                # Convert PDF to Word
                cv = Converter(pdf_file_path)
                cv.convert(word_file_path, start=0, end=None)
                cv.close()
            messagebox.showinfo("info","PDF Conversion complete!")
            status_label.config(text="Conversion complete!")
   
    else:
        messagebox.showerror("error","Please select pdf")
        

    
            
def merge_pdfs():
    pdfs = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if pdfs:
        merger = PdfMerger()
        for pdf in pdfs:
                with open(pdf , "rb") as file:
                    reader = PdfReader(file)
                    if reader.is_encrypted:
                     messagebox.showerror("Error", "PDF is encrypted")
                     return
                    else: 
                     merger.append(pdf)
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
                merger.write(save_path)
                messagebox.showinfo("info", "PDF merge successfully!")
                merger.close()

# def convert_to_txt():
#  pdf_file_path = entry.get()
#  if pdf_file_path:
#     # Open PDF file
#     with open(pdf_file_path, 'rb') as pdf_file:
#         pdf_reader = PdfReader(pdf_file)
#         # Create text file name by replacing '.pdf' with '.txt'
#         txt_file_path = pdf_file_path[:-4] + ".txt"
#         # Create and write text content to the text file
#         with open(txt_file_path, 'w') as txt_file:
#             for page_num in range(len(pdf_reader.pages)):
#                 page = pdf_reader.pages[page_num]
#                 txt_file.write(page.extract_text())
#     messagebox.showinfo("info","PDF Conversion complete!")
#     status_label.config(text="Conversion complete!")
            



root = tk.CTk()

tk.set_appearance_mode("dark")
root.title('AHK PDF Manager')
root.geometry("900x600")
root.resizable(width=False , height=False)

background_image = Image.open("b1.PNG")
background_photo = ImageTk.PhotoImage(background_image)

canvas = tk.CTkCanvas(root, width=600, height=400)
canvas.pack(fill="both", expand=True)

canvas.create_image(0, 0, image=background_photo, anchor="nw")


background_image2 = Image.open("bg1.PNG")
background_photo2 = ImageTk.PhotoImage(background_image2)


f1 = tk.CTkFrame(canvas,width=200 , height=750)
f1.pack(side = "left" , fill = "y")
# frame = tk.CTkFrame(canvas,width=600 , height=300,fg_color="transparent" , bg_color="transparent")
# frame.pack(padx=30, pady=60)

label = tk.CTkLabel(f1, text="", image=background_photo2)
label.place(x=0, y=0, relwidth=1, relheight=1)

# l4 = tk.CTkLabel(frame , text="AHK PDF Manager ")
# l4.place(x=30 , y=20)
select_button = tk.CTkButton(canvas, text='Select PDF',width=300 , height=70,fg_color="#f73d3d",hover_color='#c44848',font=("Century Gothic",24),bg_color="black",command=select_file)
select_button.place(x=400,y=100)

entry = tk.CTkEntry(canvas, width=50 , bg_color="black")
entry.place(x=700,y=120)


split_button = tk.CTkButton(f1, text='Split PDF', fg_color="#363636" ,width=150 , height=20,font=("Arial",19),bg_color="#363636",command=splitwindow)
split_button.place(x=5,y=180)

convert_button = tk.CTkButton(f1, text='Convert', fg_color="#363636" ,width=150 , height=20,font=("Arial",19),bg_color="#363636",command=cwindow)
convert_button.place(x=5,y=140)

merge_button = tk.CTkButton(f1, text='Merge',fg_color="#363636", width=150 , height=20,font=("Arial",19), bg_color="#363636",command=merge_pdfs)
merge_button.place(x=5,y=220)

encrypt_button = tk.CTkButton(f1, text='Encrypt pdf',fg_color="#363636", width=150 , height=20,font=("Arial",19), bg_color="#363636",command=pdfencrypt)
encrypt_button.place(x=5,y=260)

# convert_button2 = tk.CTkButton(f1, text='Convert to txt', fg_color="#363636" ,width=150 , height=20,font=("Arial",19),bg_color="#363636",command=convert_to_txt)
# convert_button2.place(row=4, column=0, columnspan=2, pady=5)

toolslable = tk.CTkLabel(f1, text='Tools',fg_color="#363636", width=150 , height=20,font=("Arial",19), bg_color="#363636")
toolslable.place(x=5,y=60)

toolslable2 = tk.CTkLabel(f1, text='_______',fg_color="#363636", width=150 , height=20,font=("Arial",19), bg_color="#363636")
toolslable2.place(x=5,y=80)

status_label = tk.CTkLabel(root, text='')
status_label.pack()


root.mainloop()



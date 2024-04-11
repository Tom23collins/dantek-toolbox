from tkinter import *
from tkinter import filedialog
import customtkinter
from PIL import Image
from certificate_generator import generate_certificates
from pdf_extractor import extract_pdfs
import os

def menu():
    clear_frames()
    menu_frame = Frame(root_frame, width=600, height=400, bg='#174A7D')

    dantek_toolbox_logo_image = customtkinter.CTkImage(light_image=Image.open(f'{os.getcwd()}\\static\\dantek_toolbox_logo_no_border.png'), size=(160,92.24))
    dantek_toolbox_logo = customtkinter.CTkLabel(menu_frame, image=dantek_toolbox_logo_image, text='')
    dantek_toolbox_logo.pack(pady=20)

    welcome_text = customtkinter.CTkLabel(menu_frame, text='Welcome to the Dantek Toolbox\nA lightweight desktop application with some simple\n tools to help improve efficiency at Dantek.',text_color='white', fg_color='transparent', font=custom_font)
    welcome_text.pack()

    certificate_generator_icon = customtkinter.CTkImage(light_image=Image.open(f'{os.getcwd()}\\static\\certificate_generator_icon.png'), size=(60,50))
    certificate_generator_button = customtkinter.CTkButton(menu_frame, text='Certificate Generator', image=certificate_generator_icon, command=certificate_generator, width=150, height=100, fg_color='#152F62', font=custom_font, compound=TOP)
    certificate_generator_button.pack(side=LEFT, pady=20)

    pdf_extractor_icon = customtkinter.CTkImage(light_image=Image.open(f'{os.getcwd()}\\static\\pdf_extractor_icon.png'), size=(50,50))
    pdf_extractor_button = customtkinter.CTkButton(menu_frame, text='PDF Extractor', image=pdf_extractor_icon, command=pdf_extractor, width=150, height=100, fg_color='#152F62', font=custom_font, compound=TOP)
    pdf_extractor_button.pack(side=RIGHT, pady=20)

    menu_frame.pack(pady=20)
    
def pdf_extractor():
    def task():
        if folder_path != '':
            extract_pdfs(folder_path)

    def select_directory():
        global folder_path
        folder_path = filedialog.askdirectory(title='Select a directory')
        
    global folder_path
    folder_path = ''

    clear_frames()
    pdf_extractor_frame = Frame(root_frame, width=600, height=400, bg='#174A7D')

    pdf_extractor_image = customtkinter.CTkImage(light_image=Image.open(f'{os.getcwd()}\\static\\pdf_extractor_icon.png'), size=(50,50))
    pdf_extractor_icon = customtkinter.CTkLabel(pdf_extractor_frame, image=pdf_extractor_image, text='')
    pdf_extractor_icon.pack(pady=5)

    description = customtkinter.CTkLabel(pdf_extractor_frame, text='Extract all PDF Files from a \nsource containing multiple folders.',text_color='white', fg_color='transparent', font=custom_font)
    description.pack()

    how_to_instructions = customtkinter.CTkLabel(pdf_extractor_frame, text='How to use:\n1. Select the folder to extract all of the PDFs from\n2. Click extract PDFs\n3. A folder containing the PDF files will be \n    generated and saved to your desktop',text_color='white', fg_color='transparent', font=custom_font, justify='left')
    how_to_instructions.pack(pady=20)

    select_directory_button = customtkinter.CTkButton(pdf_extractor_frame, text='Select Folder', command=select_directory, width=200, height=40, corner_radius=50, fg_color='#152F62', font=custom_font)
    select_directory_button.pack(pady=5)

    extract_pdfs_button = customtkinter.CTkButton(pdf_extractor_frame, text='Extract PDFs', command=task, width=200, height=40, corner_radius=50, fg_color='#152F62', font=custom_font)
    extract_pdfs_button.pack()

    pdf_extractor_frame.pack(pady=20)

def certificate_generator():
    global file_path, training_course
    file_path = ''
    training_course = ''

    def select_data_callback():
        global file_path
        file_path = filedialog.askopenfilename(title='Select a file', filetypes=[('Microsoft Excel File', '*.xlsx')])
        select_data_button.configure(text=file_path.split('/')[-1])

    def select_course_callback(choice):
        global training_course
        training_course = choice

    def task():
        if file_path != '' and training_course != '':
            data_results_list = generate_certificates(training_course, file_path)
            results = customtkinter.CTk()
            results.title('')
            results.iconbitmap(f'{os.getcwd()}\\static\\dantek.ico')
            results.geometry(f'300x200+{int(root.winfo_screenwidth() / 2) - 300}+{int(root.winfo_screenheight() / 2) - 200}')
            results.minsize(300,200)
            results.wm_attributes('-transparentcolor', 'grey')
            results.config(background = '#174A7D')

            results_title = customtkinter.CTkLabel(results, text='Certificate Generation Results',text_color='white', fg_color='#174A7D', bg_color='#174A7D', font=custom_font)
            results_title.pack(pady=5)

            data = f'{data_results_list[0]} Delegates Processed\n{data_results_list[1]} Certificates Generated\n{data_results_list[2]} Emails Created\n{data_results_list[3]} Did not pass\n{data_results_list[4]} Had invalid emails'

            data_results = customtkinter.CTkLabel(results, text=data, text_color='white', fg_color='#174A7D', bg_color='#174A7D', font=custom_font, justify='left')
            data_results.pack()

            results.mainloop()

    clear_frames()
    certificate_generator_frame = Frame(root_frame, width=600, height=400, bg='#174A7D')

    certificate_generator_image = customtkinter.CTkImage(light_image=Image.open(f'{os.getcwd()}\\static\\certificate_generator_icon.png'), size=(60,50))
    certificate_generator_icon = customtkinter.CTkLabel(certificate_generator_frame, image=certificate_generator_image, text='')
    certificate_generator_icon.pack(pady=5)

    description = customtkinter.CTkLabel(certificate_generator_frame, text='Generate certificates for completion of a training\n course and then prepare an email to send to the delegate.',text_color='white', fg_color='transparent', font=custom_font)
    description.pack(pady=5)

    how_to_instructions = customtkinter.CTkLabel(certificate_generator_frame, text='How to use:\n1. Select the Microsoft forms data to use\n2. Click generate certificates\n3. The certificates will then be generated\n     and emailed to the delegates',text_color='white', fg_color='transparent', font=custom_font, justify='left')
    how_to_instructions.pack()

    select_data_button = customtkinter.CTkButton(certificate_generator_frame, text='Select Data', command=select_data_callback, width=300, height=40, corner_radius=50, fg_color='#152F62', font=custom_font)
    select_data_button.pack()

    select_course_dropdown = customtkinter.CTkOptionMenu(certificate_generator_frame, values=["Legionella Awareness - Log Book User", "Legionella Awareness - Responsible Persons"], command=select_course_callback, width=300, height=40, corner_radius=50, fg_color='#152F62', button_color='#152F62', font=custom_font, dynamic_resizing=False, anchor=CENTER)
    select_course_dropdown.set("Select Training Course")
    select_course_dropdown.pack(pady=5)

    generate_certificates_button = customtkinter.CTkButton(certificate_generator_frame, text='Generate Certificates', command=task, width=300, height=40, corner_radius=50, fg_color='#152F62', font=custom_font)
    generate_certificates_button.pack()

    certificate_generator_frame.pack(pady=10)

def clear_frames():
    for frame in root_frame.winfo_children():
        frame.destroy()
        
# Initialising root 
root = customtkinter.CTk()
customtkinter.set_appearance_mode('light')
root.title('')
root.geometry(f'600x400+{int(root.winfo_screenwidth() / 2) - 300}+{int(root.winfo_screenheight() / 2) - 200}')
root.iconbitmap(f'{os.getcwd()}\\static\\dantek.ico')
root.wm_attributes('-transparentcolor', 'grey')
root.minsize(600,400)
root.config(background = '#174A7D')

custom_font = customtkinter.CTkFont(family='Poppins',  size=12, weight='bold')

# Initialising menu
root_menu = Menu(root)
root_menu.add_command(label='Menu', command=menu, font=custom_font)
root_menu.add_command(label='Certificate Generator', command=certificate_generator)
root_menu.add_command(label='PDF Extractor', command=pdf_extractor)
root.config(menu=root_menu) 

# Initialising frames
root_frame = Frame(root, width=600, height=400, bg='#174A7D')
root_frame.pack()

# Load Menu Frame
menu()

# Start Loop
root.mainloop()

import tkinter as tk
import subprocess

# Create a new window with custom width and height
window = tk.Tk()
window.title("User Input Example")
window.geometry("1400x1000")
window.configure(background='black')

# Define a function to handle button hover events
def on_enter(e):
    e.widget['background'] = 'black'
    e.widget['foreground'] = 'white'

def on_leave(e):
    e.widget['background'] = 'white'
    e.widget['foreground'] = 'black'

# Create a label and Entry widget for the user to input their name
Heading = tk.Label(window, text="WELCOME TO", anchor='center', font=("Poppins", 42), background='black', fg='white')
Heading.pack(pady=100)

# Create a label for "BAS.AI" and set its foreground color to green
bas_ai_label = tk.Label(window, text="BAS.AI", anchor='center', font=("Poppins", 52), background='black', fg='green')
bas_ai_label.pack(pady=10)

# Define a function to hide the Heading after 3 seconds
def hide_heading():
    Heading.pack_forget()

# Use the after method to call the hide_heading function after 3 seconds
window.after(3000, hide_heading)

def makebody():
    def hide_everything():
        question.pack_forget()
        word_btn.pack_forget()
        excel_btn.pack_forget()
        ppt_btn.pack_forget()
        send_mail_btn.pack_forget()

    def make_word():
        window.after(100, hide_everything)
        heading_word = tk.Label(window, text='Word File', anchor="center", background='black', fg='white')
        heading_word.config(font=("Poppins", 42), background='black', fg='white')
        heading_word.pack()
        subprocess.Popen(["python", ""]) #Add your Word-Bas.AI path 
        window.iconify()

    def make_excel():
        window.after(100, hide_everything)
        heading_excel = tk.Label(window, text='Excel File', anchor="center", background='black', fg='white')
        heading_excel.config(font=("Poppins", 42), background='black', fg='white')
        heading_excel.pack()
        subprocess.Popen(["python", ""]) #Add your Excel-Bas.AI path 
        window.iconify()

    def make_ppt():
        window.after(100, hide_everything)
        heading_ppt = tk.Label(window, text='PPT File', anchor="center", background='black', fg='white')
        heading_ppt.config(font=("Poppins", 42), background='black', fg='white')
        subprocess.Popen(["python",""]) #Add your PPT-Bas.AI path 
        heading_ppt.pack()
        window.iconify()

    def send_mail():
        window.after(100, hide_everything)
        heading_mail = tk.Label(window, text='Send Mail', anchor="center", background='black', fg='white')
        heading_mail.config(font=("Poppins", 42), background='black', fg='white')
        heading_mail.pack()
        subprocess.Popen(["python", ""]) #Add your Send_Mail-Bas.AI path 
        window.iconify()

    question = tk.Label(window, text="What would you like to do?", anchor="center", background='black', fg='white')
    question.config(font=("Poppins", 42), background='black', fg='white')
    question.pack(pady=100)

    word_btn = tk.Button(window, text="Create a Word file", command=make_word, bg='black', fg='white', font=('Poppins', 20))
    excel_btn = tk.Button(window, text="Create an Excel file", command=make_excel, bg='black', fg='white', font=('Poppins', 20))
    ppt_btn = tk.Button(window, text="Create a PowerPoint file", command=make_ppt, bg='black', fg='white', font=('Poppins', 20))
    send_mail_btn = tk.Button(window, text="Send Mail", command=send_mail, bg='black', fg='white', font=('Poppins', 20))

    # Bind the hover functions to the buttons
    word_btn.bind("<Enter>", on_enter)
    word_btn.bind("<Leave>", on_leave)
    excel_btn.bind("<Enter>", on_enter)
    excel_btn.bind("<Leave>", on_leave)
    ppt_btn.bind("<Enter>", on_enter)
    ppt_btn.bind("<Leave>", on_leave)
    send_mail_btn.bind("<Enter>", on_enter)
    send_mail_btn.bind("<Leave>", on_leave)

    word_btn.pack(pady=10)
    excel_btn.pack(pady=10)
    ppt_btn.pack(pady=10)
    send_mail_btn.pack(pady=10)

window.after(4000, makebody)

# Start the main event loop
window.mainloop()
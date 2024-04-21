from tkinter import *
from tkinter import ttk, filedialog, messagebox
import customtkinter
import numpy
import pandas as pd


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")


#### Function to put the window on the center of the screen ####
def centerWindow(w, h):
    screenW = root.winfo_screenwidth()
    screenH = root.winfo_screenheight()

    x = (screenW/2) - (w/2)
    y = (screenH/2) - (h/2)

    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

#### Function to open the file ####
def openFile():
    File = filedialog.askopenfilename(title="Open File",
        filetype=(("Excel", ".xlsx"), ("All Files", "*.*")))
    

# Checks the file and returns exeptions if there are any #
    if File:
        try:
# Reads the file #
            df = pd.read_excel(File)
            #print(df) #debugging
# Exception if not an excel file #
        except Exception as e:
            messagebox.showerror("Woah!", f"There was a problem! {e}")
    else:
        messagebox.showwarning('Oops!', 'No file was selected! ')

# Removes anything from treeview # 
    Tree.delete(*Tree.get_children())

# Making a list from the excel file #
    Tree['column'] = list(df.columns)
    Tree['show'] = 'headings'

# Iterate thru the list to put headings into treeview #
    for col in Tree['column']:
        Tree.heading(col, text=col)

# df = dataframe
# Iterate thru the rows of the excel file and then put it into the treeview #
    Rows = df.to_numpy().tolist()
    for row in Rows:
        Tree.insert("", "end", values=row)    


#### Make the window ####
root = customtkinter.CTk()
root.title("Emin Hodzic <> Excel Opener")
root.iconbitmap("Assets//e.ico")
# Centering the window on the screen #
centerWindow(700, 400)


#### Treeview for Excel File ####
Tree = ttk.Treeview(root)
Tree.pack(pady=30)


# Headings height hack #
Tree.heading('#0', text="\n")

#### Style ####
style = ttk.Style()
style.theme_use("default")


# Configure the style of treeview, colors #
style.configure("Treeview", background="#270f36", foreground="white", rowheight=30, fieldbackground="#270f36")
style.map('Treeview', background=[('selected', "#000000")])

# Configure the style of treeview headings, colors #
style.configure("Treeview.Heading", background="#606060", foreground="black")


#### Search button ####
Button = customtkinter.CTkButton(root, text="Open Excel File", command=openFile)
Button.pack(pady=30)


root.mainloop()
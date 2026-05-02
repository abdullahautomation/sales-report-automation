import tkinter as tk
from tkinter import messagebox
from main_script import start
window = tk.Tk()
window.title("Sales Report Automation")
window.geometry("400x300")
label = tk.Label(window, text="Input File Name(Don't Add .xlsx in last)")
label.pack()
entry = tk.Entry(window, width= 40)
entry.pack()
label2 = tk.Label(window, text="Enter Output Name(Don't Add .xlsx in last)")
label2.pack()
entry2 = tk.Entry(window, width=40)
entry2.pack()
def generate():
    input_file = entry.get()
    output_file = entry2.get()
    start(input_file, output_file)
    tk.messagebox.showinfo("Done!", "Report Generated Successfully.")
button = tk.Button(window, text="Generate Report", command=generate)

button.pack()

window.mainloop()
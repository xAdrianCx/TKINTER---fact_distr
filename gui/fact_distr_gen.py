from tkinter import (Tk, ttk, messagebox, StringVar, Label, Entry, Button, filedialog)
from datetime import datetime


def get_excel():
	"""
	Gets the excel database needed to generate other files.
	:return:
	"""
	global input_file_path_var
	input_file_path_entry.delete(0, "end")
	excel_file = filedialog.askopenfilename(title="Select a file...", filetypes=[("Excel files", ".xlsx", ".xls")])
	input_file_path_var.set(excel_file)


def export_to():
	"""
	This function sets the folder where we want to generate the files.
	:return: Nothing
	"""
	global export_file_path_var
	export_file_path_entry.delete(0, "end")
	export_folder = filedialog.askdirectory(title="Choose directory...")
	export_file_path_var.set(export_folder)


# Set the main window.
root = Tk()
# Set main windows title.
root.title("Generator Facturi Distributie")
# Set main window width.
width = 450
# Set main window height
height = 350
# Set min size.
root.minsize(width, height)
# Set max size.
root.maxsize(width, height)

# Declare needed variables.
year_var = StringVar()
month_var = StringVar()
input_file_path_var = StringVar()
export_file_path_var = StringVar()

title = Label(root, text="GENERARE FACTURI DISTRIBUTIE", font=("Times New Roman", 15), bg="turquoise")
title.grid(row=0, columnspan=3, padx=30, pady=30)
title.grid_rowconfigure(1, weight=1)
title.grid_columnconfigure(1, weight=1)

# Create year label and grid it to screen.
year_label = Label(root, text="Select the year:", font=("Times New Roman", 12))
year_label.grid(row=1, column=0, pady=10, padx=10, sticky="W")
# Create year combobox.
year_combo = ttk.Combobox(root, textvariable=year_var, state="readonly")
year_combo["values"] = ("2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030",
								"2031", "2032", "2033", "2034", "2035", "2036", "2037", "2038", "2039", "2040",
								"2041", "2042", "2043", "2044", "2045", "2046", "2047", "2048", "2049", "2050")
year_combo.grid(row=1, column=1, pady=10, padx=10)
# Set current year as default value.
for i in range(len(year_combo["values"])):
	if datetime.today().year == int(year_combo["values"][i]):
		year_combo.current(i)

# Create month label and grid it to screen.
month_label = Label(root, text="Select the month:", font=("Times New Roman", 12))
month_label.grid(row=2, column=0, pady=10, padx=10, sticky="W")
month_combo = ttk.Combobox(root, textvariable=month_var, state="readonly")
month_combo["values"] = ("ianuarie", "februarie", "martie", "aprilie", "mai", "iunie", "iulie", "august", "septembrie",
						 "octombrie", "noiembrie", "decembrie")
month_combo.grid(row=2, column=1, pady=10, padx=10)
month_combo.current(datetime.today().month - 1)

# Create input file label, entrybox and askforfilepath button.
input_file_path_label = Label(root, text="Select excel file:", font=("Times New Roman", 12))
input_file_path_label.grid(row=3, column=0, pady=10, padx=10, sticky="W")
input_file_path_entry = Entry(root, textvariable=input_file_path_var, state="disabled")
input_file_path_entry.grid(row=3, column=1, pady=10, padx=10, sticky="W")
input_file_path_button = Button(root, text="Choose file...", font=("Times New Roman", 10), command=get_excel)
input_file_path_button.grid(row=3, column=2, pady=10, padx=10, sticky="W")

# Create export file labe, entry and askfordirectory button.
export_file_path_label = Label(root, text="Select export path:", font=("Times New Roman", 12))
export_file_path_label.grid(row=4, column=0, pady=10, padx=10, sticky="W")
export_file_path_entry = Entry(root, textvariable=export_file_path_var, state="disabled")
export_file_path_entry.grid(row=4, column=1, pady=10, padx=10, sticky="W")
export_file_path_button = Button(root, text="Choose directory...", font=("Times New Roman", 10), command=export_to)
export_file_path_button.grid(row=4, column=2, pady=10, padx=10, sticky="W")



# Run app.
root.mainloop()




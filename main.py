from tkinter import Tk, Button, Label, filedialog, Toplevel, Text, Scrollbar
from tkinter.ttk import Combobox
import run_excel
import threading
import sys
import io

class RedirectText(io.StringIO):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, s):
        self.text_widget.insert('end', s)
        self.text_widget.see('end')

def upload_file():
    global input_file
    input_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if input_file:
        label.config(text="File uploaded successfully!")

def generate_file():
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
    entity_name = entity_name_combobox.get()
    entity_code = entity_code_combobox.get()
    if input_file and output_file and entity_name and entity_code:
        # Show log window
        show_log_window()
        # Run the file processing in a separate thread
        threading.Thread(target=process_file, args=(input_file, output_file, 
                                                    entity_name, entity_code)).start()

def process_file(input_file, output_file, entity_name, entity_code):
    run_excel.generate_soa(input_file, output_file, entity_code, entity_name)
    label.config(text="File processed and saved successfully!")

def show_log_window():
    global log_window, log_text
    log_window = Toplevel(app)
    log_window.title("Processing Log")
    log_window.geometry("600x400")
    log_text = Text(log_window, wrap='word')
    log_text.pack(expand=True, fill='both')
    scrollbar = Scrollbar(log_text)
    scrollbar.pack(side='right', fill='y')
    log_text.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=log_text.yview)
    sys.stdout = RedirectText(log_text)

def update_entity_code(event):
    selected_name = entity_name_combobox.get()
    if selected_name in entity_mapping:
        entity_code_combobox.set(entity_mapping[selected_name])

def update_entity_name(event):
    entered_code = entity_code_combobox.get()
    for name, code in entity_mapping.items():
        if code == entered_code:
            entity_name_combobox.set(name)
            return
    entity_name_combobox.set("")




entity_mapping = {
    "ABC SDN BHD": "0001",
    "DEF SDN BHD": "0002",
    "GHI SDN BHD": "0003",
}


entity_names = list(entity_mapping.keys())
entity_codes = list(entity_mapping.values())

app = Tk()
app.title("SOA Generator")

app.geometry("500x400")

label = Label(app, text="Upload an Excel file to Generate SOA")
label.pack(pady=20)

upload_button = Button(app, text="Upload Excel File", command=upload_file)
upload_button.pack(pady=10)

entity_name_label = Label(app, text="Select Entity Name:")
entity_name_label.pack(pady=10)


entity_name_combobox = Combobox(app, values=entity_names, state="readonly", width= 40)
entity_name_combobox.pack(pady=10)
entity_name_combobox.bind("<<ComboboxSelected>>", update_entity_code)  

entity_code_label = Label(app, text="Enter or Select Company Code:")
entity_code_label.pack(pady=10)


entity_code_combobox = Combobox(app, values=entity_codes, state="normal")
entity_code_combobox.pack(pady=10)
entity_code_combobox.bind("<<ComboboxSelected>>", update_entity_name)
entity_code_combobox.bind("<KeyRelease>", update_entity_name)  

generate_button = Button(app, text="Generate File", command=generate_file)
generate_button.pack(pady=10)

app.mainloop()

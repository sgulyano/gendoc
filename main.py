import tkinter as tk
from tkinter import filedialog
import pandas as pd
from docx import Document
from docx.shared import Pt
from python_docx_replace import docx_replace, docx_get_keys

class DocumentFiller:
    def __init__(self):
        self.word_template_path = None
        self.excel_data_path = None
        self.save_path = None

    def upload_word_template(self):
        # Code to handle the uploaded Word template file
        self.word_template_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        word_template_label.config(text=f"Word Template: {self.word_template_path}")

    def upload_excel_data(self):
        # Code to handle the uploaded Excel data file
        self.excel_data_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        excel_data_label.config(text=f"Excel Data: {self.excel_data_path}")

        # Load the Excel data into a DataFrame
        df = pd.read_excel(self.excel_data_path)
        # Display the data in the excel table
        excel_data_table.config(text=f"Preview Excel Data: \n{df.iloc[:5,:5].to_string()}")

    def select_save_path(self):
        # Function to handle selecting the save path
        self.save_path = filedialog.asksaveasfilename()
        save_path_label.config(text=f"Save Path: {self.save_path}")

    def fill_word_template(self):
        # Validate that word_template_path and excel_data_path are not None
        if self.word_template_path is None or self.excel_data_path is None:
            tk.messagebox.showerror("Error", "Word template and Excel data must be uploaded")
            return
        
        # Validate that save_path is not None
        if self.save_path is None:
            tk.messagebox.showerror("Error", "Save path must be selected")
            return

        # Load the Excel data into a DataFrame
        df = pd.read_excel(self.excel_data_path)

        # Check if keys in documents match columns in DataFrame
        doc = Document(self.word_template_path)
        keys = docx_get_keys(doc)

        # count number of keys that are not in columns
        not_in_columns = len(set(keys) - set(df.columns))

        # Check if keys match columns in DataFrame
        if not_in_columns > 0:
            # Check if user wants to cancel execution
            confirm = tk.messagebox.showwarning("Warning", 
                                                f"{not_in_columns} keys in Word template do not match Columns in Excel data. Are you sure you want to continue?", 
                                                type=tk.messagebox.YESNO)
            if confirm == 'no':
                return

        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            # Convert the row to a dictionary
            row_dict = row.to_dict()

            # Replace nan with empty string in values of row_dict
            row_dict = {k: str(v) if pd.notnull(v) else "" for k, v in row_dict.items()}

            print(row_dict)

            # Load the Word template
            doc = Document(self.word_template_path)

            # Call the replace function with the row dictionary
            docx_replace(doc, **row_dict)

            # Save the filled Word template
            doc.save(self.save_path + "_" + str(index) + ".docx")

# Create DocumentFiller object
document_filler = DocumentFiller()

# Create the GUI
root = tk.Tk()
root.title("Generate Word Documents")

# Create frame for word template
word_template_frame = tk.Frame(root)
word_template_frame.pack()

# Create buttons for uploading Word template
word_template_button = tk.Button(word_template_frame, text="Upload Word Template", command=document_filler.upload_word_template)
word_template_button.pack(side=tk.LEFT)

# Create labels to display the uploaded file paths
word_template_label = tk.Label(word_template_frame, text="Word Template: ")
word_template_label.pack(side=tk.RIGHT)

# Create frame for excel data
excel_data_frame = tk.Frame(root)
excel_data_frame.pack()

# Create buttons for uploading Excel data
excel_data_button = tk.Button(excel_data_frame, text="Upload Excel Data", command=document_filler.upload_excel_data)
excel_data_button.pack(side=tk.LEFT)

# Create labels to display the uploaded file paths
excel_data_label = tk.Label(excel_data_frame, text="Excel Data: ")
excel_data_label.pack(side=tk.RIGHT)

# Create frame for excel data
excel_data_table_frame = tk.Frame(root)
excel_data_table_frame.pack()

# Create table to display data from Excel file
excel_data_table = tk.Label(excel_data_table_frame, text="Preview Excel Data: ", justify=tk.LEFT, wraplength=500)
excel_data_table.pack()

# Create frame for save path
save_path_frame = tk.Frame(root)
save_path_frame.pack()

# Create a button for selecting the save path
save_path_button = tk.Button(save_path_frame, text="Select Save Path", command=document_filler.select_save_path)
save_path_button.pack(side=tk.LEFT)

# Create a label to display the save path
save_path_label = tk.Label(save_path_frame, text="Save Path: ")
save_path_label.pack(side=tk.RIGHT)

# Create a button for filling the Word template
fill_button = tk.Button(root, text="Generate Word Document", command=document_filler.fill_word_template)
fill_button.pack()

# Run the GUI
root.mainloop()
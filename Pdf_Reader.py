# import module
import os
from pypdf import PdfReader
import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext
# import datetime as datetime

def Exctract(F_Name, output_text):
    pdf_Text = PdfReader(F_Name).pages[0].extract_text().replace("\n"," ")
    keyword = ["License Number:", "और पता:" ,"2.Address", "Issued On / िदनांक:", "Valid Upto: / वैधता:"]
    extracted_data = {}
    
    # Comapny Name
    if keyword[1] in pdf_Text:
        start_index = pdf_Text.find(keyword[1])
        x = pdf_Text[start_index + len(keyword[1]):pdf_Text.find(keyword[2])].split(" ")
        x = x[0:x.index("")]
        extracted_data["Company Name"] = " ".join(x)
        print("Company Name : ", extracted_data["Company Name"])
        output_text.insert(tk.END, "\nCompany Name : "+extracted_data["Company Name"])
        
    # License Number
    if keyword[0] in pdf_Text:
        start_index = pdf_Text.find(keyword[0])
        extracted_data["License Number"] = pdf_Text[start_index + len(keyword[0]):].split(" ")[0]
        print("License Number :", extracted_data["License Number"])
        output_text.insert(tk.END, "\nLicense Number : "+extracted_data["License Number"])
        
    # Issued On 
    if keyword[3] in pdf_Text:
        start_index = pdf_Text.find(keyword[3])
        extracted_data["Issued On"] = pdf_Text[start_index + len(keyword[3]):].split(" ")[0]
        print("Issued On :", extracted_data["Issued On"])
        output_text.insert(tk.END, "\nIssued On : "+extracted_data["Issued On"])
        
    # Valid upto
    if keyword[4] in pdf_Text:
        start_index = pdf_Text.find(keyword[4])
        extracted_data["Valid upto"] = (pdf_Text[start_index + len(keyword[4]):].split(" ")[0])
        print("Valid upto :",extracted_data["Valid upto"])
        output_text.insert(tk.END, "\nValid upto : "+extracted_data["Valid upto"])
        output_text.update() # Update the GUI
        
    return extracted_data


def process_files(output_text):
    directory = filedialog.askdirectory(title="Select directory containing PDF files")
    if not directory:
        return
    print('\nFolder path: ',directory)
    output_text.insert(tk.END, '\nFolder path: '+directory)
    pdf_file = [file for file in os.listdir(directory) if file.lower().endswith(".pdf")]
    print("Files in folder are :", pdf_file, "\nTotal ", len(pdf_file)+1, " files")
    output_text.insert(tk.END, "\nFiles in folder are :"+ str(pdf_file)+ "\nTotal "+ str(len(pdf_file)+1)+ " files")

    # Calling Extract funtion for data exctraction
    data = []
    for pdf_file in pdf_file:
        print(f"\nProcessing file: {pdf_file}")
        output_text.insert(tk.END, f"\n\nProcessing file: {pdf_file}")
        extracted_data = Exctract(os.path.join(directory, pdf_file), output_text)
        if extracted_data:
            data.append(extracted_data)
        else:
            print("No data found in this file")
            output_text.insert(tk.END, "\nNo data found in this file")
            output_text.update() # Update the GUI

    # saving in excel file
    if data:
        data = pd.DataFrame(data) # DataFrame created
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if output_file:
            data.to_excel(output_file, "PDF Data", index=False)
            print('Extracted data is available in ',output_file.split("/")[-1],'\n')
            output_text.insert(tk.END, '\nDataFrame is written to Excel File successfully.')
            output_text.update() # Update the GUI


# Create Tkinter window
root = tk.Tk()
root.title("PDF Data Extractor")

# Title label
title_label = tk.Label(root, text="PDF Data Extractor", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

# Created a text widget for displaying Output
output_text = scrolledtext.ScrolledText(root, width=45, height=20, font = ("Times New Roman", 15))
output_text.pack(padx=10, pady=10)

# Create a button to select the directory containing PDF files
process_button = tk.Button(root, text="Select Directory", command=lambda: process_files(output_text))
process_button.pack(pady=20)

root.mainloop()
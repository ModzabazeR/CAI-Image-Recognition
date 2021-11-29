#!/usr/bin/env python
# -*- coding: utf-8 -*-
import InvoiceExtract as ie
import tkinter as tk
from tkinter import filedialog
import os

processed = 0
skipped = 0




def process():
    for file in files:
        print(f"Processing {os.path.basename(file)} ...")
        if all(x in ie.ie_extract_text(file) for x in ie.KBANK_KEYWORDS):
            inv = ie.KBANKInvoice(file)
        elif all(x in ie.ie_extract_text(file) for x in ie.BBL_KEYWORDS):
            inv = ie.BBLInvoice(file)
        else:
            print("Unsupported Bank / Invoice Format")
            continue

        inv.get_invoice_info()
        inv.to_json()
        inv.close()

def locate():
    global files
    files = filedialog.askopenfilenames(
        initialdir="./", title=f"Select Invoice Files", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))

def gui():
        window = tk.Tk()
        window.title('File tool to JSON')
        window.minsize(width=720, height=360)
        title_label = tk.Label(master=window, text='\n What files do you want to convert\n',
                               font=("Arial", 28))
        title_label.pack()
        ins_button = tk.Button(master=window, text='Select', font=("Arial", 30), command=locate)
        ins_button.pack()
        process_button = tk.Button(master=window, text='Convert', font=("Arial", 30), command=process)
        process_button.pack()
        window.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    gui()

''' print(f"Processed {processed} Invoices.")
        print(f"Skipped {skipped} Invoices.")
        print("Done processing. DO NOT forget to check errors.")'''
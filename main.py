#!/usr/bin/env python
# -*- coding: utf-8 -*-
import InvoiceExtract as Ie
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
import logging
import utils

files = ()
scan_passed = 0
skipped = 0


def process():
    title_label.config(text="\nLoading...\n")
    scan_passed = 0
    skipped = 0
    for file in files:
        cur_file = os.path.basename(file)
        title_label.config(text=f"\nProcessing {cur_file} ...\n")
        print(f"Processing {cur_file} ...")
        logger.warning(f"Processing {cur_file} ...")
        if file.endswith(".pdf"):
            if all(x in utils.ie_extract_text(file) for x in Ie.KBANK_KEYWORDS):
                inv = Ie.KBANKInvoice(file)
            elif all(x in utils.ie_extract_text(file) for x in Ie.BBL_KEYWORDS):
                inv = Ie.BBLInvoice(file)
            elif all(x in utils.ie_extract_text(file) for x in Ie.SCB_KEYWORDS):
                inv = Ie.SCBInvoice(file)
            else:
                logger.warning(f"Unsupported Bank / Invoice Format : {cur_file}")
                skipped += 1
                continue
            try:
                inv.get_invoice_info()
                inv.to_excel()
            except:
                logging.exception("Error occurred while processing")
            finally:
                inv.close()

        scan_passed += 1
        # logger.warning(f'"{cur_file}.xlsx" written to output folder.')
        title_label.config(text=f"\nProcessed {scan_passed} Invoices.\n")

    if scan_passed + skipped == len(files):
        os.system("start output")
        messagebox.showinfo(title=None, message=f'Extract complete {scan_passed} file \n Skipped {skipped} file')
        if os.path.exists(r"output/temp"):
            os.rmdir(r"output/temp")
        # logger.warning(f"Processed {scan_passed} Invoices.")
        # logger.warning(f"Skipped {skipped} Invoices.")
        # logger.warning("Done processing. DO NOT forget to check errors.")

def start_submit_thread():
    global files, scan_passed, skipped
    if len(files) != 0 :
        ask = messagebox.askquestion(title=None, message='Do you wish to process?')
        if ask == 'yes':
            scan_passed = 0
            skipped = 0
            process_thread = threading.Thread(target=process)
            process_thread.daemon = True
            process_thread.start()
        else:
            files = ()
            files_list.config(text="Your file : \n None", font=("Arial", 9))
    else:
        messagebox.showerror(title='Error', message='Select file to convert')

def locate():
    global files, files_list
    files_list = tk.Label (master=root, font=("Arial", 9))
    files = filedialog.askopenfilenames(
        initialdir="./", title=f"Select Invoice Files", filetypes=(('PDF files','*.pdf'), ("All files", "*.*")))
    str= '\n'.join(files)
    if len(files) > 0:
        files_list.config(text=f'Your file : \n {str}', font=("Arial", 9))
    else:
        files_list.config(text="Your file : \n None", font=("Arial", 9))

    files_list.grid(column=0, row=2, columnspan=3, padx=5, pady=5)
    
    if not os.path.exists(r"output"):
        os.mkdir(r"output")

def run ():
    global files, scan_passed, skipped
    if len(files) != 0 :
        ask = messagebox.askquestion(title=None, message='Do you wish to process?')
        if ask == 'yes':
            scan_passed = 0
            skipped = 0
            files = ()
            files_list.config(text="Your file : \n Completed", font=("Arial", 9))
        else:
            files = ()
            files_list.config(text="Your file : \n None", font=("Arial", 9))
    else:
        messagebox.showerror(title='Error', message='Select file to convert')

if __name__ == "__main__":
    logging.basicConfig(filename="log.txt", filemode="w", format="%(asctime)s - %(message)s", level=logging.WARNING)
    logger = logging.getLogger()
    root = tk.Tk()
    root.title('Doc Juice')
    root.geometry("480x360")
    root.iconbitmap('icon.ico')

    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)
    root.rowconfigure(2, weight=1)

    title_label = tk.Label(master=root, text='Doc Juice!', font=("Arial", 28))
    title_label.grid(column=0, row=0, padx=5, pady=5, columnspan=2)

    description_label = tk.Label(master=root,
                        text='Select your invoice to convert to excel \nFile support : PDF',
                        font=("Arial", 12))
    description_label.grid(column=0, row=1, padx=5, pady=5, columnspan=2)

    insert_file = tk.Button(master=root, text='Choose file', font=("Arial", 10), width=20, command=locate)
    insert_file.grid(column=0, row=3, padx=5, pady=5)

    extract_file = tk.Button(master=root, text='Extract file', font=("Arial", 10), width=20, command=start_submit_thread)
    extract_file.grid(column=1, row=3, padx=5, pady=5)

    author_info = tk.Label(master=root, text = 'Build by CAI-C Gen 4 Doc Juice! Team')
    author_info.grid(column=1, row=4, padx=5, pady=5)

    root.mainloop()

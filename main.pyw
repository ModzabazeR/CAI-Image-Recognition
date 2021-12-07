#!/usr/bin/env python
# -*- coding: utf-8 -*-
import InvoiceExtract as Ie
import tkinter as tk
from tkinter import filedialog
import os
import threading
import logging


def process():
    title_label.config(text="\nLoading...\n")
    global processed, skipped
    processed = 0
    skipped = 0
    for file in files:
        cur_file = os.path.basename(file)
        title_label.config(text=f"\nProcessing {cur_file} ...\n")
        logger.warning(f"Processing {cur_file} ...")
        if all(x in Ie.ie_extract_text(file) for x in Ie.KBANK_KEYWORDS):
            inv = Ie.KBANKInvoice(file)
        elif all(x in Ie.ie_extract_text(file) for x in Ie.BBL_KEYWORDS):
            inv = Ie.BBLInvoice(file)
        else:
            logger.warning("Unsupported Bank / Invoice Format")
            skipped += 1
            continue

        try:
            inv.get_invoice_info()
            inv.to_excel()
        except:
            logging.exception("Error occurred while processing")
        finally:
            inv.close()
        processed += 1
        logger.warning(f'"{cur_file}.xlsx" written to output folder.\n')
        title_label.config(text=f"\nProcessed {processed} Invoices.\n")

    if processed + skipped == len(files):
        os.system("start output")
        if os.path.exists(r"output/temp"):
            os.rmdir(r"output/temp")
        logger.warning(f"Processed {processed} Invoices.")
        logger.warning(f"Skipped {skipped} Invoices.")
        logger.warning("Done processing. DO NOT forget to check errors.")


def start_submit_thread():
    process_thread = threading.Thread(target=process)
    process_thread.daemon = True
    process_thread.start()


def locate():
    global files
    files = filedialog.askopenfilenames(
        initialdir="./", title=f"Select Invoice Files", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
    if len(files) > 0:
        process_button.config(state="normal")
    title_label.config(text=f"\nSelected {len(files)} files.\n")


if __name__ == "__main__":
    logging.basicConfig(filename="log.txt", filemode="w", format="%(asctime)s - %(message)s", level=logging.WARNING)
    logger = logging.getLogger()
    window = tk.Tk()
    window.title('Invoice2data')
    window.minsize(width=720, height=360)
    title_label = tk.Label(master=window, text='\n What files do you want to convert\n',
                           font=("Arial", 28))
    title_label.pack()
    ins_button = tk.Button(master=window, text='Select', font=("Arial", 30), command=locate)
    ins_button.pack()
    process_button = tk.Button(master=window, text='Convert', font=("Arial", 30), state="disabled",
                               command=start_submit_thread)
    process_button.pack()
    window.mainloop()

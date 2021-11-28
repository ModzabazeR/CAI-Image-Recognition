#!/usr/bin/env python
# -*- coding: utf-8 -*-
import InvoiceExtract as ie
import tkinter as tk
from tkinter import filedialog
import os


def run():
    files = filedialog.askopenfilenames(
        initialdir="./", title=f"Select Invoice Files", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))

    processed = 0
    skipped = 0

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

    print(f"Processed {processed} Invoices.")
    print(f"Skipped {skipped} Invoices.")
    print("Done processing. DO NOT forget to check errors.")
    input("Press Enter to exit...")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    run()
import pdfplumber
import os
import re
import json
import pandas as pd
from collections import OrderedDict

BBL_KEYWORDS = ("Currency THB Date", "By the instruction of", "Beneficiary name :", "Beneficiary Account :",
                "Invoice details as follows (if any)", "Payment Net")
KBANK_KEYWORDS = ("KASIKORNBANK PCL", "On behalf of", "Payment details are as follows")
MAPPING = json.load(open("mapping.json", "r", encoding='utf-8'))


def pretty_save_json(file: str, data: dict) -> None:
    with open(file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def ie_extract_text(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

class PDFInvoice:
    # Invoice metadata
    pdf = None
    invoice_types = ['Credit Advice Report',
                     'Pre-Advice Report', 'Payment Advice']
    invoice_extension = ''
    invoice_type = ''
    text = ''

    # Invoice data
    receiver = None
    bank = None
    receiver_account = None
    payment_date = None
    total_after_tax = None
    cheque_id = None
    sender = None
    bank_charge = None

    def close(self) -> None:
        self.pdf.close()

    def get_invoice_info(self) -> None:
        pass

    def extract(self, file: str) -> None:
        pass

    def get_entries(self) -> dict:
        data = {
            "บริษัท (ผู้รับเงิน)": self.receiver,
            "ธนาคาร": self.bank,
            "เลขที่บัญชีธนาคารบริษัทที่โอนเข้า": self.receiver_account,
            "วันที่ชำระ": self.payment_date,
            "จำนวนเงินที่ชำระ": self.total_after_tax,
            "เลขที่เช็ค": self.cheque_id,
            "ชื่อลูกค้า": self.sender,
            "items": self.extract(self.file_path),
            "ค่าธรรมเนียมธนาคาร (ถ้ามี)": self.bank_charge
        }
        return data

    def __init__(self, file_path: str) -> None:
        if not os.path.exists('output'):
            os.makedirs('output')

        self.file_path = file_path
        if file_path.endswith('.pdf'):
            self.invoice_extension = 'pdf'
            self.pdf = pdfplumber.open(file_path)
            self.get_invoice_info()
        else:
            raise Exception('File type not supported')

    def to_txt(self) -> None:
        "Extract raw text from the pdf file"
        with open(f'output/{os.path.basename(self.file_path)}.txt', 'w', encoding='utf-8') as f:
            for page in self.info:
                f.write(f'[PAGE {page.page_number}]\n') if page.page_number == 1 else f.write(
                    f'\n[PAGE {page.page_number}]\n')
                f.write(page.extract_text())
        print(f'"{os.path.basename(self.file_path)}.txt" written to output folder')

    def to_json(self) -> None:
        data = self.get_entries()
        pretty_save_json(
            f'output/{os.path.basename(self.file_path)}.json', data)

        print(f'"{os.path.basename(self.file_path)}.json" written to output folder\n')

    def to_excel(self) -> None:
        pd.DataFrame.from_dict(self.get_entries(), orient='columns').to_excel(
            f'output/{os.path.basename(self.file_path)}.xlsx', index=False)
        print(f'"{os.path.basename(self.file_path)}.xlsx" written to output folder')


class KBANKInvoice(PDFInvoice):
    def correct_words(self, text: str, mapping: dict) -> str:
        for word in mapping:
            text = text.replace(word, mapping[word])
        return text

    def parse_row(self, row):
        return OrderedDict([
            ("INV.NUMBER", row[0:18].strip()),
            ("INV.DATE", row[18:32].strip()),
            ("INV.AMOUNT", row[32:45].strip()),
            ("VAT AMT", row[45:55].strip()),
            ("Amt. (รวม Vat)", ""),
            ("WHT AMT", row[55:65].strip()),
            ("NET AMOUNT", row[65:].strip())
        ])

    def extract(self, path: str) -> dict:
        pdf = pdfplumber.open(path)
        p0 = pdf.pages[1]
        text = p0.extract_text()
        core_pat = re.compile(r"NET AMOUNT\n=+\n(.*)\n=+\nTOTAL", re.DOTALL)
        core = re.search(core_pat, text).group(1)
        core = core.split("\n")

        parsed = [self.parse_row(x) for x in core]
        cols = list(parsed[0].keys())
        data = pd.DataFrame(parsed, columns=cols)
        data = data.drop(["INV.DATE"], 1)
        data = data.rename(columns={"INV.NUMBER": "เลขที่ Invoice", "INV.AMOUNT": "Amt. (ก่อน Vat)",
                                    "VAT AMT": "Vat. Amt", "WHT AMT": "WHT Amt. (แต่ละ Inv)", "NET AMOUNT": "จำนวนเงินสุทธิ (แต่ละ Inv)"})
        data = data.replace([''], [None])
        data_dict = data.to_dict(orient="records")
        return data_dict

    def get_invoice_info(self) -> None:
        self.info = self.pdf.pages

        for page in self.info:
            self.text += page.extract_text()
        
        self.text = self.correct_words(self.text, MAPPING)

        type_match = re.search(r"(Subject : )(\w)+", self.text)
        if type_match:
            self.invoice_type = type_match.group(2)

        date_match = re.search(
            r"(Cheque Date : )(\d{2}\/\d{2}/\d{4})", self.text)
        if date_match:
            self.payment_date = date_match.group(2)
        else:
            print(
                f'Payment date not found in {os.path.basename(self.file_path)}')

        sender_match = re.search(
            r"(Payer Name\s+: )([\w ()ก-๛.,]+)", self.text)
        if sender_match:
            self.sender = sender_match.group(2)
        else:
            print(
                f'Sender name not found in {os.path.basename(self.file_path)}')

        receiver_match = re.search(r"(To : )([\w ()ก-๛.,]+)", self.text)
        if receiver_match:
            self.receiver = receiver_match.group(2)
        else:
            print(
                f'Receiver name not found in {os.path.basename(self.file_path)}')

        total_match = re.search(
            r"(Total Invoice after VAT : \*+)([\d,.]+)", self.text)
        if total_match:
            self.total_after_tax = total_match.group(2)
        else:
            print(
                f'Total after tax not found in {os.path.basename(self.file_path)}')

        bank_charge_match = re.search(
            r"(Benef Charges : \*+)([\d,.]+)", self.text)
        if bank_charge_match:
            self.bank_charge = bank_charge_match.group(2)
            self.bank_charge = self.bank_charge.replace(".00", "0")

    def __init__(self, file_path: str) -> None:
        super().__init__(file_path)
        self.bank = 'กสิกรไทย (KBANK)'


class BBLInvoice(PDFInvoice):
    def get_invoice_info(self) -> None:
        self.info = self.pdf.pages

        for page in self.info:
            self.text += page.extract_text()

        for type in self.invoice_types:
            if type in self.text:
                self.invoice_type = type
                break

        if self.invoice_type not in self.invoice_types:
            raise ValueError(
                f'Invoice type not found in {os.path.basename(self.file_path)}')

        date_match = re.search(
            r'(Payment Date : )(\d{2}-\w{3}-\d{2})', self.text)
        if date_match:
            self.payment_date = date_match.group(2)
        else:
            print(
                f'Payment Date not found in {os.path.basename(self.file_path)}')

        sender_match = re.search(
            r'(By the instruction of : )([\w ()ก-๛.,]+)', self.text)
        if sender_match:
            self.sender = sender_match.group(2)
        else:
            print(f'Sender not found in {os.path.basename(self.file_path)}')

        receiver_match = re.search(
            r'(Beneficiary name : )([\w ()ก-๛.,]+)', self.text)
        if receiver_match:
            self.receiver = receiver_match.group(2)
        else:
            print(f'Receiver not found in {os.path.basename(self.file_path)}')

        receiver_account_match = re.search(
            r'(Beneficiary [Aa]ccount : )(\d+)', self.text)
        if receiver_account_match:
            self.receiver_account = receiver_account_match.group(2)

        total_match = re.search(r'(Payment Net : )([\d,.]+)', self.text)
        if total_match:
            self.total_after_tax = total_match.group(2)

        cheque_id_match = re.search(r'(Cheque No. : )(\d+)', self.text)
        if cheque_id_match:
            self.cheque_id = cheque_id_match.group(2)

    def __init__(self, file_path: str) -> None:
        super().__init__(file_path)
        self.bank = 'กรุงเทพ (BBL)'

    def extract(self, file: str) -> dict:
        pdf = pdfplumber.open(file)

        credit_advice_cols = ["Item No", "Invoice No.", "Date",
                              "Gross Amount", "WHT Amount", "VAT Amount", "Income Type"]
        pre_advice_cols = ["Item No", "Invoice No.",
                           "Date", "Gross Amount", "WHT Amount"]

        cols = credit_advice_cols if self.invoice_type == 'Credit Advice Report' else pre_advice_cols

        data = pd.DataFrame(columns=cols)

        for page in pdf.pages:
            table = page.extract_table()
            if type(table) != list:
                return None

            # Filter out empty rows
            table = list(filter(lambda a: a != ['', '', '', '', '', '', ''] and a != [
                         '', '', '', '', ''], table))

            # replace \n with space in the table
            for i in range(len(table)):
                for j in range(len(table[i])):
                    table[i][j] = table[i][j].replace('\n', ' ')

            if page.page_number == 1 and self.invoice_type == 'Credit Advice Report':
                df = pd.DataFrame(table[0:-1], columns=cols)
            else:
                df = pd.DataFrame(
                    table[1:-1], columns=cols) if self.invoice_type == 'Credit Advice Report' else pd.DataFrame(table, columns=cols)

            data = pd.concat([data, df], ignore_index=True)

        data = data.drop(0)
        data = data[["Invoice No.", "Gross Amount", "WHT Amount"]]
        placeholder = [None for i in range(len(data))]
        data.insert(1, "Amt. (ก่อน Vat)", placeholder)
        data.insert(2, "Vat. Amt", placeholder)
        data["จำนวนเงินสุทธิ (แต่ละ Inv)"] = placeholder

        data = data.rename(columns={"Invoice No.": "เลขที่ Invoice",
                                    "Gross Amount": "Amt. (รวม Vat)", "WHT Amount": "WHT Amt. (แต่ละ Inv)"})
        data = data.replace([''], [None])
        
        # if all the invoice data are None, return None
        if all(data.isna().all()):
            return None

        data_dict = data.to_dict(orient="records")
        return data_dict

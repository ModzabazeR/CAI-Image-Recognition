import pdfplumber
import os
import re
import json
import csv
import openpyxl as xl
from openpyxl.styles import PatternFill, Alignment, Border, Side
import pandas as pd
from collections import OrderedDict
import utils
from google.cloud import vision
import tempfile

def create_temp_key(api_key: dict):
    "Create temp key file"
    f = tempfile.NamedTemporaryFile(mode="w+", delete=False)
    json.dump(api_key, f)
    f.flush()
    os.system("attrib +h " + f.name)
    return f

def authorize_google_vision(api_key_file: tempfile._TemporaryFileWrapper):
    "Authorize Google Vision API"
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = api_key_file.name

def delete_temp_key(api_key_file):
    "Delete temp key file"
    try:
        api_key_file.close()
        os.unlink(api_key_file.name)
    except FileNotFoundError:
        print(f"File {api_key_file.name} not found")

API_KEY = # YOUR API KEY HERE

f = create_temp_key(API_KEY)
authorize_google_vision(f)
client = vision.ImageAnnotatorClient()

BULLETIN_KEYWORDS = ("Bulletin Co.,Ltd", " ")
METRO_KEYWORDS = ("เมโทร ยูนิฟอร์ม", " ")
LEEKA_KEYWORDS = ("ลีก้า", "LEEKA", " ")

BBL_KEYWORDS = ("Currency THB Date", "By the instruction of", "Beneficiary name :", "Beneficiary Account :",
                "Invoice details as follows (if any)", "Payment Net")
KBANK_KEYWORDS = ("KASIKORNBANK PCL", "On behalf of", "Payment details are as follows")
MAPPING = json.load(open("mapping.json", "r", encoding='utf-8'))


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

    def extract(self, file: str, mode: str = "records") -> None:
        pass

    def get_entries(self, mode: str = "records") -> dict:
        data = {
            "บริษัท (ผู้รับเงิน)": self.receiver,
            "ธนาคาร": self.bank,
            "เลขที่บัญชีธนาคารบริษัทที่โอนเข้า": self.receiver_account,
            "วันที่ชำระ": self.payment_date,
            "จำนวนเงินที่ชำระ": self.total_after_tax,
            "เลขที่เช็ค": self.cheque_id,
            "ชื่อลูกค้า": self.sender,
            "items": self.extract(self.file_path, mode),
            "ค่าธรรมเนียมธนาคาร (ถ้ามี)": self.bank_charge
        }
        return data

    def __init__(self, file_path: str) -> None:
        self.info = None

        self.file_path = file_path
        if file_path.endswith('.pdf'):
            self.invoice_extension = 'pdf'
            self.pdf = pdfplumber.open(file_path)
            self.get_invoice_info()
        else:
            raise Exception('File type not supported')

    def to_txt(self) -> None:
        with open(f'output/{os.path.basename(self.file_path)}.txt', 'w', encoding='utf-8') as f:
            for page in self.info:
                f.write(f'[PAGE {page.page_number}]\n') if page.page_number == 1 else f.write(
                    f'\n[PAGE {page.page_number}]\n')
                f.write(page.extract_text())
        print(f'"{os.path.basename(self.file_path)}.txt" written to output folder')

    def to_json(self) -> None:
        data = self.get_entries()
        utils.pretty_save_json(
            f'output/{os.path.basename(self.file_path)}.json', data)

        print(f'"{os.path.basename(self.file_path)}.json" written to output folder\n')

    def to_excel(self) -> None:
        if not os.path.exists(r'output/temp'):
            os.mkdir(r'output/temp')

        file = self.get_entries(mode="list")

        metadata = {k: v for k, v in file.items() if k != "items" and k !=
                    "ค่าธรรมเนียมธนาคาร (ถ้ามี)"}
        items = file["items"]
        fee = file["ค่าธรรมเนียมธนาคาร (ถ้ามี)"]
        header = [i for i in metadata.keys()] + [i for i in items.keys()] + \
                 ["ค่าธรรมเนียมธนาคาร (ถ้ามี)"]

        # to csv
        with open(f"output/temp/{os.path.basename(self.file_path)}.csv", "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            if len(items["เลขที่ Invoice"]) > 0:
                for i in range(len(items["เลขที่ Invoice"])):
                    if i == 0:
                        writer.writerow([metadata[k] for k in metadata.keys()] +
                                        [items[k][i] for k in items.keys()] + [fee])
                    else:
                        writer.writerow([None for _ in metadata.keys()] +
                                        [items[k][i] for k in items.keys()])
            else:
                writer.writerow([metadata[k] for k in metadata.keys()] +
                                [None for _ in items.keys()] + [fee])

        df = pd.read_csv(f"output/temp/{os.path.basename(self.file_path)}.csv", encoding="utf-8")
        df.to_excel(f"output/{os.path.basename(self.file_path)}.xlsx", index=False)
        os.remove(f"output/temp/{os.path.basename(self.file_path)}.csv")

        # Excel operation
        cols = ["A", "B", "C", "D", "E", "F", "G", "N"]
        null_fill = PatternFill(patternType="solid", fgColor="D9D9D9")
        normal_align = Alignment(horizontal="left", vertical="top")
        border = Border(left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))
        header_fill = PatternFill(patternType="solid", fgColor="FFFF99")

        wb = xl.load_workbook(f"output/{os.path.basename(self.file_path)}.xlsx")
        ws = wb["Sheet1"]

        c = ws["C2"]
        ws.freeze_panes = c

        # merge cells
        for col in cols:
            try:
                ws.merge_cells(f"{col}2:{col}{len(items['เลขที่ Invoice']) + 1}")
            except ValueError:
                pass

        # set border
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length

        # styling
        for row in ws.iter_rows():
            for cell in row:
                if cell.row == 1:
                    cell.fill = header_fill
                if cell.value is None:
                    cell.fill = null_fill
                cell.alignment = normal_align
                cell.border = border

        wb.save(f"output/{os.path.basename(self.file_path)}.xlsx")
        print(f'"{os.path.basename(self.file_path)}.xlsx" written to output folder\n')


def correct_words(text: str, mapping: dict) -> str:
    for word in mapping:
        text = text.replace(word, mapping[word])
    return text


class KBANKInvoice(PDFInvoice):

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

    def extract(self, path: str, mode="records") -> dict:
        pdf = pdfplumber.open(path)
        p0 = pdf.pages[1]
        text = p0.extract_text()
        core_pat = re.compile(r"NET AMOUNT\n=+\n(.*)\n=+\nTOTAL", re.DOTALL)
        core = re.search(core_pat, text).group(1)
        core = core.split("\n")

        parsed = [self.parse_row(x) for x in core]
        cols = list(parsed[0].keys())
        data = pd.DataFrame(parsed, columns=cols)
        data = data.drop(columns="INV.DATE")
        data = data.rename(columns={"INV.NUMBER": "เลขที่ Invoice", "INV.AMOUNT": "Amt. (ก่อน Vat)",
                                    "VAT AMT": "Vat. Amt", "WHT AMT": "WHT Amt. (แต่ละ Inv)",
                                    "NET AMOUNT": "จำนวนเงินสุทธิ (แต่ละ Inv)"})
        data = data.replace([''], [None])
        data_dict = data.to_dict(orient=mode)
        return data_dict

    def get_invoice_info(self) -> None:
        self.info = self.pdf.pages

        for page in self.info:
            self.text += page.extract_text()

        self.text = correct_words(self.text, MAPPING)

        type_match = re.search(r"(Subject : )(\w)+", self.text)
        if type_match:
            self.invoice_type = type_match.group(2)

        date_match = re.search(
            r"(Cheque Date : )(\d{2}/\d{2}/\d{4})", self.text)
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
    credit_advice_cols = ["Item No", "Invoice No.", "Date",
                          "Gross Amount", "WHT Amount", "VAT Amount", "Income Type"]
    pre_advice_cols = ["Item No", "Invoice No.",
                       "Date", "Gross Amount", "WHT Amount"]

    def get_invoice_info(self) -> None:
        self.info = self.pdf.pages

        for page in self.info:
            self.text += page.extract_text()

        if all(x in self.text for x in self.credit_advice_cols):
            self.invoice_type = self.invoice_types[0]
        elif all(x in self.text for x in self.pre_advice_cols):
            self.invoice_type = self.invoice_types[1]

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

    def extract(self, file: str, mode: str = "records") -> dict:
        pdf = pdfplumber.open(file)

        cols = self.credit_advice_cols if self.invoice_type == 'Credit Advice Report' else self.pre_advice_cols

        data = pd.DataFrame(columns=cols)

        for page in pdf.pages:
            table = page.extract_table()
            if type(table) != list:
                return {
                    "เลขที่ Invoice": [],
                    "Amt. (ก่อน Vat)": [],
                    "Vat. Amt": [],
                    "Amt. (รวม Vat)": [],
                    "WHT Amt. (แต่ละ Inv)": [],
                    "จำนวนเงินสุทธิ (แต่ละ Inv)": []
                }

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
                    table[1:-1], columns=cols) if self.invoice_type == 'Credit Advice Report' else pd.DataFrame(table,
                                                                                                                columns=cols)

            data = pd.concat([data, df], ignore_index=True)

        data = data.drop(0)
        data = data[["Invoice No.", "Gross Amount", "WHT Amount"]]
        placeholder = [None for _ in range(len(data))]
        data.insert(1, "Amt. (ก่อน Vat)", placeholder)
        data.insert(2, "Vat. Amt", placeholder)
        data["จำนวนเงินสุทธิ (แต่ละ Inv)"] = placeholder

        data = data.rename(columns={"Invoice No.": "เลขที่ Invoice",
                                    "Gross Amount": "Amt. (รวม Vat)", "WHT Amount": "WHT Amt. (แต่ละ Inv)"})
        data = data.replace([''], [None])

        # if all the invoice data are None, return empty dict
        if all(data.isna().all()):
            return {
                    "เลขที่ Invoice": [],
                    "Amt. (ก่อน Vat)": [],
                    "Vat. Amt": [],
                    "Amt. (รวม Vat)": [],
                    "WHT Amt. (แต่ละ Inv)": [],
                    "จำนวนเงินสุทธิ (แต่ละ Inv)": []
                }

        data_dict = data.to_dict(orient=mode)
        return data_dict

class ImgInvoice:
    supplier_name = None
    date = None
    invoice_no = None
    sub_total = None
    grand_total = None
    vat_total = None
    text = ""

    def __init__(self, image_path):
        # Image Preprocessing
        self.image_to_show = image_path
        image_to_scan = utils.preprocess(image_path)
        content = utils.get_content(image_path, image_to_scan)

        # Text Extraction
        image = vision.Image(content=content)
        self.response = client.text_detection(image=image)
        self.annotations = self.response.text_annotations
        self.full_text_annotation = self.response.full_text_annotation
        self.text = str(self.annotations[0].description)

        if not os.path.exists(r'output'):
            os.mkdir(r'output')

    def get_entries(self):
        entries = {
            "ชื่อ Supplier": self.supplier_name,
            "วันที่ในเอกสาร": self.date,
            "เลขที่เอกสาร": self.invoice_no,
            "Amt. (ก่อน Vat)": self.sub_total,
            "Vat. Amt": self.vat_total,
            "Amt. (รวม Vat)": self.grand_total,
        }
        return entries

    def to_json(self):
        entries = self.get_entries()
        utils.pretty_save_json(f"{os.path.basename(self.image_to_show)}.json", entries)

    def to_excel(self):
        entries = self.get_entries()
        df = pd.DataFrame(entries, index=[0])
        df.to_excel(f"output/{os.path.basename(self.image_to_show)}.xlsx", index=False)
        null_fill = PatternFill(patternType="solid", fgColor="D9D9D9")
        normal_align = Alignment(horizontal="left", vertical="top")
        border = Border(left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))
        header_fill = PatternFill(patternType="solid", fgColor="FFFF99")

        wb = xl.load_workbook(f"output/{os.path.basename(self.image_to_show)}.xlsx")
        ws = wb["Sheet1"]
        ws.freeze_panes = ws['A2']

        # set border
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length

        # styling
        for row in ws.iter_rows():
            for cell in row:
                if cell.row == 1:
                    cell.fill = header_fill
                if cell.value is None:
                    cell.fill = null_fill
                cell.alignment = normal_align
                cell.border = border
        
        wb.save(f"output/{os.path.basename(self.image_to_show)}.xlsx")
        print(f'"{os.path.basename(self.image_to_show)}.xlsx" written to output folder\n')

class BulletinInvoice(ImgInvoice):
    def __init__(self, image_path):
        super().__init__(image_path)
        self.supplier_name = "Bulletin Co.,Ltd"

    def get_invoice_info(self):
        text = self.text
        float_sub_total = None
        float_grand_total = None
        # Find payment date
        date_match = re.search(r"Date :[\n ]*(\w+ \d+, \d+)", text)
        if date_match:
            self.date = date_match.group(1)

        # Find invoice number
        inv_no_match = re.search(r"Inv. No :[\n ]*([\w\d -]+)", text)
        if inv_no_match:
            self.invoice_no = inv_no_match.group(1)

        # Find sub total
        sub_total_match = re.search(r"Sub-Total[\n ]*([\d,.]+)", text)
        if sub_total_match:
            self.sub_total = sub_total_match.group(1)
            float_sub_total = utils.to_float(self.sub_total)

        # Find grand total
        grand_total_match = re.search(r"Grand Total[\n ]*([\d,.]+)", text)
        if grand_total_match:
            self.grand_total = grand_total_match.group(1)
            float_grand_total = utils.to_float(self.grand_total)
        
        # Calculate VAT total
        if float_sub_total is not None and float_grand_total is not None:
            float_vat_total = float_grand_total - float_sub_total
            self.vat_total = utils.to_string(float_vat_total)

class MetroUniformInvoice(ImgInvoice):
    def __init__(self, image_path):
        super().__init__(image_path)
        self.supplier_name = "Metro Uniform Co.,Ltd"

    def get_invoice_info(self):
        text = self.text
        float_sub_total = None
        float_grand_total = None

        # Find payment date
        date_match = re.search(r"วันที่[\n ]*(\d{2}/\d{2}/\d{4})", text)
        if date_match:
            self.date = date_match.group(1)

        # Find invoice number
        inv_no_match = re.search(r"เลขที่[\n ]*(IV\d+-\d+)", text)
        if inv_no_match:
            self.invoice_no = inv_no_match.group(1)

        # Find sub total
        sub_total_match = re.search(r"รวมเงิน[\n ]*([\d,.]+)", text)
        if sub_total_match:
            self.sub_total = sub_total_match.group(1)
            float_sub_total = utils.to_float(self.sub_total)

        # Find grand total
        grand_total_match = re.search(r"จำนวนเงินทั้งสิ้น[\n ]*[ก-๛ ]*[\n ]*([\d,.]+)", text)
        if grand_total_match:
            self.grand_total = grand_total_match.group(1)
            float_grand_total = utils.to_float(self.grand_total)

        # Calculate VAT total
        if float_sub_total is not None and float_grand_total is not None:
            float_vat_total = float_grand_total - float_sub_total
            self.vat_total = utils.to_string(float_vat_total)

class LeekaInvoice(ImgInvoice):
    def __init__(self, image_path):
        super().__init__(image_path)
        self.supplier_name = "Leeka Business Co.,Ltd"

    def get_invoice_info(self):
        text = self.text
        float_sub_total = None
        float_grand_total = None

        # Find payment date
        date_match = re.search(r"(วันที่|Date)[\n ]*(\d{2}/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})", text)
        if date_match:
            self.date = date_match.group(2)

        # Find invoice number
        inv_no_match = re.search(r"(เลขที่)*[ก-๛]*[\n ]*(IV\d+|SR\d+)", text)
        if inv_no_match:
            self.invoice_no = inv_no_match.group(2)

        # Find sub total
        sub_total_match = re.search(r"(รวม|ยอดรวม[\n ]*TOTAL)[\n ]*([\d,.]+)", text)
        if sub_total_match:
            self.sub_total = sub_total_match.group(2)
            float_sub_total = utils.to_float(self.sub_total)

        # Find grand total
        grand_total_match = re.search(r"(รวมทั้งสิ้น|NET TOTAL)[\n ]*([\d,.]+)", text)
        if grand_total_match:
            self.grand_total = grand_total_match.group(2)
            float_grand_total = utils.to_float(self.grand_total)

        # Calculate VAT total
        if float_sub_total is not None and float_grand_total is not None:
            float_vat_total = float_grand_total - float_sub_total
            self.vat_total = utils.to_string(float_vat_total)
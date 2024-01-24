from pypdf import PdfWriter, PdfReader
from tkinter import Tk, ttk, Frame, StringVar, BooleanVar, Button
from tkinter.filedialog import askopenfilename, askdirectory
import pandas as pd
import json


class Delivery(Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(background="white")
        self.rowconfigure(5, weight=1)
        self.columnconfigure(1, weight=1)
        self.watermark = PdfReader("asset/delivery.pdf")
        self.reader = ""
        self.file_name = ""
        self.file_location = ""
        self.last_location_file = ""
        self.last_location_save = ""
        self.file_txt = StringVar()
        self.save_txt = StringVar()
        self.error_txt = StringVar()
        self.type = BooleanVar()
        self._load_location()
        self.widget = [
            ttk.Checkbutton(self, text="ปิดชื่อบริษัท", variable=self.type, onvalue=True, offvalue=False,
                            command=lambda: self._toggle_type()),
            ttk.Label(self, anchor="e", width=13, background="white", text="ใบกำกับภาษี"),
            ttk.Label(self, anchor="e", width=13, background="white", text="บันทึกไฟล์ไว้ที่"),
            ttk.Label(self, anchor="e", background="white", foreground="red", textvariable=self.error_txt),
            ttk.Entry(self, state="disabled", textvariable=self.file_txt),
            ttk.Entry(self, state="disabled", textvariable=self.save_txt),
            ttk.Button(self, width=10, text="เลือก", command=lambda: self._select_file()),
            ttk.Button(self, width=10, text="เลือก", command=lambda: self._select_save()),
            ttk.Button(self, width=10, text="บันทึก", command=lambda: self._merg_pdf()),
        ]
        self.widget[0].grid(row=0, column=1, pady=(15, 0), padx=5)
        self.widget[1].grid(row=1, column=0, pady=15, padx=5)
        self.widget[2].grid(row=2, column=0, pady=15, padx=5)
        self.widget[3].grid(row=3, column=1, pady=15, padx=5)
        self.widget[4].grid(row=1, column=1, pady=15, padx=5, sticky="ew")
        self.widget[5].grid(row=2, column=1, pady=15, padx=5, sticky="ew")
        self.widget[6].grid(row=1, column=2, pady=15, padx=5)
        self.widget[7].grid(row=2, column=2, pady=15, padx=5)
        self.widget[8].grid(row=3, column=2, pady=15, padx=5)

    def _toggle_type(self):
        if self.type.get():
            self.watermark = PdfReader("asset/delivery-com.pdf")
        else:
            self.watermark = PdfReader("asset/delivery.pdf")

    def _load_location(self):
        try:
            with open("asset/temp.json", 'r') as openfile:
                data = json.load(openfile)
                self.last_location_file = data["last_location_file"]
                self.last_location_save = data["last_location_save"]
                self.file_location = self.last_location_save
                self.save_txt.set(self.last_location_save)
        except:
            pass

    def _save_locations(self):
        try:
            with open("asset/temp.json", "w") as outfile:
                json.dump({
                    "last_location_file": self.last_location_file,
                    "last_location_save": self.last_location_save
                }, outfile)
        except:
            pass

    def _select_file(self):
        pdf_file = askopenfilename(initialdir=self.last_location_file, filetypes=[('pdf file', '*.pdf')])
        if pdf_file:
            self.file_name = pdf_file.split("/")[-1]
            self.last_location_file = "/".join(pdf_file.split("/")[:-1])
            self.reader = PdfReader(pdf_file)
            self.file_txt.set(pdf_file)
        else:
            pass

    def _select_save(self):
        pdf_folder = askdirectory(initialdir=self.last_location_save)
        if pdf_folder:
            self.file_location = pdf_folder
            self.save_txt.set(pdf_folder)
            self.last_location_save = pdf_folder
        else:
            pass

    def _merg_pdf(self):
        writer = PdfWriter()
        if self.file_location and self.file_name:
            try:
                path = "{}/{}".format(self.file_location, self.file_name)
                self.page = self.reader.pages[0]
                self.page.merge_page(self.watermark.pages[0])
                writer.add_page(self.page)
                with open(path, "wb") as fp:
                    writer.write(fp)
                fp.close()
                self.reader = ""
                self.file_name = ""
                self.file_txt.set("")
                self._save_locations()
                self.error_txt.set("")
            except PermissionError:
                self.error_txt.set("PermissionError. Please Select Another Save location.")
        else:
            pass


class Signature(Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(background="white")
        self.rowconfigure(4, weight=1)
        self.columnconfigure(1, weight=1)
        self.watermark_sing = PdfReader("asset/signature.pdf")
        self.reader = ""
        self.file_name = ""
        self.file_location = ""
        self.last_location_file = ""
        self.last_location_save = ""
        self.file_txt = StringVar()
        self.save_txt = StringVar()
        self.error_txt = StringVar()
        self._load_location()
        self.widget = [
            ttk.Label(self, anchor="e", width=10, background="white", text=""),
            ttk.Label(self, anchor="e", width=13, background="white", text="ใบสั่งซื้อสินค้า"),
            ttk.Label(self, anchor="e", width=13, background="white", text="บันทึกไฟล์ไว้ที่"),
            ttk.Label(self, anchor="e", background="white", foreground="red", textvariable=self.error_txt),
            ttk.Entry(self, state="disabled", textvariable=self.file_txt),
            ttk.Entry(self, state="disabled", textvariable=self.save_txt),
            ttk.Button(self, width=10, text="เลือก", command=lambda: self._select_file()),
            ttk.Button(self, width=10, text="เลือก", command=lambda: self._select_save()),
            ttk.Button(self, width=10, text="บันทึก", command=lambda: self._merg_pdf()),
        ]
        self.widget[0].grid(row=0, column=0, pady=(17, 0), padx=5)
        self.widget[1].grid(row=1, column=0, pady=15, padx=5)
        self.widget[2].grid(row=2, column=0, pady=15, padx=5)
        self.widget[3].grid(row=3, column=1, pady=15, padx=5)
        self.widget[4].grid(row=1, column=1, pady=15, padx=5, sticky="ew")
        self.widget[5].grid(row=2, column=1, pady=15, padx=5, sticky="ew")
        self.widget[6].grid(row=1, column=2, pady=15, padx=5)
        self.widget[7].grid(row=2, column=2, pady=15, padx=5)
        self.widget[8].grid(row=3, column=2, pady=15, padx=5)

    def _load_location(self):
        try:
            with open("asset/temp.json", 'r') as openfile:
                data = json.load(openfile)
                self.last_location_file = data["last_location_file"]
                self.last_location_save = data["last_location_save"]
                self.file_location = self.last_location_save
                self.save_txt.set(self.last_location_save)
        except:
            pass

    def _save_locations(self):
        try:
            with open("asset/temp.json", "w") as outfile:
                json.dump({
                    "last_location_file": self.last_location_file,
                    "last_location_save": self.last_location_save
                }, outfile)
        except:
            pass

    def _select_file(self):
        pdf_file = askopenfilename(initialdir=self.last_location_file, filetypes=[('pdf file', '*.pdf')])
        if pdf_file:
            self.file_name = pdf_file.split("/")[-1]
            self.last_location_file = "/".join(pdf_file.split("/")[:-1])
            self.reader = PdfReader(pdf_file)
            self.file_txt.set(pdf_file)
        else:
            pass

    def _select_save(self):
        pdf_folder = askdirectory(initialdir=self.last_location_save)
        if pdf_folder:
            self.file_location = pdf_folder
            self.save_txt.set(pdf_folder)
            self.last_location_save = pdf_folder
        else:
            pass

    def _merg_pdf(self):
        writer = PdfWriter()
        if self.file_location and self.file_name:
            try:
                path = "{}/{}".format(self.file_location, self.file_name)
                self.page = self.reader.pages[0]
                self.page.merge_page(self.watermark_sing.pages[0])
                writer.add_page(self.page)
                with open(path, "wb") as fp:
                    writer.write(fp)
                fp.close()
                self.reader = ""
                self.file_name = ""
                self.file_txt.set("")
                self._save_locations()
                self.error_txt.set("")
            except PermissionError:
                self.error_txt.set("PermissionError. Please Select Another Save location.")
        else:
            pass


class CheckReceipt(Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(background="white")
        self.rowconfigure(4, weight=1)
        self.columnconfigure(1, weight=1)
        self.file_txt = StringVar()
        self.target_txt = StringVar()
        self.result_txt = StringVar()
        self.widget = [
            ttk.Label(self, anchor="e", width=10, background="white", text=""),
            ttk.Label(self, anchor="e", width=13, background="white", text="ไฟล์ Excel"),
            ttk.Label(self, anchor="e", width=13, background="white", text="จำนวนเงิน"),
            ttk.Label(self, anchor="e", background="white", textvariable=self.result_txt),
            ttk.Entry(self, state="disabled", textvariable=self.file_txt),
            ttk.Entry(self, justify='center', textvariable=self.target_txt),
            ttk.Button(self, width=10, text="เลือก", command=lambda: self._select_file()),
            ttk.Button(self, width=10, text="ตรวจสอบ", command=lambda: self._check()),
        ]
        self.widget[0].grid(row=0, column=0, pady=(17, 0), padx=5)
        self.widget[1].grid(row=1, column=0, pady=15, padx=5)
        self.widget[2].grid(row=2, column=0, pady=15, padx=5)
        self.widget[3].grid(row=3, column=1, pady=15, padx=5)
        self.widget[4].grid(row=1, column=1, pady=15, padx=5, sticky="ew")
        self.widget[5].grid(row=2, column=1, pady=15, padx=5, sticky="ew")
        self.widget[6].grid(row=1, column=2, pady=15, padx=5)
        self.widget[7].grid(row=2, column=2, pady=15, padx=5)

    def _select_file(self):
        excel_file = askopenfilename(filetypes=[('xls file', '*.xls')])
        if excel_file:
            self.file_txt.set(excel_file)
        else:
            pass

    def _check(self):
        result, debtor = [], {}
        if self.file_txt.get() and self.target_txt.get():
            data = pd.read_excel(self.file_txt.get())
            target = float(self.target_txt.get())
            for i in range(1, len(data) - 4):
                customer = data.iloc[i, 5]
                invoices = data.iloc[i, 1][0:6]
                debt_val = data.iloc[i, 7]
                if customer not in debtor:
                    debtor[customer] = {invoices: float(debt_val)}
                else:
                    if invoices in debtor[customer]:
                        debtor[customer][invoices] += float(debt_val)
                    else:
                        debtor[customer][invoices] = float(debt_val)
            for customer, invoices in debtor.items():
                for inv, val in invoices.items():
                    total_debt = sum(invoices.values())
                    if val == target or (target + 25) >= val >= (target - 25):
                        result.append(f"{customer} {inv} ยอด {round(val, 2)}")
                    elif total_debt == target or (target + 25) >= total_debt >= (target - 25):
                        result.append(f"{customer} {invoices}")
                        break
            if not result:
                self.result_txt.set("ไม่พบข้อมูล")
            else:
                self.result_txt.set('\n'.join(result))


class MainFrame(Frame):
    def __init__(self, container):
        super().__init__(container)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.pack(fill="both", expand=True)
        self.style = ttk.Style()
        self.style.configure("TCheckbutton", background="white")
        self.style.configure("TEntry", background="white")
        self.number = 0
        self.menu = Frame(self, background="white")
        self.content = Frame(self, background="white", borderwidth=1)
        self.child = [
            Signature(self.content),
            Delivery(self.content),
            CheckReceipt(self.content)
        ]
        self.widget = [
            Button(self.menu, width=20, bg="gray80", bd=1, relief="solid", text="เซ็นใบสั่งซื้อสินค้า",
                   command=lambda: self._switch_frame(0)),
            Button(self.menu, width=20, bg="white", bd=0, relief="solid", text="ปิดราคาขายสินค้า",
                   command=lambda: self._switch_frame(1)),
            Button(self.menu, width=20, bg="white", bd=0, relief="solid", text="ตรวจสอบยอดรับ",
                   command=lambda: self._switch_frame(2))
        ]
        self.menu.grid(row=0, column=0, sticky="nsew")
        self.content.grid(row=1, column=0, sticky="nsew")
        self.widget[0].grid(row=0, column=0, padx=2, sticky="w")
        self.widget[1].grid(row=0, column=1, padx=2, sticky="w")
        self.widget[2].grid(row=0, column=2, sticky="w")
        self.child[0].pack(fill="both", expand=True, pady=15, padx=1)

    def _switch_frame(self, n):
        if n != self.number:
            for i in range(len(self.child)):
                self.child[i].pack_forget()
                self.widget[i].configure(bg="white", bd=0)
            self.number = n
            self.child[n].pack(fill="both", expand=True, pady=15, padx=1)
            self.widget[n].configure(bg="gray80", bd=1)


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("BDF Editor")
        self.posRight = int(self.winfo_screenwidth() / 2 - 300)
        self.posDown = int(self.winfo_screenheight() / 2 - 150)
        self.geometry("600x300+{}+{}".format(self.posRight, self.posDown))
        self.resizable(False, False)
        self.mainFrame = MainFrame(self)


if __name__ == '__main__':
    app = App()
    app.mainloop()

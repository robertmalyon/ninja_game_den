from tkinter import Tk, Frame, BOTTOM, Entry, Label, END, Text, WORD, IntVar, Checkbutton, Button, messagebox
#import sys
import pypyodbc
from datetime import datetime
# import win32print
import re
import csv
import string
from woocommerce import API

# sys.path.append(r'c:\users\rob\appdata\local\programs\python\python37-32\lib\site-packages')

# Remove syspath
# Update CSV path
# Update DB path throughout, remove .accdb to correct driver issue
# Uncomment win32print import statement and print functionality lines
# Don't forget to get Zebra barcode font
# Rename file to .pyw for consoleless operation
dbDriver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
dbPath = r"C:\Users\Rob\Desktop\Test\Gamestar - Brighton.accdb;"
csvPath = r'C:\Users\Rob\Desktop\Products.csv'
printer = "ZDesigner GK420d"


def woo_connect(data):
    wcapi = API(
        url="https://ninjagameden.co.uk",
        consumer_key="",
        consumer_secret="",
        query_string_auth=True)

    new_online_product = wcapi.post("products", data).json()
    print(new_online_product)
    online_id = new_online_product.get('id')
    return online_id


class DatabaseConnection:
    def __init__(self):
        self.conn = pypyodbc.connect(r"Driver=" + dbDriver + r"DBQ=" + dbPath)
        self.cursor = self.conn.cursor()


class CsvWriter:
    def __init__(self, fields):
        with open(csvPath, mode='a', newline='') as products_csv:
            product_writer = csv.writer(
                products_csv, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            product_writer.writerow(fields)


class CreateEposBat:
    def __init__(self, serial_no, part_no, price):
        db = DatabaseConnection()
        db.cursor.execute("SELECT MAX([Batch No]) FROM [Batch Details]")
        self.bat_count = db.cursor.fetchone()[0] + 1
        db.cursor.execute('''
                            INSERT INTO [Batch Details] ( [Batch No], [Serial No], Status, [Part No], Qty, Price, 
                            [Date Acquired], Saleable ) 
                            VALUES ( ?, ?, 2, ?, 1, ?, ?, Yes)''', (self.bat_count, serial_no, part_no, price,
                                                                    datetime.today()))
        db.cursor.commit()
        db.cursor.close()


class Product:
    def __init__(self, barcode_part):
        self.barcode_part = barcode_part
        db = DatabaseConnection()
        if re.search('[a-zA-Z]', self.barcode_part):
            db.cursor.execute("SELECT Stock.[Part No], Stock.[Bar Code], Stock.[Description], Stock.[Platform], "
                              "Stock.[Second Hand Price] FROM Stock WHERE Stock.[Part No]= ?", (self.barcode_part,))
        else:
            db.cursor.execute("SELECT Stock.[Part No], Stock.[Bar Code], Stock.[Description], Stock.[Platform], "
                              "Stock.[Second Hand Price] FROM  [Barcode Lookup] INNER JOIN Stock ON "
                              "[Barcode Lookup].[Part No] = Stock.[Part No] WHERE [Barcode Lookup].[Bar Code] = ?", (self.barcode_part,))
            if db.cursor.fetchone() is not None:
                db.cursor.execute("SELECT Stock.[Part No], Stock.[Bar Code], Stock.[Description], Stock.[Platform], "
                                  "Stock.[Second Hand Price] FROM  [Barcode Lookup] INNER JOIN Stock ON "
                                  "[Barcode Lookup].[Part No] = Stock.[Part No] WHERE [Barcode Lookup].[Bar Code] = ?", (self.barcode_part,))
            else:
                db.cursor.execute("SELECT Stock.[Part No], Stock.[Bar Code], Stock.[Description], Stock.[Platform], "
                                  "Stock.[Second Hand Price] FROM Stock WHERE Stock.[Bar Code]= ?",
                                  (self.barcode_part,))

        self.row = db.cursor.fetchone()
        self.part = str(self.row[0])
        self.barcode = str(self.row[1])
        # Convert this variable to string for string concatenation with ZPL
        self.description = str(self.row[2])
        self.platform = str(self.row[3])
        self.secondHandPrice = str(round(self.row[4], 2))
        self.online_title = self.description.lower()
        self.online_title = string.capwords(self.online_title)
        self.online_title = self.online_title.replace(
            " Cart Only", " (Cartridge Only)")
        self.online_title = self.online_title.replace(
            " Cart", " (Cartridge Only)")
        self.online_title = self.online_title.replace(" Boxed", " (Boxed)")
        self.online_title = self.online_title.replace("ntsc-j", "**NTSC-J**")
        self.online_title = self.online_title.replace("Ntsc-j", "**NTSC-J**")
        self.online_title = self.online_title.replace("ntsc-u", "**NTSC-U**")
        self.online_title = self.online_title.replace("Ntsc-u", "**NTSC-U**")
        self.online_title = re.sub(r"[\[].*?[\]]", "", self.online_title)
        self.online_name = self.online_title.strip()
        self.online_name = re.sub(r"[\([].*?[\])]", "", self.online_name)
        self.online_name = self.online_name.strip()
        self.generic_desription = "This game is in very good clean condition."
        if "(Cartridge Only)" in self.online_title:
            self.online_description = "This is a cartridge only copy of " + self.online_name + " so there is no box " \
                                                                                               "or manual. The label " \
                                                                                               "is in very good " \
                                                                                               "condition with no " \
                                                                                               "defects apart from " \
                                                                                               "some minor shelf wear."
        elif "(Boxed)" in self.online_title:
            self.online_description = "This is a boxed copy of " + self.online_name + " which includes the manual. " \
                "The cover is original and is in very good condition with no defects. The outer box is also in very" \
                " good condition with some minor shelf wear."

        else:
            self.online_description = "This copy of " + self.online_name + " is in very good clean condition with " \
                                                                           "some minor shelf wear to the outer case."

        db.conn.close()

    def update_price(self, updated_price):
        # Establish database connection
        db = DatabaseConnection()
        db.cursor.execute("UPDATE Stock SET Stock.[Second Hand Price] = ? WHERE (((Stock.[Part No])= ?))",
                          (updated_price, self.part))
        db.cursor.commit()
        db.conn.close()


class LabelFront:
    def __init__(self, second_hand_price, description, date, barcode):
        self.raw_data_string = r"^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ~DG000.GRF," \
                               r"01024,016,,::::::::::L040H040H040H040H040H040H040,:::::T03C1E07F83FHF87F8," \
                               r"N040H040F84F9FFE3FHF9FFE,T0780F3E0E3C0F078F,T0780F3C0C3C0H0787,T0780F3C003C0H078" \
                               r"780T0780F1F003C0H078780L03FC30I0780F07F83C0H078780L07FHFJ0780F007E3FF0078780L07FFE" \
                               r"0I0F80F801F3C0H078780L040H040H0F80FC01F7C004787C0T0780F0C1F3C0H078F,T07E3F1F3F3C078" \
                               r"78F,T03FFE3FHF3FHF87BE,T01FFC7FFE7FHF87FC,U0HF83FFC3FHF87F8,U07F00FF01FHF9FE0,," \
                               r":::::::::::::::::::::::::::::~DG001.GRF,01024,032,,:::::::H040H040H040H040H040H040" \
                               r"H040H040H040H040H040H040H040H040,,:::::H0600F8FC01F0H0F03E0L0FE007C03C0607FHFJ01F" \
                               r"E07FHF600F80,H0707FBEE07F403E47F040H043FFC0FE41E0707FHFH0407FF87FHF707F80,H0781F1E" \
                               r"F03E001E0FF0K078F81FE01F0F0781E0I01E3C781E781F,H07C1F1EF83E001E03F0K0F03007E01F9F8" \
                               r"780K01E1C78007C1F,H07E1F1EFC3E001E03F80J0E03007F01FHF8780K01E1E78007E1F,H07F9F1EFF" \
                               r"3E001E03F80I01E0I07F01FHFC780K01E1E78007F9F,H07FHF1EFHFE001E077C0I01E0I0EF81FHFC78" \
                               r"0K01E1E78007FHF,H07BFF1EF7FE001E0E7C0I01E0H01CF81EFFC7FE0J01E1E7FE07BFF,H078FF1EF" \
                               r"1FE001E0E3E0I01E1FE1C7C1E77C780K01E1E780078FF,H0781F5EF07E005E1C7E00401E4FE3C7C1" \
                               r"EC3E780040H05E1E7800781F,H0781F1EF03E1C1E7FHFJ01F07EFHFE1E03E780K01E3C7800781F,H0" \
                               r"781F1EF03E3C3E381F0J0F8FE703E1E03E780F0I01E3C780F781F,H0781F3EF03E7E7C781F80I0IFC" \
                               r"F03F1E03E7FHFJ01EF87FHF781F,H0781F3EF03E7FF8781F80I07FFCF03F3E03EFIFJ01FF0FIF78" \
                               r"1F,H0781F1EF03E7FF8781F80I03FF0F03F1E03C7FHFJ01FE07FHF781F,H0781F9EF03F1FE0781F" \
                               r"C0J0FC0F03F9F0383FHFJ07F803FHF781F80,H07C0I0F80K03830O07060I070V07C,^XA" \
            r"^MMT" \
            r"^PW384" \
            r"^LL0256" \
            r"^LS0" \
            r"^FT224,32^XG000.GRF,1,1^FS" \
            r"^FT0,32^XG001.GRF,1,1^FS" \
            r"^FT8,137^A0N,99,98^FH\^FD\9C" + second_hand_price + "^FS" \
            r"^FT8,195^A0N,28,28^FH\^FD" + date + "^FS" \
            r"^FT7,238^A0N,28,28^FH\^FD" + description + "^FS" \
            r"^BY2,2,36^FT178,190^BEN,,N,N" \
            r"^FD" + barcode + "^FS" \
            r"^PQ1,0,1,Y^XZ" \
            r"^XA^ID000.GRF^FS^XZ" \
            r"^XA^ID001.GRF^FS^XZ"

        self.raw_data = bytes(self.raw_data_string, "utf-8")


class LabelBack:
    def __init__(self, bat_no, description, selectBarcodeSmall):
        if selectBarcodeSmall == False:
            self.raw_data_string = r"^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ" \
                r"^XA" \
                r"^MMT" \
                r"^PW384" \
                r"^LL0256" \
                r"^LS0" \
                r"^BY4,3,160^FT11,213^BCN,,Y,N" \
                r"^FD>:" + bat_no + "^FS" \
                                    r"^FT52,37^A0N,28,28^FH\^FD" + description + "^FS" \
                                    r"^PQ1,0,1,Y^XZ"
        else:
            self.raw_data_string = r"^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ" \
                                   r"^XA" \
                                   r"^MMT" \
                                   r"^PW384" \
                                   r"^LL0256" \
                                   r"^LS0" \
                                   r"^BY4,3,160^FT11,213^BCN,,Y,N" \
                                   r"^FD>:" + bat_no + "^FS" \
                                                       r"^FT52,37^A0N,28,28^FH\^FD" + description + "^FS" \
                                                                                                    r"^PQ1,0,1,Y^XZ"

            self.raw_data_string = "^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ^XA^MMT^" \
                "PW384^LL0256^LS0^BY1,3,29^FT132,36^BCN,,Y,N^FD>:" + bat_no + "^FS^" \
                "PQ1,0,1,Y^XZ"

        self.raw_data = bytes(self.raw_data_string, "utf-8")


class MainWindow:
    def __init__(self):
        self.part_no = "--  "
        self.platform = "--  "
        self.description = "--  "
        self.price = "0.00"
        self.root = Tk()
        self.root.title("NGD POS")
        # Fetch date and convert to string
        self.date = str(datetime.today().strftime('%d/%m/%Y'))

        self.topframe = Frame(self.root, relief='flat', borderwidth=20)
        self.topframe.pack()
        self.bottomframe = Frame(self.root, relief='flat', borderwidth=20)
        self.bottomframe.pack(side=BOTTOM)

        self.lblPartNo = Label(self.topframe, text="Enter barcode or part no: ", font=("Ariel", 14), width=25,
                               anchor="e", pady=10)
        self.lblPartNo.grid(row=0, column=0)
        self.txtPartNo = Entry(self.topframe, width=20)
        self.txtPartNo.grid(row=0, column=1, sticky='w')
        self.txtPartNo.config(font=('ariel', 14,))
        self.txtPartNo.bind("<Return>", self.on_enter)
        self.txtPartNo.focus_set()

        self.labelPartNo2 = Label(self.topframe, text="Part No: ", font=(
            "Ariel", 10), width=25, anchor="e", pady=5)
        self.labelPartNo2.grid(row=1, column=0)
        self.lblPartNo2 = Label(self.topframe, width=10,
                                text=self.part_no, anchor='w')
        self.lblPartNo2.grid(row=1, column=1, sticky='w')
        self.lblPartNo2.config(font=('ariel', 10,))

        self.lblPlatform = Label(self.topframe, text="Platform: ", font=(
            "Ariel", 10), width=25, anchor="e", pady=5)
        self.lblPlatform.grid(row=2, column=0)
        self.lblPlatform2 = Label(
            self.topframe, width=5, text=self.platform, anchor='w')
        self.lblPlatform2.grid(row=2, column=1, sticky='w')
        self.lblPlatform2.config(font=('ariel', 10,))

        self.lblDescription = Label(self.topframe, text="Description: ", font=("Ariel", 10), width=25,
                                    anchor="e", pady=5)
        self.lblDescription.grid(row=3, column=0)
        self.lblDescription2 = Label(
            self.topframe, width=30, height=2, text=self.description, anchor='w')
        self.lblDescription2.grid(row=3, column=1, sticky='w')
        self.lblDescription2.config(font=('ariel', 10,))

        self.lblSecondHandPrice = Label(self.topframe, text="Second Hand Price: £", font=("Ariel", 14),
                                        width=25, anchor="e", pady=5)
        self.lblSecondHandPrice.grid(row=4, column=0)
        self.txtSecondHandPrice = Entry(self.topframe, width=10)
        self.txtSecondHandPrice.grid(row=4, column=1, sticky='w')
        self.txtSecondHandPrice.config(font=('ariel', 14,))
        self.set_price(self.price)

        self.lblDatePriced = Label(self.topframe, text="Date Priced: ", font=("Ariel", 10), width=25,
                                   anchor="e", pady=10)
        self.lblDatePriced.grid(row=5, column=0)
        self.txtDatePriced = Entry(self.topframe, width=10)
        self.txtDatePriced.grid(row=5, column=1, sticky='w')
        self.txtDatePriced.insert(END, self.date)
        self.txtDatePriced.config(font=('ariel', 10,))

        self.lblSku = Label(self.topframe, text="SKU: ", font=(
            "Ariel", 10), width=25, anchor='e', pady=10)
        self.lblSku.grid(row=6, column=0)
        self.txtSku = Label(self.topframe, text="-- ",
                            font=("Ariel", 10), width=25, anchor='w')
        self.txtSku.grid(row=6, column=1, sticky='w')

        self.qtyBox = Label(self.topframe, text="Qty: ",
                            font=("Ariel", 10), anchor='e')
        self.qtyBox.grid(row=6, column=2, sticky='e')
        self.qty_entry = Entry(self.topframe, width=10)
        self.qty_entry.grid(row=6, column=3, sticky='w')
        self.qty_entry.config(font=('ariel', 10,))

        self.online_title = Label(self.topframe, text="Online Title: ", font=("Ariel", 10), width=25, anchor='e',
                                  pady=10)
        self.online_title.grid(row=7, column=0)
        self.online_title_entry = Text(self.topframe, width=50, height=2)
        self.online_title_entry.grid(row=7, column=1, sticky='w')
        self.online_title_entry.config(font=('ariel', 10,), wrap=WORD)

        self.online_price = Label(self.topframe, text="Online Price: ", font=("Ariel", 10), width=25, anchor='e',
                                  pady=10)
        self.online_price.grid(row=7, column=2)
        self.online_price_entry = Entry(self.topframe, width=10)
        self.online_price_entry.grid(row=7, column=3, sticky='w')
        self.online_price_entry.config(font=('ariel', 10,))

        self.featured_product_label = Label(self.topframe, text="Featured Product? ", font=("Ariel", 14),
                                            anchor='e')
        self.featured_product_label.grid(row=10, column=0,)
        self.featured_product_var = IntVar()
        self.featured_product_tick = Checkbutton(self.topframe, variable=self.featured_product_var, onvalue=True,
                                                 offvalue=False, anchor='e')
        self.featured_product_tick.grid(row=10, column=0, sticky='e')
        self.featured_product_label.config(font=('ariel', 10,))

        self.visible_product_label = Label(self.topframe, text="Visible? ", font=("Ariel", 14),
                                           anchor='w')
        self.visible_product_label.grid(row=10, column=1)
        self.visible_product_tick = Checkbutton(
            self.topframe, onvalue=1, offvalue=0, anchor='e', padx=30)
        self.visible_product_tick.grid(row=10, column=1, sticky='e')
        self.visible_product_label.config(font=('ariel', 10,))

        self.selectBarcodeSmall_label = Label(self.topframe, text="Small Barcode? ", font=("Ariel", 10), width=25,
                                              anchor='w')
        self.selectBarcodeSmall_label.grid(row=10, column=2)
        self.selectBarcodeSmall = IntVar()
        self.barcodeSize = Checkbutton(self.topframe, variable=self.selectBarcodeSmall, onvalue=True, offvalue=False,
                                       height=5, anchor='e')
        self.barcodeSize.grid(row=10, column=2, sticky='e')

        self.online_description = Label(self.topframe, text="Online Description: ", font=("Ariel", 10), width=25,
                                        anchor='e', pady=10)
        self.online_description.grid(row=11, column=0)
        self.online_description_entry = Text(self.topframe, width=70, height=5)
        self.online_description_entry.grid(
            row=11, column=1, columnspan=3, sticky='w')
        self.online_description_entry.config(font=('ariel', 10,), wrap=WORD)

        self.btnPrint = Button(
            self.bottomframe, text="Print F12", command=self.label_print)
        self.btnPrint.config(font=("Ariel", 16, "bold"), padx=20, pady=10)
        self.btnPrint.grid(row=1, column=1, padx=20, pady=20)
        self.btnExit = Button(self.bottomframe, text="Exit",
                              command=self.root.destroy)
        self.btnExit.config(font=("Ariel", 16, "bold"), padx=20, pady=10)
        self.btnExit.grid(row=1, column=2, padx=20, pady=20)

        self.root.bind('<Escape>', lambda c: self.root.destroy())
        self.root.bind('<F12>', lambda a: self.label_print())
        self.txtPartNo.selection_range(0, END)
        self.root.mainloop()

    def set_price(self, set_price):
        self.txtSecondHandPrice.delete(0, END)
        self.txtSecondHandPrice.insert(0, set_price)
        self.price = set_price

    def get_next_bat(self):
        db = DatabaseConnection()
        db.cursor.execute("SELECT MAX([Batch No]) FROM [Batch Details]")
        return db.cursor.fetchone()[0] + 1

    def requery(self):
        try:
            self.product = Product(self.txtPartNo.get())
        except TypeError:
            messagebox.showerror("Barcode/Part No Does Not Exist", "Please try a different Barcode or Part No or call "
                                                                   "support helpline on 0121-do-1.")
        else:
            self.platform = self.product.platform
            self.online_title = self.product.online_title
            self.description = self.product.description
            self.set_price(self.product.secondHandPrice)
            self.lblPlatform2.config(text=self.platform)
            self.lblDescription2.config(text=self.description)
            self.part_no = self.product.part
            self.lblPartNo2.config(text=self.part_no)
            self.online_title_entry.delete(1.0, END)
            self.online_title_entry.insert(END, self.online_title)
            self.online_price_entry.delete(0, END)
            self.online_price_entry.insert(END, self.price)
            self.txtSku.config(text="BAT" + str(self.get_next_bat()))
            self.qty_entry.delete(0, END)
            self.qty_entry.insert(END, 1)
            self.online_description_entry.delete(1.0, END)
            self.online_description_entry.insert(
                END, self.product.online_description)

    def on_enter(self, event):
        self.requery()
        self.txtSecondHandPrice.focus_set()
        self.txtSecondHandPrice.selection_range(0, END)

    def create_csv(self):
        title = self.online_title_entry.get(1.0, END)
        created_at = datetime.now()
        sku = self.bat_no
        price = self.online_price_entry.get()
        stock_quantity = self.qty_entry.get()
        in_stock = 'True'
        back_orders_allowed = 'No'
        featured = 'True' if self.featured_product_var.get() == True else 'False'
        description = self.online_description_entry.get(1.0, END)
        categories = self.platform

        fields = [title, created_at, sku, price, stock_quantity, in_stock, back_orders_allowed, featured, description,
                  categories]

        CsvWriter(fields)

    def addUniqueBatch(self):
        msgAddBatch = messagebox.askquestion("Add product to online stock", "Do you want to add this product to "
                                             "online stock and create a unique BAT no? Select No if you just "
                                             "want to print labels", icon='warning')

        self.price = self.product.secondHandPrice
        if msgAddBatch == 'yes':
            self.barcode = self.product.barcode
            self.front_label = LabelFront(
                self.product.secondHandPrice, self.description, self.date, self.barcode)
            print(self.front_label.raw_data)
            # printer_name = "ZDesigner GK420d"
            # hPrinter = win32print.OpenPrinter(printer_name)
            # hJob = win32print.StartDocPrinter(hPrinter, 1, ("test of raw data", None, "RAW"))
            # win32print.StartPagePrinter(hPrinter)
            # win32print.WritePrinter(hPrinter, self.front_label.raw_data)
            # win32print.EndPagePrinter(hPrinter)
            # win32print.EndDocPrinter(hPrinter)
            # win32print.ClosePrinter(hPrinter)
            self.bat = CreateEposBat(
                "available_online", self.product.part, self.product.secondHandPrice)
            self.bat_no = "BAT" + str(self.bat.bat_count)
            self.txtSku.config(text="BAT" + str(self.get_next_bat()))
            self.back_label = LabelBack(
                self.bat_no, self.product.description, self.selectBarcodeSmall.get())
            print(self.back_label.raw_data)
            # printer_name = "ZDesigner GK420d"
            # hPrinter = win32print.OpenPrinter(printer_name)
            # hJob = win32print.StartDocPrinter(hPrinter, 1, ("test of raw data", None, "RAW"))
            # win32print.StartPagePrinter(hPrinter)
            # win32print.WritePrinter(hPrinter, self.back_label.raw_data)
            # win32print.EndPagePrinter(hPrinter)
            # win32print.EndDocPrinter(hPrinter)
            # win32print.ClosePrinter(hPrinter)
            # fields = [self.date, self.part_no, self.description, self.platform, self.price]
            # CsvWriter(fields)
            self.create_csv()

            name = str(self.online_title_entry.get(1.0, END))
            sku = str(self.bat_no)
            price = str(self.online_price_entry.get())
            stock_quantity = str(self.qty_entry.get())
            featured = 'True' if self.featured_product_var.get() == True else 'False'
            description = str(self.online_description_entry.get(1.0, END))
            #categories = str(self.platform)

            data = {
                "name": name,
                "type": "simple",
                "sku": sku,
                "status": "private",
                "featured": featured,
                "description": description,
                "regular_price": price,
                "stock_quantity": stock_quantity,
                # "categories": [categories],
                "images": []
            }

            online_id = woo_connect(data)
            db = DatabaseConnection()
            db.cursor.execute("UPDATE [Batch Details] SET [Batch Details].[Serial No] =? "
                              "WHERE ((([Batch Details].[Batch No])=?))", (online_id, self.bat.bat_count))
            db.cursor.commit()
            db.cursor.close()

        else:
            self.barcode = self.product.barcode
            front_label = LabelFront(
                self.price, self.description, self.date, self.barcode)
            print(front_label.raw_data)
            # printer_name = "ZDesigner GK420d"
            # hPrinter = win32print.OpenPrinter(printer_name)
            # hJob = win32print.StartDocPrinter(hPrinter, 1, ("test of raw data", None, "RAW"))
            # win32print.StartPagePrinter(hPrinter)
            # win32print.WritePrinter(hPrinter, front_label.raw_data)
            # win32print.EndPagePrinter(hPrinter)
            # win32print.EndDocPrinter(hPrinter)
            # win32print.ClosePrinter(hPrinter)

    def label_print(self):
        try:
            self.product.update_price(self.txtSecondHandPrice.get())
            self.requery()
        except pypyodbc.DataError:
            messagebox.showerror(
                "Invalid Price", "Please make sure that Second Hand Price is a number.")
        except AttributeError:
            messagebox.showerror(
                "No product selected", "Please enter a barcode or Part No and press enter.")
        else:
            self.addUniqueBatch()


MainWindow()

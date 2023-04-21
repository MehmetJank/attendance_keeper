from tkinter import Tk, Frame, Entry, Button, END, RIGHT, Label, Listbox
from tkinter import ttk

from openpyxl import load_workbook

class AttendanceKeeperApp(Frame):

    def import_list(self):
        # Excel dosyasını açar
        wb = load_workbook('ENGR 102 Student List.xlsx')

        # Aktif sayfayı seçer
        ws = wb.active

        # ----------------- Bölüm Seçimi -----------------
        # combo_box_items listesini tanımlar. Yani bolumlerin listesi
        combo_box_items = []

        # Ilk satırı haric D sütunundaki tüm hücreleri ekler
        for cell in ws['D'][1:]:
            combo_box_items.append(cell.value)

        # combo_box_items listesindeki tekrar eden elemanları siler
        combo_box_items = list(set(combo_box_items))

        # combo_box_items listesindeki elemanları alfabetik sıraya koyar
        combo_box_items.sort()

        # combo_box_items listesindeki elemanları ekrana yazdırır
        # for items in combo_box_items:
            # print(items)

        # user_choice_section = int(input("Bolum Seciniz:"))
        # user_choice_section = combo_box_items[user_choice_section - 1]
        # print(user_choice_section)


        # min row 2. satirdan baslar cunku basliklar olmayacak, min col 1. sutundan baslayip, max col 4. sutuna kadar döngüye alir

        # ----------------- Öğrenci Seçimi Function -----------------


        def section_students(section: str):
            # min row 2. satirdan baslar cunku basliklar olmayacak, min col 1. sutundan baslayip, max col 4. sutuna kadar döngüye alir
            for row in ws.iter_rows(min_row=2, min_col=1, max_col=4):
                cell = row[3]  # 4. sutun
                if cell.value == section:  # 4. sutunun degeri bolumun degerine esitse
                    # 2. sutunun degerini bosluklara gore ayirir
                    name_surname = row[1].value.split()
                    # 2. sutunun ilk elemani soyadi, ikinci elemani adi, 1. sutunun degeri numarasi
                    print(name_surname[1], name_surname[0], row[0].value)

        return combo_box_items

    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()

    # Custom Button
    def custom_button(
        self,
        text: str,
        row: int,
        column: int,
        sticky: str,
        command: object,
        width: int = 0,
        height: int = 0,
        columnspan: int = 1,
        rowspan: int = 1,
        padx: int = 0,
        pady: int = 0,
    ):
        button = Button(text=text, width=width, height=height, command=command)
        button.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            rowspan=rowspan,
            padx=padx,
            pady=pady,
            sticky=sticky,
        )

    # Custom Text Label
    def custom_text_label(
        self,
        text: str,
        row: int,
        column: int,
        sticky: str,
        width: int = 0,
        height: int = 0,
        columnspan: int = 1,
        rowspan: int = 1,
        padx: int = 0,
        pady: int = 0,
    ):
        label = Label(
            text=text,
            width=width,
            height=height,
        )
        label.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            rowspan=rowspan,
            padx=padx,
            pady=pady,
            sticky=sticky,
        )

    def custom_listbox(self, row: int, column: int, list_values : list = []):
        listBox = Listbox(width=30, height=5)
        listBox.grid(
            row=row,
            column=column,
            columnspan=2,
            rowspan=3,
            padx=3,
            pady=3,
            sticky="nswe",
        )
        for item in list_values:
            listBox.insert(END, item)

    def custom_combo_box(self, row: int, column: int, width: int, height: int, list_values: list = [0]):
        combo = ttk.Combobox(width=width, height=height)
        combo.grid(
            row=row,
            column=column,
            columnspan=1,
            rowspan=1,
            padx=3,
            pady=3,
            sticky="nswe",
        )
        combo["values"] = list_values
        combo.current(0)

    def initUI(self):
        # Boyutlandırma
        for i in range(5):
            self.master.columnconfigure(i, weight=1)
        for i in range(7):
            self.master.rowconfigure(i, weight=1)

        # uygulama başlığı
        self.parent.title("Attendance Keeper")

        # Row 0
        self.custom_text_label(
            text="AttandanceKeeper v1.0",
            row=0,
            column=0,
            sticky="nswe",
            padx=3,
            pady=3,
            columnspan=5,
        )

        # Row 1
        self.custom_text_label(
            text="Select Student list Excel File",
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            padx=3,
            pady=3,
        )

        self.custom_button(
            text="Import List",
            row=1,
            column=2,
            columnspan=1,
            width=15,
            sticky="nswe",
            command="",
            padx=3,
            pady=3,
        )

        # Row 2
        self.custom_text_label(
            text="Select a Student",
            row=2,
            column=0,
            columnspan=2,
            sticky="w",
            padx=3,
            pady=3,
        )

        self.custom_text_label(
            text="Section:",
            row=2,
            column=2,
            columnspan=1,
            sticky="nswe",
            padx=3,
            pady=3,
        )

        self.custom_text_label(
            text="Attanted Students",
            row=2,
            column=3,
            columnspan=2,
            sticky="e",
            padx=3,
            pady=3,
        )

        # Row 3
        my_test_list = ["test1", "test2", "test3", "test4", "test5"]
        self.custom_listbox(row=3, column=0, list_values=my_test_list)

        self.custom_combo_box(row=3, column=2, width=15, height=5, list_values=self.import_list())

        self.custom_listbox(row=3, column=3)

        # Row 4
        self.custom_button(
            text="Add =>",
            row=4,
            column=2,
            width=15,
            height=0,
            sticky="nswe",
            command="",
        )

        # Row 5
        self.custom_button(
            "<= Remove", row=5, column=2, width=15, height=0, sticky="nswe", command=""
        )

        # Row 6
        self.custom_text_label(
            text="Please select file type:", row=6, column=0, width=5, sticky="nswe"
        )

        self.custom_combo_box(row=6, column=1, width=5, height=5)

        self.custom_text_label(
            text="Please enter file name:", row=6, column=2, width=5, sticky="nswe"
        )

        input1 = Entry(self.parent, width=15)
        input1.grid(row=6, column=3, sticky="nswe", padx=3, pady=3)

        self.custom_button(
            text="Export as File", row=6, column=4, width=10, sticky="nswe", command=""
        )


def main():
    root = Tk()
    root.geometry("750x250")
    app = AttendanceKeeperApp(root)
    root.mainloop()


main()

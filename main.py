from tkinter import Tk, Frame, Entry, Button, END, RIGHT, Label, Listbox
from tkinter import ttk


class AttendanceKeeperApp(Frame):
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
            background="red",
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

    def custom_listbox(self, row: int, column: int):
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

    def custom_combo_box(self, row: int, column: int, width: int, height: int):
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
        self.custom_listbox(row=3, column=0)

        self.custom_combo_box(row=3, column=2, width=15, height=5)

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

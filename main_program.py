from tkinter import *
from tkinter import font
import matplotlib.pyplot as plt
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import Workbook

ERROR_FILE_NOT_FOUND = -1
ERROR_PRODUCT_NOT_FOUND = -2
PRODUCTS_FILE_NAME = 'products.xls'


def open_excel_file():
    try:
        wb = open_workbook(PRODUCTS_FILE_NAME)
    except FileNotFoundError:
        return ERROR_FILE_NOT_FOUND
    else:
        return wb


def create_excel_file():
    wb = Workbook()
    wb.add_sheet('Sheet1')
    wb.save(PRODUCTS_FILE_NAME)


root = Tk()
root.title("SAIL PRODUCT INFORMATION")
root.bind("<Button-1>", lambda e:root.destroy())
root.bind("<Key>", lambda e:root.destroy())

photo = PhotoImage(file='./pic.gif')
text1 = Text(root, height = photo.height()/16, width = int(photo.width()/7.5))
text1.image_create(END, image=photo)
text1.pack(side=TOP)

x = font.families()

text2 = Text(root, height = 4, width = int(photo.width()/7.5))
text2.tag_configure('italics', font=(x[7], 12, 'italic'))
text2.tag_configure('big', font=(x[23], 25, 'bold'))
text2.insert(END, '\t SAIL PRODUCT INFORMATION', 'big')
text2.insert(END, '\n\t\t\t\t Press any key to continue...', 'italics')
text2.pack(side=LEFT)

file = open_excel_file()
if file == ERROR_FILE_NOT_FOUND:
    create_excel_file()

root.mainloop()


def get_range_list(param):
    list1 = []
    for x in range(param):
        list1.append(x)
    return list1


def get_index_if_product_exists(product_name1):
    file1 = open_excel_file()
    if file1 == ERROR_FILE_NOT_FOUND:
        return ERROR_FILE_NOT_FOUND
    number_of_products = file1.sheet_by_index(0).nrows
    if number_of_products <= 0:
        return ERROR_PRODUCT_NOT_FOUND
    sheet = file1.sheet_by_index(0)
    for i in range(number_of_products):
        if product_name1.lower() == sheet.cell(i,0).value.lower():
            return i
    return ERROR_PRODUCT_NOT_FOUND


class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.instruction = Label(self, text="Enter Product Name:")
        self.instruction.grid(row=0, column=0, columnspan=2, sticky=W)
        self.product_name = Entry(self, bd=2, width=22)
        self.product_name.grid(row=0, column=1, sticky=W)
        self.product_name.bind("<Return>", lambda e: self.search_product())

        self.search_button = Button(self, text="Search", height=1, width=40, command=self.search_product)
        self.search_button.grid(row=1, column=0, columnspan=2, sticky=W)

        self.text = Text(self, width=36, height=2, wrap=WORD)
        self.text.grid(row=3, column=0, columnspan=2, sticky=W)

        Label(self, height=1, width=40).grid(row=4, column=0, columnspan=2)

        Button(self, text="Product content", command=self.show_product_content, width=18).grid(row=6,
                                                                                               column=0,
                                                                                               sticky=W,
                                                                                               pady=5, padx=5)

        Button(self, text="Product Description", command=self.show_product_description, width=18).grid(row=6, column=1,
                                                                                                       sticky=W, pady=5,
                                                                                                       padx=5)

        Button(self, text="Product Uses", command=self.show_product_uses, width=18).grid(row=7, column=0, sticky=W,
                                                                                         pady=5, padx=5)

        Button(self, text="Enter new product", command=self.enter_new_product, width=18).grid(row=7, column=1, sticky=W,
                                                                                              pady=5, padx=5)

        Button(self, text='Quit', command=quit, width=40).grid(row=8, column=0, sticky=W, columnspan=2, pady=12)

    def display_message(self, message):
        self.text.delete(0.0, END)
        self.text.tag_configure("red",font=( "Arial",11, 'bold'), foreground = "red")
        if message.__contains__("Sorry") or message.__contains__("enter"):
            self.text.insert(0.0, message, "red")
        else:
            self.text.insert(0.0, message)

    def search_product(self):
        product_name1 = self.product_name.get().strip()

        if product_name1 == "":
            self.display_message("Please enter product name!")
        else:
            result = get_index_if_product_exists(product_name1)
            if result == ERROR_FILE_NOT_FOUND:
                self.display_message("Sorry, database error. Kindly restart the app.")
            elif result == ERROR_PRODUCT_NOT_FOUND:
                self.display_message("Sorry, product not in database.")
            else:
                self.display_message("Product Found!")

    def show_product_content(self):
        product_name1 = self.product_name.get().strip()
        if product_name1 == "":
            self.display_message("Please enter product name!")
            return
        else:
            result = get_index_if_product_exists(product_name1)
            if result == ERROR_FILE_NOT_FOUND:
                self.display_message("Sorry, database error. Kindly restart the app.")
                return
            elif result == ERROR_PRODUCT_NOT_FOUND:
                self.display_message("Sorry, product not in database.")
                return

        file1 = open_excel_file()
        sheet = file1.sheet_by_index(0)
        raw_data = sheet.cell(result, 1).value
        if len(raw_data)==0:
            self.display_message("Product contents not found.")
            return
        raw_data = raw_data.split(";")
        materials = []
        quantities = []
        try:
            for i in raw_data:
                j = i.split("=")
                materials.append(j[0].strip())
                quantities.append(float(j[1].strip()))
        except (IndexError, ValueError):
            self.display_message("Wrong data in contents. Please check.")
            return

        x = get_range_list(len(materials))
        plt.axis([-1, len(materials), 0, 100])
        plt.xticks(x, materials)
        plt.plot(x, quantities, 'ro')
        plt.grid(True)

        for i in range(len(quantities)):
            plt.annotate(quantities[i], xy=(i, quantities[i]),
                         xytext=(i+0.05, quantities[i]+ 1)
                         )

        plt.ylabel('Constituent %')
        plt.xlabel('Materials')
        plt.show()

    def show_product_description(self):
        product_name1 = self.product_name.get().strip()
        if product_name1 == "":
            self.display_message("Please enter product name!")
            return
        else:
            result = get_index_if_product_exists(product_name1)
            if result == ERROR_FILE_NOT_FOUND:
                self.display_message("Sorry, database error. Kindly restart the app.")
                return
            elif result == ERROR_PRODUCT_NOT_FOUND:
                self.display_message("Sorry, product not in database.")
                return

        file = open_excel_file()
        sheet = file.sheet_by_index(0)
        raw_data = sheet.cell(result, 2).value
        if len(raw_data)==0:
            self.display_message("Product description not found.")
            return
        lines = raw_data.split("\n")

        class Apps(Frame):

            def __init__(self, master):
                Frame.__init__(self, master)
                self.grid()
                self.create_widgets()

            def create_widgets(self):

                self.text = Text(self, wrap=WORD)
                self.text.grid(row=5, column=0, columnspan=2, sticky=W)

                description = ""
                for i in lines:
                    description += i
                    description += "\n\n"
                print("description = ", description)
                self.text.delete(0.0, END)
                self.text.insert(0.0, description)

        root = Tk()
        root.title("Product Description")
        app = Apps(root)

        root.mainloop()

    def show_product_uses(self):
        product_name1 = self.product_name.get().strip()
        if product_name1 == "":
            self.display_message("Please enter product name!")
            return
        else:
            result = get_index_if_product_exists(product_name1)
            if result == ERROR_FILE_NOT_FOUND:
                self.display_message("Sorry, database error. Kindly restart the app.")
                return
            elif result == ERROR_PRODUCT_NOT_FOUND:
                self.display_message("Sorry, product not in database.")
                return

        file = open_excel_file()
        sheet = file.sheet_by_index(0)
        raw_data = sheet.cell(result, 3).value
        if len(raw_data) == 0:
            self.display_message("Product uses not found.")
            return
        lines = raw_data.split("\n")

        class Apps(Frame):

            def __init__(self, master):
                Frame.__init__(self, master)
                self.grid()
                self.create_widgets()

            def create_widgets(self):

                self.text = Text(self, wrap=WORD)
                self.text.grid(row=5, column=0, columnspan=2, sticky=W)
                uses = ""
                for i in lines:
                    uses += i
                    uses += "\n\n"
                print("uses = ", uses)
                self.text.delete(0.0, END)
                self.text.insert(0.0, uses)

        root = Tk()
        root.title("USES")
        app = Apps(root)

        root.mainloop()

    def enter_new_product(self):

        class Apps(Frame):

            contents_cleared = False

            def __init__(self, master):
                Frame.__init__(self, master)
                self.grid()
                self.create_widgets()

            def show_message(self, message):
                self.message_window.delete(0.0, END)
                self.message_window.tag_configure("red",font=( x[20],12, 'bold'), foreground = "red")
                self.message_window.insert(0.0, message, "red")

            def clear_text(self):
                if not self.contents_cleared:
                    self.contents.delete(0.0, END)
                    self.contents_cleared = True

            def create_widgets(self):
                self.new_product_label = Label(self, text="Product Name:")
                self.new_product_label.grid(row=0, column=0, columnspan=2, sticky=W)
                self.new_product_name = Entry(self, width=25, bd = 2)
                self.new_product_name.grid(row=0, column=1, sticky=W)

                Label(self, height=1, width=40).grid(row=1, column=0, columnspan=2)

                self.contents_label = Label(self, text="Enter the contents (in %):")
                self.contents_label.grid(row=2, column=0, columnspan=2, sticky=W)
                self.contents = Text(self, width = 44, height = 4, wrap = WORD)
                self.contents.grid(row=3, column=0, columnspan=4, sticky=W)
                self.contents.insert(0.0, "Example: iron ore = 12.4; steel ore = 15.2; Cr = 28.5 (Separate the contents"
                                          " by a semi-colon';')")
                self.contents.bind("<Button-1>", lambda e:self.clear_text())
                self.contents.bind("<FocusIn>", lambda e: self.clear_text())

                Label(self, height=1, width=40).grid(row=4, column=0, columnspan=2)

                self.description_label = Label(self, text="Description:")
                self.description_label.grid(row=5, column=0, columnspan=2, sticky=W)
                self.description = Text(self, width = 44, height = 4, wrap = WORD)
                self.description.grid(row=6, column=0, columnspan=4, sticky=W)

                Label(self, height=1, width=40).grid(row=7, column=0, columnspan=2)

                self.uses_label = Label(self, text="Uses:")
                self.uses_label.grid(row=8, column=0, columnspan=2, sticky=W)
                self.uses = Text(self, width = 44, height = 4, wrap = WORD)
                self.uses.grid(row=9, column=0, columnspan=4, sticky=W)

                Label(self, height=1, width=40).grid(row=10, column=0, columnspan=2)

                self.submit_button = Button(self, text="Submit", width=20, command=self.add_data_to_database)
                self.submit_button.grid(row=11, column=1, sticky=W)

                Label(self, height=1, width=40).grid(row=12, column=0, columnspan=2)

                self.message_window = Text(self, width=44, height=2, wrap=WORD)
                self.message_window.grid(row=13, column=0, columnspan=2, sticky=W)

            def add_data_to_database(self):
                product_name = self.new_product_name.get().strip().lower()
                if len(product_name)==0:
                    self.show_message("Enter product name!")
                    return
                if get_index_if_product_exists(product_name)>=0:
                    self.show_message("Product already exists!")
                    return 
                contents = self.contents.get(0.0, END).strip().replace("\n", "")
                if len(contents)==0 or self.contents_cleared is False:
                    self.show_message("Enter product contents!")
                    return
                description = self.description.get(0.0, END).strip()
                if len(description)==0:
                    self.show_message("Enter product description!")
                    return
                uses = self.uses.get(0.0, END).strip()
                if len(uses)==0:
                    self.show_message("Enter product uses!")
                    return

                file = open_excel_file()
                if file == ERROR_FILE_NOT_FOUND:
                    create_excel_file()

                first_empty_row_index = file.sheet_by_index(0).nrows
                wb = copy(file)
                
                s = wb.get_sheet(0)
                s.write(first_empty_row_index,0,product_name)
                s.write(first_empty_row_index,1,contents)
                s.write(first_empty_row_index,2,description)
                s.write(first_empty_row_index,3,uses)
                wb.save(PRODUCTS_FILE_NAME)

                self.show_message("Product added in the database!")

        root = Tk()
        root.title("SAIL PRODUCT INFORMATION")
        app = Apps(root)

        root.mainloop()


root = Tk()
root.title("SAIL PRODUCT INFORMATION")
app = Application(root)

root.mainloop()

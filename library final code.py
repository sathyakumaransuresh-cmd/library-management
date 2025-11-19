import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from datetime import date, timedelta
import pandas as pd

frame = tk.Tk()
frame.title("LIBRARY MANEGMENT SYSTEM")
frame.configure(bg='light cyan')
frame.attributes('-fullscreen', True)

def issuetab():
    issuetab_window = tk.Toplevel(frame)
    issuetab_window.title('ISSUE BOOK')
    issuetab_window.geometry('1300x700')
    issuetab_window.configure(bg='light blue')

    a = tk.Label(issuetab_window, text='lib_no')
    a.place(x=350, y=220)
    lib_id = tk.Entry(issuetab_window, width='30')
    lib_id.place(x=400, y=220)

    b = tk.Label(issuetab_window, text='book_id')
    b.place(x=340, y=280)
    book_id = tk.Entry(issuetab_window, width='30')
    book_id.place(x=400, y=280)

    c = tk.Label(issuetab_window, text="date of issue")
    c.place(x=320, y=340)
    date_i = tk.Entry(issuetab_window, width='30')
    date_i.place(x=400, y=340)
    date_i.insert(0, date.today())

    e = tk.Label(issuetab_window, text='Name')
    e.place(x=600, y=220)
    lib_name = tk.Entry(issuetab_window, width='30')
    lib_name.place(x=680, y=220)

    f = tk.Label(issuetab_window, text='Book Name')
    f.place(x=600, y=280)
    book_name = tk.Entry(issuetab_window, width='30')
    book_name.place(x=680, y=280)

    g = tk.Label(issuetab_window, text='Return Date')
    g.place(x=600, y=340)
    return_date = tk.Entry(issuetab_window, width='30')
    return_date.place(x=680, y=340)

    def return_interval():
        return_date.insert(0, date.today() + timedelta(days=6))

    d = tk.Button(issuetab_window, text='continue', command=return_interval)
    d.place(x=520, y=380)

    def add_entry():
        data_entries = []
        name = lib_name.get()
        book = book_name.get()
        idate = date_i.get()
        rdate = return_date.get()

        if name and book and idate and rdate:
            messagebox.showinfo("Entry Added", "Entry added successfully!")
        else:
            messagebox.showwarning("Incomplete Data", "Please fill in all fields.")

        filepath = 'final_out.xlsx'
        file = openpyxl.load_workbook(filepath)
        entries = file.active
        entries.append([name, book, idate, rdate])
        file.save(filepath)

    h = tk.Button(issuetab_window, text='Add Entry', command=add_entry)
    h.place(x=820, y=380)


def view():
    view_window = tk.Toplevel(frame)
    view_window.title('VIEW')
    view_window.geometry('1300x700')
    view_window.configure(bg='light blue')

    file_path = "E:/project backup/final_out.xlsx"
    df = pd.read_excel(file_path)

    tree1 = ttk.Treeview(view_window)
    columns = list(df.columns)
    tree1["columns"] = columns

    tree1.column("#0", width=100, minwidth=100)
    tree1.heading("#0", text=columns[0], anchor=tk.W)
    tree1.heading("#1", text=columns[1], anchor=tk.W)
    tree1.heading("#2", text=columns[2], anchor=tk.W)
    tree1.heading("#3", text=columns[3], anchor=tk.W)

    for index, row in df.iterrows():
        tree1.insert("", "end", text=row.iloc[0], values=[row.iloc[i] for i in range(1, len(columns))])

    def delete_selected():
        df = pd.read_excel(file_path)
        selected_items = tree1.selection()
        tree1.delete(selected_items)

        df.drop(index=0)
        remaining_items = []
        updated_data = [tree1.item(item, 'values') for item in remaining_items]
        remaining_items.append(updated_data)
        updated_df = pd.DataFrame(updated_data, columns=df.columns)
        updated_df.to_excel('final_out.xlsx', index=False)

    tree1.pack(side=tk.TOP, fill=tk.X, pady=50)

    deletebutton = tk.Button(view_window, height=3, width=20, text='Book recived', command=delete_selected)
    deletebutton.place(x=850, y=450)


photo = tk.PhotoImage(file="E:/project backup/total.png")
label = tk.Label(image=photo, bg='light cyan')
label.place(x=310, y=-50)

photo_issuebutton = tk.PhotoImage(file="E:/project backup/kk.png")
photo_viewbutton = tk.PhotoImage(file="E:/project backup/kk.png")
    
issuebutton = tk.Button(frame, height=180, width=380, image=photo_issuebutton, bg='light cyan', command=issuetab)
viewbutton = tk.Button(frame, height=180, width=380, image=photo_viewbutton, bg='light cyan', command=view)

issuebutton.place(x=250, y=450)
viewbutton.place(x=850, y=450)

print(issuebutton)
print(viewbutton)

frame.mainloop()

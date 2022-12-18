import os
from tkinter import *
from tkinter import messagebox
import docx2pdf
from tkcalendar import Calendar, DateEntry
from docx import Document

root = Tk()
root.title("مؤسسة نوارة الشمال للصيانة والنظافة")
root.geometry('450x500')
root.maxsize(400, 500)
root.minsize(400, 500)
root.configure(bg="#F2F2F0")

frm_Header = Frame(root, bg='#464AA6', height=70)
frm_Header.pack(side=TOP, fill=X)
lbl_Title = Label(frm_Header, text="مؤسسة نوارة الشمال للصيانة والنظافة", fg='#F2F2F0', font=("Arial Bold", 20),
                  bg='#464AA6')
lbl_Title.place(relx=0.5, rely=0.5, anchor=CENTER)

frm_Footer = Frame(root, bg='#464AA6', height=15)
frm_Footer.pack(side=BOTTOM, fill=X)
lbl_CR = Label(frm_Footer, text="An application powered B I K A +201094207117", fg='#F2F2F0', font=("Arial", 8),
               bg='#464AA6')
lbl_CR.place(relx=0.5, rely=0.5, anchor=CENTER)

frm_Center = Frame(root, padx=15, pady=15)
frm_Center.pack(fill=BOTH, expand=1)
frm_Center.grid_rowconfigure((1, 3, 4), weight=1)
frm_Center.grid_columnconfigure(0, weight=1)
lbl_Cli = Label(frm_Center, text=":بيانات العميل",
                fg='#1B208C', font=("Arial Bold", 12))
lbl_Cli.grid(row=0, column=0, sticky=SE)
lbl_Con = Label(frm_Center, text=':تفاصيل العقد',
                fg='#1B208C', font=("Arial Bold", 12))
lbl_Con.grid(row=2, column=0, sticky=SE)

frm_Cli = Frame(frm_Center, highlightthickness=1, highlightbackground='#F2BC79', highlightcolor='#787CBF', pady=10,
                padx=5)
frm_Cli.grid(row=1, column=0, sticky=EW)
frm_Cli.grid_rowconfigure((0, 1), weight=1)
frm_Cli.grid_columnconfigure((0, 1), weight=1)
lbl_cName = Label(frm_Cli, text=":الأسم", fg='#1B208C', font=("Arial", 13))
lbl_cName.grid(row=0, column=1, sticky=E)
lbl_cNum = Label(frm_Cli, text=":السجل التجاري",
                 fg='#1B208C', font=("Arial", 13))
lbl_cNum.grid(row=0, column=0, sticky=E)
txt_cName = Entry(frm_Cli, fg='#1B208C', font=("Arial Bold", 13), highlightthickness=2, highlightcolor='#464AA6',
                  highlightbackground='#787CBF', bg='#F2F2F0', justify=CENTER)
txt_cName.grid(row=1, column=1, sticky=EW)
txt_cNum = Entry(frm_Cli, fg='#1B208C', font=("Arial Bold", 13), highlightthickness=2, highlightcolor='#464AA6',
                 highlightbackground='#787CBF', bg='#F2F2F0', justify=CENTER)
txt_cNum.grid(row=1, column=0, sticky=EW, padx=(0, 2))

frm_Con = Frame(frm_Center, highlightthickness=1, highlightbackground='#F2BC79', highlightcolor='#787CBF', pady=10,
                padx=5)
frm_Con.grid(row=3, column=0, sticky=EW)
frm_Con.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=1)
frm_Con.grid_columnconfigure((0, 1), weight=1)
lbl_sName = Label(frm_Con, text=":نوع الخدمة",
                  fg='#1B208C', font=("Arial", 13))
lbl_sName.grid(row=0, column=0, sticky=E, columnspan=2)
txt_sName = Entry(frm_Con, fg='#1B208C', font=("Arial Bold", 13), highlightthickness=2, highlightcolor='#464AA6',
                  highlightbackground='#787CBF', bg='#F2F2F0', justify=CENTER, width=100)
txt_sName.grid(row=1, column=0, columnspan=2)
lbl_sPrice = Label(frm_Con, text=":سعر الخدمة",
                   fg='#1B208C', font=("Arial", 13))
lbl_sPrice.grid(row=2, column=0, sticky=E, columnspan=2)
txt_sPrice = Entry(frm_Con, fg='#1B208C', font=("Arial Bold", 13), highlightthickness=2, highlightcolor='#464AA6',
                   highlightbackground='#787CBF', bg='#F2F2F0', justify=CENTER, width=100)
txt_sPrice.grid(row=3, column=0, columnspan=2)
lbl_Date1 = Label(frm_Con, text=":تاريخ البدأ",
                  fg='#1B208C', font=("Arial", 13))
lbl_Date1.grid(row=4, column=1, sticky=E)
lbl_Date2 = Label(frm_Con, text=":تاريخ الأنتهاء",
                  fg='#1B208C', font=("Arial", 13))
lbl_Date2.grid(row=4, column=0, sticky=E)
cal1 = DateEntry(frm_Con, fg="#1B208C", year=2022)
cal1.grid(row=5, column=1, sticky=EW)
cal2 = DateEntry(frm_Con, fg="#1B208C", year=2022)
cal2.grid(row=5, column=0, sticky=EW)


def btn_Click():
    if len(txt_sName.get()) == 0 or len(txt_cName.get()) == 0 or len(txt_cNum.get()) == 0 or len(txt_sPrice.get()) == 0:
        messagebox.showerror("حطأ", "يجب ملئ جميع الحقول")
        return

    if not os.path.exists("./doc.docx"):
        messagebox.showerror("حطأ", "ملف الورد الأساسي غير موجود")
        return

    contract_period = (cal2.get_date().year - cal1.get_date().year) * 12 + (
        cal2.get_date().month - cal1.get_date().month)

    if contract_period == 0:
        contract_period = 1

    dic = {"Day": cal1.get_date().day, "Month": cal1.get_date().month, "Year": cal1.get_date().year,
           "Client_Name": txt_cName.get(), "Client_Num": txt_cNum.get(), "sName": txt_sName.get(),
           "Price": txt_sPrice.get(), "Date2": str(cal1.get_date()), "Date3": str(cal2.get_date()),
           "loan": str(int(txt_sPrice.get()) / contract_period)}

    doc = Document('./doc.docx')

    for a in dic:
        for p in doc.paragraphs:
            if p.text.find(a) >= 0:
                inline = p.runs
                for i in range(len(inline)):
                    if a in inline[i].text:
                        text = inline[i].text.replace(a, str(dic[a]))
                        inline[i].text = text

    name = txt_cName.get() + "_" + str(cal1.get_date().month) + "_" + \
        str(cal1.get_date().day) + "_" + str(cal1.get_date().year)
    doc.save(f'./Contracts/{name}.docx')
    docx2pdf.convert(f'./Contracts/{name}.docx', f'./Contracts/{name}.pdf')

    if os.path.exists(f'./Contracts/{name}.docx') and os.path.exists(f'./Contracts/{name}.pdf'):
        messagebox.showinfo("نجح", "تم إصدار العقد بنجاح..")
        txt_cName.delete(0, END)
        txt_sPrice.delete(0, END)
        txt_cNum.delete(0, END)
        txt_sName.delete(0, END)
        txt_cName.focus()

    else:
        messagebox.showerror("خطأ", "حدث خطأ ما")


btn = Button(frm_Center, text="طباعة العقد", bg="#F2BC79", fg="#1B208C", borderwidth=0,
             width=15, height=1, font=("Arial", 16), command=btn_Click)
btn.grid(row=4, column=0)

txt_cName.focus()

print(os.getcwd())

root.mainloop()

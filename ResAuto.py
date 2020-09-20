from tkinter import *
from tkinter.filedialog import askopenfilename as Open, asksaveasfilename as SaveAs
from tkinter import messagebox as mb
from tkinter.ttk import Progressbar
import requests as r
from selenium import webdriver
from time import *
from xlsxwriter import *
import os


def roll_from_to(frm, to):
    k1 = frm
    l = list(k1)
    k2 = to
    s1 = list(k2)
    m = []
    m.append(k1)
    c = k1
    while(c != k2):
        if(((l[-2]) == "9") and (l[-1] == "9") or (l[-2].isdigit() == False)):
            if(l[-2].isdigit()):
                l[-2], l[-1] = "a", "0"
                c = "".join(l)
                m.append(c)
            else:
                if(l[-1] == "9"):
                    ed = ord(l[-2])+1
                    l[-2], l[-1] = chr(ed), "0"
                    c = "".join(l)
                    m.append(c)
                else:
                    l[-1] = str(int(l[-1])+1)
                    c = "".join(l)
                    m.append(c)
        else:
            if(int(l[-1]) != 9):
                l[-1] = str(int(l[-1])+1)
                c = "".join(l)
                m.append(c)
            else:
                l[-2] = str(int(l[-2])+1)
                l[-1] = "0"
                c = "".join(l)
                m.append(c)
    return m

def submit():
    global name, sgpas
    w1.iconify()
    # Window 3 Single User

    def sing():
        w2.iconify()
        w3 = Toplevel()
        w3.geometry("1080x720")
        w3.configure(bg="#3b5998")
        Label(w3, text="Search for Single Student", bg="#3b5998",fg="orange", font=("Calibri", 25, "bold")).place(x=20, y=20)
        Label(w3, text="Enter URL :", bg="#3b5998", fg="white",font=("Calibri", 20, "bold italic")).place(x=250, y=90)
        E30 = Entry(w3, fg="#000", bg="white", font=("Times New Roman", 14, "bold"),width=45)
        E30.place(x=400, y=100)
        Label(w3, text="Enter Reg. No :", bg="#3b5998", fg="white",font=("Calibri", 20, "bold")).place(x=250, y=140)
        E31 = Entry(w3, fg="#000", bg="white", font=("Times New Roman", 14, "bold italic"))
        E31.place(x=475, y=150)

        def ssing():
            w3.withdraw()
            c = webdriver.Chrome("chromedriver.exe")
            url = E30.get()
            c.get(url)
            ss = c.find_element_by_id('ht')
            ss.clear()
            ss.send_keys(E31.get())
            x = c.find_element_by_class_name('ci')
            x.click()
            while(1):
                if(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr') != []):break
            rows, cols = len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr')), len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr[1]/th'))
            l = []
            for i in range(2, rows):
                t = []
                for j in range(1, cols+1):
                    t.append(c.find_element_by_xpath('//*[@id="rs"]/table/tbody/tr['+str(i)+']/td['+str(j)+']').text)
                l.append(t)
            c.quit()
            os.system("taskkill /f cmd.exe")
            w3.deiconify()
            y, cr_sum, sum_grades, no_of_fails = 250, 0, 0, 0
            for i in l:
                cr_sum += int(i[3])
                if(i[2] == "O"):sum_grades = sum_grades+10*int(i[3])
                elif(i[2] == "S"):sum_grades = sum_grades+9*int(i[3])
                elif(i[2] == "A"):sum_grades = sum_grades+8*int(i[3])
                elif(i[2] == "B"):sum_grades = sum_grades+7*int(i[3])
                elif(i[2] == "C"):sum_grades = sum_grades+6*int(i[3])
                elif(i[2] == "D"):sum_grades = sum_grades+5*int(i[3])
                elif(i[2] == "F"):no_of_fails += 1
                Label(w3, text=i[0], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=100, y=y)
                Label(w3, text=i[1], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=250, y=y)
                Label(w3, text=i[2], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=750, y=y)
                Label(w3, text=i[3], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=973, y=y)
                y += 25
            if(no_of_fails <= 0):
                Label(w3, text="Your SGPA is : %.3f" % (sum_grades/cr_sum), bg="#3b5998",fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=600)
                Label(w3, text="Subject Code", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=100, y=200)
                Label(w3, text="Subject Name", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=250, y=200)
                Label(w3, text="Grade/Status", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=750, y=200)
                Label(w3, text="Credits", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=950, y=200)
                try:
                    Label(w3, text="All Clear...!", bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold")).place(x=700, y=550)
                    Label(w3, text="Your SGPA is : %.3f" % (sum_grades/cr_sum), bg="#3b5998",fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=600)
                except Exception:
                    mb.showerror("Oops :(", "Session Timed Out...")
            else:
                Label(w3, text="You have Failed!", bg="#3b5998", fg="red",font=("Calibri", 25, "bold")).place(x=700, y=600)
                Label(w3, text="So SGPA cannot be calculated... :(", bg="#3b5998",fg="#ff4500", font=("Calibri", 16, "bold")).place(x=775, y=650)
        Button(w3, text="Get Result", bg="grey", fg="#00ff00", font=("Times New Roman", 12, "bold"), command=ssing).place(x=725, y=145)
        w3.mainloop()
    # Window 4 Multi User

    def muti():
        global sgpas, path
        w2.iconify()
        def WEB():
            w4.iconify()
            def browse():
                global sgpas, path
                w5.iconify()
                def GET():
                    global sgpas, path
                    def EXCEL():
                        global sgpas, path
                        def close():
                            w7.destroy()
                            w5.deiconify()
                        def Open():
                            global path
                            w7.destroy()
                            os.startfile(path.get())
                        def Save_As():
                            global sgpas
                            path.set(SaveAs(title="Save File", initialdir="C:/", filetypes=(("Excel Workbook", "*.xlsx"), ("All Files", "*.*")))+".xlsx")
                            if(path.get() != ""):
                                XL_File = Workbook("%s" % (path.get()))
                                keys = list(sgpas.keys())
                                s1 = XL_File.add_worksheet("%s to %s" % (keys[0], keys[-1]))
                                s1.write("A1", "Roll Num")
                                s1.write("B1", "Result")
                                roll_num, res = [], []
                                for i in keys:
                                    roll_num.append(i)
                                    x = XL_File.add_worksheet("%s" % (i))
                                    x.write("C1", i)
                                    x.write("A2", "Subject Code")
                                    x.write("B2", "Subject Name")
                                    x.write("C2", "Grade / Status")
                                    x.write("D2", "Credits")
                                    sgpa, fails = 0, 0
                                    gra, crs = [], []
                                    for j in range(len(sgpas[i])):
                                        l = ["A", "B", "C", "D"]
                                        for k in sgpas[i][j]:
                                            ind = l[sgpas[i][j].index(k)]
                                            x.write("%s%i" %(ind, j+3), "%s" % (k))
                                            if(l.index(ind) == 2):
                                                if(k == "O"):gra.append(10)
                                                elif(k == "S"):gra.append(9)
                                                elif(k == "A"):gra.append(8)
                                                elif(k == "B"):gra.append(7)
                                                elif(k == "C"):gra.append(6)
                                                elif(k == "D"):gra.append(5)
                                                elif(k == "F"):fails += 1
                                                else:gra.append(0)
                                            if(l.index(ind) == 3):crs.append(int(k))
                                    for z in range(len(gra)):sgpa += gra[z]*crs[z]
                                    try:
                                        if(fails <= 0):
                                            Y = sgpa/sum(crs)
                                            res.append(Y)
                                            x.write("C%i" % (j+4), "SGPA:")
                                            x.write("D%i" %(j+4), "%.3f" % (Y))
                                        else:
                                            x.write("D%i" % (j+4), "Failed!")
                                            res.append("Failed!")
                                    except Exception:pass
                                    for i in range(len(roll_num)):
                                        s1.write("A%i" % (i+2), roll_num[i])
                                        s1.write("B%i" % (i+2), res[i])
                                XL_File.close()
                                mb.showinfo("Success :)","File Saved Successfully...")
                            else:path.set("Please Try Again...")
                        w7 = Tk()
                        w7.geometry("400x350")
                        w7.config(bg="#3b5998")
                        Label(w7, text="PATH", bg="#3b5998", fg="white",font=("Calibri", 12, "bold")).place(x=10, y=120)
                        Label(w7, textvariable=path, bg="#3b5998", fg="white", font=("Calibri", 12, "bold")).place(x=100, y=120)
                        Button(w7, text="Save As", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=Save_As).place(x=75, y=200)
                        Button(w7, text="Open File", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=Open).place(x=200, y=200)
                        Button(w7, text="Close", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=close).place(x=150, y=250)
                    def PDF():
                        pass
                    global no_of_fails, S1, S2, S3, S4
                    try:
                        if(no_of_fails == 0):
                            S1.set("")
                            S2.set("")
                        else:
                           S3.set("")
                           S4.set("")
                    except Exception as e:
                        print(e)
                    l = sgpas[selection.get()]
                    y = 325
                    Label(w5, text="Subject Code", bg="#3b5998", fg="#00ff00", font=("Calibri", 18, "bold")).place(x=100, y=275)
                    Label(w5, text="Subject Name", bg="#3b5998", fg="#00ff00", font=("Calibri", 18, "bold")).place(x=250, y=275)
                    Label(w5, text="Grade/Status", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=750, y=275)
                    Label(w5, text="Credits", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=950, y=275)
                    Button(w5, text="Save As PDF", bg="#FF3500", fg="white", command=PDF, font=("Calibri", 14, "bold")).place(x=100, y=625)
                    Button(w5, text="Save As Excel", bg="#1d9f42", fg="white", comman=EXCEL, font=("Calibri", 14, "bold")).place(x=300, y=625)
                    cr_sum, sum_grades, no_of_fails = 0, 0, 0
                    for i in l:
                        cr_sum += int(i[3])
                        if(i[2] == "O"):sum_grades = sum_grades+10*int(i[3])
                        elif(i[2] == "S"):sum_grades = sum_grades+9*int(i[3])
                        elif(i[2] == "A"):sum_grades = sum_grades+8*int(i[3])
                        elif(i[2] == "B"):sum_grades = sum_grades+7*int(i[3])
                        elif(i[2] == "C"):sum_grades = sum_grades+6*int(i[3])
                        elif(i[2] == "D"):sum_grades = sum_grades+5*int(i[3])
                        elif(i[2] == "F"):no_of_fails += 1
                        Label(w5, text=i[0], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=100, y=y)
                        Label(w5, text=i[1], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=250, y=y)
                        Label(w5, text=i[2], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=750, y=y)
                        Label(w5, text=i[3], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=973, y=y)
                        y += 25
                    if(no_of_fails == 0):
                        try:
                            S1.set("All Clear...!")
                            S2.set("Your SGPA is : %.3f" % (sum_grades/cr_sum))
                            Label(w5, textvariable=S3, bg="#3b5998", fg="red", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w5, textvariable=S4, bg="#3b5998", fg="#ff4500", font=("Calibri", 15, "bold")).place(x=775, y=625)
                            Label(w5, textvariable=S1, bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w5, textvariable=S2, bg="#3b5998", fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=625)
                        except Exception:pass
                    else:
                        try:
                            S3.set("You have Failed!")
                            S4.set("So SGPA cannot be calculated... :(")
                            Label(w5, textvariable=S1, bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w5, textvariable=S2, bg="#3b5998", fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=625)
                            Label(w5, textvariable=S3, bg="#3b5998", fg="red", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w5, textvariable=S4, bg="#3b5998", fg="#ff4500", font=("Calibri", 15, "bold")).place(x=775, y=625)
                        except Exception:pass
                frm, to = int(E41.get()[-3:]), int(E42.get()[-3:])
                total = to-frm
                m = roll_from_to(E41.get().lower(), E42.get().lower())
                w8 = Tk()
                w8.config(bg='#3b5998')
                w8.geometry("400x400")
                pgbar = Progressbar(w8, length=250, orient="horizontal", maximum=total, value=10)
                pgbar.place(x=100, y=300)
                sgpas = {}
                c = webdriver.Chrome(executable_path="chromedriver.exe")
                url = E40.get()
                c.get(url)
                to_assure = []
                w4.withdraw()
                for rollno in m:
                    pgbar['value']=int(rollno[-3:])-frm
                    Label(w8, text=str(rollno)).pack()
                    w8.update_idletasks()
                    ss = c.find_element_by_id('ht')
                    ss.clear()
                    ss.send_keys(rollno)
                    x = c.find_element_by_class_name('ci')
                    x.click()
                    while(1):
                        if(c.find_element_by_id('rs').text[16:26] == rollno):break
                    rows, cols = len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr')), len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr[1]/th'))
                    if(rows != 0 and cols != 0):
                        to_assure.append(True)
                        l = []
                        for i in range(2, rows):
                            t = []
                            for j in range(1, cols+1):
                                x = c.find_element_by_xpath('//*[@id="rs"]/table/tbody/tr['+str(i)+']/td['+str(j)+']').text
                                t.append(x)
                            l.append(t)
                        sgpas[rollno] = l
                    else:to_assure.append(False)
                if(all(to_assure)):pass
                else:mb.showerror("Oops :(", "Session Timed Out for atleast One Student")
                c.quit()
                os.system("taskkill /f cmd.exe")
                w5.deiconify()
                mb.showinfo("Success", "Data Successfully Read :)")
                selection = StringVar()
                selection.set("--select--")
                Label(w5, text="Select a Reg. Number", bg="#3b5998", fg="white", font=("Calibri", 18, "bold")).place(x=250, y=210)
                lst=list(sgpas.keys())
                drop_down = OptionMenu(w5, selection,*lst)
                drop_down.place(x=470, y=210)
                Button(w5, text="View", command=GET, bg="white", fg="#3b5998", font=("Calibri", 14, "bold")).place(x=610, y=205)

            # Window 5
            w5 = Toplevel()
            w5.geometry("1080x720")
            w5.config(bg="#3b5998")
            Label(w5, text="Fetch from Internet", bg="#3b5998", fg="#ffff00", font=("calibri", 25, "bold italic")).place(x=10, y=10)
            Label(w5, text="Enter URL :", bg="#3b5998", fg="white",font=("Calibri", 20, "bold italic")).place(x=250, y=60)
            E40 = Entry(w5, fg="#000", bg="white", font=("Times New Roman", 14, "bold"),width=45)
            E40.place(x=400, y=70)
            Label(w5, text="Enter Reg. No's", bg="#3b5998", fg="white",font=("Calibri", 20, "bold italic")).place(x=40, y=125)
            Label(w5, text="From: ", bg="#3b5998", fg="white",font=("Calibri", 20, "bold")).place(x=225, y=125)
            Label(w5, text="To: ", bg="#3b5998", fg="white",font=("Calibri", 20, "bold")).place(x=535, y=125)
            E41 = Entry(w5, fg="#000", bg="white", font=("Times New Roman", 14, "bold"))
            E41.place(x=305, y=135)
            E42 = Entry(w5, fg="#000", bg="white", font=("Times New Roman", 14, "bold"))
            E42.place(x=590, y=135)
            Button(w5, text="Fetch Result", bg="white", fg="#3b5998", font=("Calibri", 14, "bold"), command=browse).place(x=850, y=130)

        def from_PDF():
            def browse():
                global sgpas, path
                w6.iconify()
                def GET():
                    global sgpas, path
                    def EXCEL():
                        global sgpas, path
                        def close():
                            w7.destroy()
                            w6.deiconify()
                        def Open():
                            global path
                            w7.destroy()
                            os.startfile(path.get())
                        def Save_As():
                            global sgpas
                            path.set(SaveAs(title="Save File", initialdir="C:/", filetypes=(("Excel Workbook", "*.xlsx"), ("All Files", "*.*")))+".xlsx")
                            if(path.get() != ""):
                                XL_File = Workbook("%s" % (path.get()))
                                keys = list(sgpas.keys())
                                s1 = XL_File.add_worksheet("%s to %s" % (keys[0], keys[-1]))
                                s1.write("A1", "Roll Num")
                                s1.write("B1", "Result")
                                roll_num, res = [], []
                                for i in keys:
                                    roll_num.append(i)
                                    x = XL_File.add_worksheet("%s" % (i))
                                    x.write("C1", i)
                                    x.write("A2", "Subject Code")
                                    x.write("B2", "Subject Name")
                                    x.write("C2", "Grade / Status")
                                    x.write("D2", "Credits")
                                    sgpa, fails = 0, 0
                                    gra, crs = [], []
                                    for j in range(len(sgpas[i])):
                                        l = ["A", "B", "C", "D"]
                                        for k in sgpas[i][j]:
                                            ind = l[sgpas[i][j].index(k)]
                                            x.write("%s%i" %(ind, j+3), "%s" % (k))
                                            if(l.index(ind) == 2):
                                                if(k == "O"):gra.append(10)
                                                elif(k == "S"):gra.append(9)
                                                elif(k == "A"):gra.append(8)
                                                elif(k == "B"):gra.append(7)
                                                elif(k == "C"):gra.append(6)
                                                elif(k == "D"):gra.append(5)
                                                elif(k == "F"):fails += 1
                                                else:gra.append(0)
                                            if(l.index(ind) == 3):crs.append(int(k))
                                    for z in range(len(gra)):sgpa += gra[z]*crs[z]
                                    try:
                                        if(fails <= 0):
                                            Y = sgpa/sum(crs)
                                            res.append(Y)
                                            x.write("C%i" % (j+4), "SGPA:")
                                            x.write("D%i" %(j+4), "%.3f" % (Y))
                                        else:
                                            x.write("D%i" % (j+4), "Failed!")
                                            res.append("Failed!")
                                    except Exception:pass
                                    for i in range(len(roll_num)):
                                        s1.write("A%i" % (i+2), roll_num[i])
                                        s1.write("B%i" % (i+2), res[i])
                                XL_File.close()
                                mb.showinfo("Success :)","File Saved Successfully...")
                            else:path.set("Please Try Again...")
                        w7 = Tk()
                        w7.geometry("400x350")
                        w7.config(bg="#3b5998")
                        Label(w7, text="PATH", bg="#3b5998", fg="white",font=("Calibri", 12, "bold")).place(x=10, y=120)
                        Label(w7, textvariable=path, bg="#3b5998", fg="white", font=("Calibri", 12, "bold")).place(x=100, y=120)
                        Button(w7, text="Save As", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=Save_As).place(x=75, y=200)
                        Button(w7, text="Open File", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=Open).place(x=200, y=200)
                        Button(w7, text="Close", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=close).place(x=150, y=250)
                    def PDF():
                        pass
                    global no_of_fails, S1, S2, S3, S4
                    try:
                        if(no_of_fails == 0):
                            S1.set("")
                            S2.set("")
                        else:
                           S3.set("")
                           S4.set("")
                    except Exception as e:
                        print(e)
                    l = sgpas[selection.get()]
                    y = 325
                    Label(w6, text="Subject Code", bg="#3b5998", fg="#00ff00", font=("Calibri", 18, "bold")).place(x=100, y=275)
                    Label(w6, text="Subject Name", bg="#3b5998", fg="#00ff00", font=("Calibri", 18, "bold")).place(x=250, y=275)
                    Label(w6, text="Grade/Status", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=750, y=275)
                    Label(w6, text="Credits", bg="#3b5998", fg="#00ff00",font=("Calibri", 18, "bold")).place(x=950, y=275)
                    Button(w6, text="Save As PDF", bg="#FF3500", fg="white", command=PDF, font=("Calibri", 14, "bold")).place(x=100, y=625)
                    Button(w6, text="Save As Excel", bg="#1d9f42", fg="white", comman=EXCEL, font=("Calibri", 14, "bold")).place(x=300, y=625)
                    cr_sum, sum_grades, no_of_fails = 0, 0, 0
                    for i in l:
                        cr_sum += int(i[3])
                        if(i[2] == "O"):sum_grades = sum_grades+10*int(i[3])
                        elif(i[2] == "S"):sum_grades = sum_grades+9*int(i[3])
                        elif(i[2] == "A"):sum_grades = sum_grades+8*int(i[3])
                        elif(i[2] == "B"):sum_grades = sum_grades+7*int(i[3])
                        elif(i[2] == "C"):sum_grades = sum_grades+6*int(i[3])
                        elif(i[2] == "D"):sum_grades = sum_grades+5*int(i[3])
                        elif(i[2] == "F"):no_of_fails += 1
                        Label(w6, text=i[0], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=100, y=y)
                        Label(w6, text=i[1], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=250, y=y)
                        Label(w6, text=i[2], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=750, y=y)
                        Label(w6, text=i[3], bg="#3b5998", fg="white", font=("Calibri", 14, "bold")).place(x=973, y=y)
                        y += 25
                    if(no_of_fails == 0):
                        try:
                            S1.set("All Clear...!")
                            S2.set("Your SGPA is : %.3f" % (sum_grades/cr_sum))
                            Label(w6, textvariable=S3, bg="#3b5998", fg="red", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w6, textvariable=S4, bg="#3b5998", fg="#ff4500", font=("Calibri", 15, "bold")).place(x=775, y=625)
                            Label(w6, textvariable=S1, bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w6, textvariable=S2, bg="#3b5998", fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=625)
                        except Exception:pass
                    else:
                        try:
                            S3.set("You have Failed!")
                            S4.set("So SGPA cannot be calculated... :(")
                            Label(w6, textvariable=S1, bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w6, textvariable=S2, bg="#3b5998", fg="yellow", font=("Calibri", 20, "bold")).place(x=750, y=625)
                            Label(w6, textvariable=S3, bg="#3b5998", fg="red", font=("Calibri", 25, "bold")).place(x=700, y=575)
                            Label(w6, textvariable=S4, bg="#3b5998", fg="#ff4500", font=("Calibri", 15, "bold")).place(x=775, y=625)
                        except Exception:pass
                frm, to = int(E41.get()[-3:]), int(E42.get()[-3:])
                total = to-frm
                m = roll_from_to(E41.get().lower(), E42.get().lower())
                w8 = Tk()
                w8.config(bg='#3b5998')
                w8.geometry("400x400")
                pgbar = Progressbar(w8, length=250, orient="horizontal", maximum=total, value=10)
                pgbar.place(x=100, y=300)
                sgpas = {}
                c = webdriver.Chrome(executable_path="chromedriver.exe")
                url = "https://jntukresults.edu.in/view-results-56736077.html"
                c.get(url)
                to_assure = []
                w4.withdraw()
                for rollno in m:
                    pgbar['value']=int(rollno[-3:])-frm
                    Label(w8, text=str(rollno)).pack()
                    w8.update_idletasks()
                    ss = c.find_element_by_id('ht')
                    ss.clear()
                    ss.send_keys(rollno)
                    x = c.find_element_by_class_name('ci')
                    x.click()
                    while(1):
                        if(c.find_element_by_id('rs').text[16:26] == rollno):break
                    rows, cols = len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr')), len(c.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr[1]/th'))
                    if(rows != 0 and cols != 0):
                        to_assure.append(True)
                        l = []
                        for i in range(2, rows):
                            t = []
                            for j in range(1, cols+1):
                                x = c.find_element_by_xpath('//*[@id="rs"]/table/tbody/tr['+str(i)+']/td['+str(j)+']').text
                                t.append(x)
                            l.append(t)
                        sgpas[rollno] = l
                    else:to_assure.append(False)
                if(all(to_assure)):pass
                else:mb.showerror("Oops :(", "Session Timed Out for atleast One Student")
                c.quit()
                os.system("taskkill /f cmd.exe")
                w6.deiconify()
                mb.showinfo("Success", "Data Successfully Read :)")
                selection = StringVar()
                selection.set("--select--")
                Label(w6, text="Select a Reg. Number", bg="#3b5998", fg="white", font=("Calibri", 18, "bold")).place(x=250, y=210)
                lst=list(sgpas.keys())
                drop_down = OptionMenu(w6, selection,*lst)
                drop_down.place(x=470, y=210)
                Button(w6, text="View", command=GET, bg="white", fg="#3b5998", font=("Calibri", 14, "bold")).place(x=610, y=205)
            # Window 6
            w6 = Toplevel()
            w6.geometry("1080x720")
            w6.config(bg="#3b5998")
            PDF_Path=StringVar()
            PDF_Path.set("")
            Label(w6, text="Fetch from PDF", bg="#3b5998", fg="orange",font=("Calibri", 25, "bold italic")).place(x=10, y=15)
            Label(w6, textvariable=PDF_Path, bg="#3b5998", fg="#000",font=("Calibri", 18, "bold italic")).place(x=360, y=70)
            Label(w6, text="Enter Reg. No's", bg="#3b5998", fg="white",font=("Calibri", 20, "bold italic")).place(x=40, y=125)
            Label(w6, text="From: ", bg="#3b5998", fg="white",font=("Calibri", 20, "bold")).place(x=225, y=125)
            Label(w6, text="To: ", bg="#3b5998", fg="white",font=("Calibri", 20, "bold")).place(x=535, y=125)
            E41 = Entry(w6, fg="#000", bg="white", font=("Times New Roman", 14, "bold"))
            E41.place(x=305, y=135)
            E42 = Entry(w6, fg="#000", bg="white", font=("Times New Roman", 14, "bold"))
            E42.place(x=590, y=135)
            Button(w6, text="Fetch Result", bg="white", fg="#3b5998", font=("Calibri", 14, "bold"), command=browse).place(x=850, y=130)
            def choose():
                if(PDF_Path!=""):
                    PDF_Path.set(Open(title="Choose File", initialdir="C:/", filetypes=(("PDF File", "*.pdf"), ("All Files", "*.*"))))
                else:
                    PDF_Path.set("Please Choose PDF File Again...:(")
            Button(w6, text="Choose PDF", bg="white", fg="#3b5998", font=("Calibri", 14, "bold"), command=choose).place(x=250, y=70)


        # Window 4
        w4 = Toplevel()
        w4.geometry("1080x720")
        w4.configure(bg="#3b5998")
        Label(w4, text="Search for Multiple Students", bg="#3b5998",fg="orange", font=("Calibri", 25, "bold italic")).place(x=20, y=20)
        pdf = PhotoImage(file="PDF.png")
        web = PhotoImage(file="web.png")
        Label(w4, text="Read from a PDF File", fg="white", bg="#3b5998",font=("Calibri", 20, "bold")).place(x=150, y=150)
        Label(w4, text="Browse from Internet", fg="white", bg="#3b5998",font=("Calibri", 20, "bold")).place(x=665, y=150)
        Button(w4, text="Fetch from PDF", bg="white", fg="#3b5998", font=("Calibri", 16, "bold"), command=from_PDF).place(x=200, y=600)
        Button(w4, text="Fetch from Internet", bg="white", fg="#3b5998",font=("Calibri", 16, "bold"), command=WEB).place(x=700, y=600)
        Label(w4, image=pdf, bg="#3b5998").place(x=150, y=275)
        Label(w4, image=web, bg="#3b5998").place(x=675, y=275)
        w4.mainloop()

    # Window 2 Main Frame
    w2 = Toplevel()
    w2.title("Counsellor Dashboard")
    w2.geometry("1080x720")
    w2.configure(bg="#3b5998")
    result = PhotoImage(file="result.png")
    Label(w2, text="Welcome", bg="#3b5998", fg="#00ff55",font=("Calibri", 20, "bold italic")).place(x=20, y=20)
    if(name.get() != ""):
        Label(w2, text="Mr. / Mrs. "+name.get().title(), bg="#3b5998",fg="white", font=("Calibri", 20, "bold italic")).place(x=130, y=20)
    else:
        Label(w2, text="Mr. / Mrs. Anonymous", bg="#3b5998", fg="white",font=("Calibri", 20, "bold italic")).place(x=130, y=20)
    Label(w2, text="Search for Single Student :", bg="#3b5998",fg="white", font=("Calibri", 20, "bold")).place(x=600, y=170)
    Button(w2, text="Click Here", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=sing).place(x=900, y=175)
    Label(w2, text="OR", bg="#3b5998", fg="#00ff00", font=("Calibri", 25, "bold italic")).place(x=750, y=285)
    # Result Image
    Label(w2, image=result, bg="#3b5998").place(x=75, y=180)
    Label(w2, text="Search for Multiple Students :", bg="#3b5998",fg="white", font=("Calibri", 20, "bold")).place(x=600, y=400)
    Button(w2, text="Click Here", bg="white", fg="#3b5998", font=("Calibri", 12, "bold"), command=muti).place(x=940, y=405)
    w2.mainloop()

# Window 1
w1 = Tk()
w1.title("Counsellor Login")
w1.geometry("1080x720")
w1.config(bg="#3b5998")
name, no_of_fails, S1, S2, S3, S4, sgpas, path = StringVar(), 0, StringVar(), StringVar(), StringVar(), StringVar(), {}, StringVar()
logo = PhotoImage(file="logo.png")
Label(w1, text="Enter Counsellor's Name :", fg="white",bg="#3b5998", font=("Calibri", 20, "bold")).place(x=250, y=100)
Entry(w1, textvariable=name, fg="#123456", bg="#ffffff", font=("Times New Roman", 18, "bold italic")).place(x=570, y=105)
Label(w1, image=logo, bg="#3b5998").place(x=150, y=275)
Button(w1, text="Submit", fg="#00ff00", bg="grey", font=("Calibri", 14, "bold"), command=submit).place(x=850, y=100)
w1.mainloop()
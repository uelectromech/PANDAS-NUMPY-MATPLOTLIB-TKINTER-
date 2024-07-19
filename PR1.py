# IMPORTING REQUIRED LIBERIES.
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import random

# OPENING DIALOG FUNCTION TO SELECT EXCELL FILE (*.xlsx; *.xls) AND LOAD
# VALUES TO COMBOBOXES.


def opendialog():

    try:

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx; *.xls")])
        # IN CASE OF SELECTING A FILE.
        if file_path != "":
            label['text'] = file_path
            # CLEARING OLD VALUES.
            combo1.set('')
            combo2.set('')
            label2['text'] = ""
            label3['text'] = ""
            # READ EXCELL FILE CONTENTS AND STORE THEM IN PANDAS DATA FRAME.
            gas = pd.read_excel(file_path)
            # LOADING COLUMN HEADERS INTO COMBOBOX LIST
            combo1['values'] = list(gas.columns)
            combo2['values'] = list(gas.columns)
            # SETTING COMBOBOXES TO READONLY.
            combo1['state'] = 'readonly'
            combo2['state'] = 'readonly'
            # CALLING SELECTION CHANGED FUNCTION ON ComboboxSelected EVENT.
            combo1.bind("<<ComboboxSelected>>", selection_changed1)
            combo2.bind("<<ComboboxSelected>>", selection_changed2)
        # IN CASE OF SELECTING NO FILE.
        else:
            # CLEARING EVERY THING FROM OLD VALUES.
            label['text'] = ""
            combo1.set('')
            combo2.set('')
            combo1['values'] = []
            combo2['values'] = []
            label2['text'] = ""
            label3['text'] = ""

    # EXCEPTION HANDLER TO CATCH ERRORS.
    except BaseException:
        # ON EXCEPTION DO NOTHING.
        pass

#  SELECTION CHANGED FUNCTION ON Combobox1Selected EVENT.


def selection_changed1(event):
    # CLEARING OLD VALUES
    label2['text'] = ""
    # READ EXCELL FILE CONTENTS AND STORE THEM IN PANDAS DATA FRAME.
    gas = pd.read_excel(label['text'])
    # LOADING EXCELL SHEET SELECTED COLUMN IN COMBOBOX SIZE,MIN & MAX.
    label2['text'] = "Count = " + str(gas[combo1.get()].size)
    label2['text'] = label2['text'] + \
        "\nMinimum = " + str(gas[combo1.get()].min())
    label2['text'] = label2['text'] + \
        "\nMaximum = " + str(gas[combo1.get()].max())

#  SELECTION CHANGED FUNCTION ON Combobox2Selected EVENT.


def selection_changed2(event):
    # CLEARING OLD VALUES
    label3['text'] = ""
    # READ EXCELL FILE CONTENTS AND STORE THEM IN PANDAS DATA FRAME.
    gas = pd.read_excel(label['text'])
    # LOADING EXCELL SHEET SELECTED COLUMN IN COMBOBOX SIZE,MIN & MAX.
    label3['text'] = "Count = " + str(gas[combo2.get()].size)
    label3['text'] = label3['text'] + \
        "\nMinimum = " + str(gas[combo2.get()].min())
    label3['text'] = label3['text'] + \
        "\nMaximum = " + str(gas[combo2.get()].max())

# PLOT FUNCTION FOR PLOTTING EXCELL SHEET DATA ON X & Y AXIS.


def plot():

    try:
        # READ EXCELL FILE CONTENTS AND STORE THEM IN PANDAS DATA FRAME.
        gas = pd.read_excel(label['text'])
        # DETERMINING PLOT SIZE.
        plt.figure(figsize=(10, 5))
        # DETERMINING PLOT TITLE LABEL AND FORMATTING.
        plt.title(
            combo2.get(),
            fontdict={
                'fontweight': 'bold',
                'fontsize': 18})
        # PLOTTING DATA
        plt.plot(gas[combo1.get()], gas[combo2.get()],
                 'b.-', label=combo2.get())
        # SCALING X AXIS DATA.
        plt.xticks(gas[combo1.get()][::int(
            gas[combo1.get()].size / 5)].tolist() + [gas[combo1.get()].max()])
        # SETTING X AXIS LABEL.
        plt.xlabel(combo1.get())
        # PRINTING LEGEND FOR THE PLOT DATA.
        plt.legend()
        # SAVING A COPY OF PLOT AS A PICTURE.
        plt.savefig('Gas_price_figure.png', dpi=300)
        # VIEWING THE FINAL PLOT.
        plt.show()
    # EXCEPTION HANDLER TO CATCH ERRORS.
    except BaseException:
        # SHOWING WARNING MESSAGEBOX ON ERRORS.
        messagebox.showwarning(
            "Warning", "Plot function requirments not satisfied")

# ADDING NEW CURVES TO CURRENT PLOT FUNCTION FOR ADDING MORE EXCELL SHEET
# DATA ON X & Y AXIS.


def addtoplot():

    try:
        # READ EXCELL FILE CONTENTS AND STORE THEM IN PANDAS DATA FRAME.
        gas = pd.read_excel(label['text'])
        # CLEARING PLOT TITLE LABEL.
        plt.title('')
        # SETTING RANDOM CURVE COLOR FOR NEW CURVE.
        rndcolor = [
            "#" + ''.join([random.choice('0123456789ABCDEF') for j in range(6)])]
        # PLOTTING DATA
        plt.plot(gas[combo1.get()], gas[combo2.get()],
                 color=rndcolor[0], marker='.', label=combo2.get())
        # SCALING X AXIS DATA.
        plt.xticks(gas[combo1.get()][::int(
            gas[combo1.get()].size / 5)].tolist() + [gas[combo1.get()].max()])
        # SETTING X AXIS LABEL.
        plt.xlabel(combo1.get())
        # PRINTING LEGEND FOR THE PLOT DATA.
        plt.legend()

        # SAVING A COPY OF PLOT AS A PICTURE.
        plt.savefig('Gas_price_figure.png', dpi=300)

        # VIEWING THE FINAL PLOT.
        plt.show()

    # EXCEPTION HANDLER TO CATCH ERRORS.
    except BaseException:
        # SHOWING WARNING MESSAGEBOX ON ERRORS.
        messagebox.showwarning(
            "Warning", "Plot function requirments not satisfied")


# LOADING NEW FORM.
root = tk.Tk()
# SETTING FORM TITLE , DIMENTIONS AND STARTUP LOCATON.
root.title("EXCEL ANALIZER")
root.geometry('1400x850+50+50')
# SETTING FORM RESIZABLE PROPERITY & ICON PICTURE.
root.resizable(True, True)
root.iconbitmap('icon.ico')
# SETTING FORM BACKGROUND PICTURE.
background_image = tk.PhotoImage(file='landscape.png')
background_label = tk.Label(root, image=background_image)
background_label.place(relwidth=1, relheight=1)
# ADDING FRAME.
frame = tk.Frame(root, bg='#64B2FA', bd=5, borderwidth=1, relief="solid")
frame.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.10, anchor='n')
# ADDING LABEL.
label = tk.Label(
    frame,
    bg='#DDEEFA',
    font=(
        "Arial",
        12),
    borderwidth=1,
    relief="solid")
label.place(relx=0.01, rely=0.04, relwidth=0.98, relheight=0.40)
# ADDING BUTTON.
button = tk.Button(
    frame,
    bg='#B0CEE3',
    text="Open excel file",
    font=(
        "Arial",
        18),
    # CALLING OPENDIALOG FUNCTION ON BUTTON CLICKING EVENT.
    command=lambda: opendialog(),
    borderwidth=1,
    relief="solid")
button.place(relx=0.35, rely=0.5, relheight=0.45, relwidth=0.3)
# ADDING NEW FRAME
lower_frame = tk.Frame(
    root,
    bg='#64B2FA',
    bd=10,
    borderwidth=1,
    relief="solid")
lower_frame.place(
    relx=0.5,
    rely=0.25,
    relwidth=0.75,
    relheight=0.6,
    anchor='n')
# ADDING BUTTON.
button1 = tk.Button(
    lower_frame,
    bg='#B0CEE3',
    text="Plot",
    font=(
        "Arial",
        25),
    # CALLING PLOT FUNCTION ON BUTTON CLICKING EVENT.
    command=lambda: plot(),
    borderwidth=1,
    relief="solid")
button1.place(relx=0.25, rely=0.8, relheight=0.15, relwidth=0.225)
# ADDING BUTTON.
button2 = tk.Button(
    lower_frame,
    bg='#B0CEE3',
    text="Add to Plot",
    font=(
        "Arial",
        25),
    # CALLING ADDTOPLOT FUNCTION ON BUTTON CLICKING EVENT.
    command=lambda: addtoplot(),
    borderwidth=1,
    relief="solid")
button2.place(relx=0.525, rely=0.8, relheight=0.15, relwidth=0.25)

# FORMATTING COMBOBOXES.
style = ttk.Style()
style.theme_use('clam')
style.configure("combo1", background='#64B2FA')
style.configure("combo2", background='#64B2FA')
# ADDING COMBOBOX.
combo1 = ttk.Combobox(lower_frame, font=("Arial", 25))
combo1.place(relx=0.15, rely=0.05, relheight=0.10, relwidth=0.2)
# ADDING COMBOBOX.
combo2 = ttk.Combobox(lower_frame, font=("Arial", 25))
combo2.place(relx=0.65, rely=0.05, relheight=0.10, relwidth=0.2)
# ADDING LABEL.
label2 = tk.Label(
    lower_frame,
    bg='#DDEEFA',
    font=(
        "Arial",
        26),
    borderwidth=1,
    relief="solid")
label2.place(relx=0.05, rely=0.2, relwidth=0.4, relheight=0.55)
# ADDING LABEL.
label3 = tk.Label(
    lower_frame,
    bg='#DDEEFA',
    font=(
        "Arial",
        26),
    borderwidth=1,
    relief="solid")
label3.place(relx=0.55, rely=0.2, relwidth=0.4, relheight=0.55)
# FORM MAIN LOOP.
root.mainloop()

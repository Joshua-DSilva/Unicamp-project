from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import make_interp_spline
import ctypes
from openpyxl import Workbook
import os


def open_file():
   # browse file 
   global my_dir, l1, btn_setlim
   my_dir = filedialog.askopenfilename()
   sheet_name = os.path.basename(my_dir)
   l1 = Label(frame, text=sheet_name, bd = 1, font= ('Helvetica', 10), relief=SUNKEN, anchor=W)
   l1.grid(row = 0, column = 1)
   l1.grid_propagate(0)
   # Enable limit and graph button
   btn_setlim = ttk.Button(frame, text="Set Limits", command=set_limits)
   btn_setlim.grid(row = 1, column = 0)
   btn_getgraph['state']=NORMAL


def build_graph(dir):
    # Read directory 
    global file
    l1.destroy()
    global Freq, X_, Y_mi, Y_ma, check_setlim
    dict = pd.read_excel(dir, None)
    file = list(dict.keys())    
    # Data received in first 4 worksheets
    res = []
    for i in range(4):
        if i==0:
            df = dict.get(file[i])
            arr = df.to_numpy()
            arr = arr.T
            Freq = arr[0][0:61]
            X_ = np.linspace(Freq.min(), Freq.max(), 1000000)
            res.append(X_)
            res = res + get_graph(df)
        else:
            res = res + get_graph(dict.get(file[i]))
    # Display graph using Z-Values and Freq values from res
    display_graphs(res)
    plt.show()
    # Disable buttons after graph display
    btn_getgraph['state']=DISABLED
    if check_setlim==1:
        frame_setlim.grid_forget()
        check_setlim = 0
    btn_setlim['state']=DISABLED
    Y_mi = 100
    Y_ma = 10000000


def get_graph(df):
    arr = df.to_numpy()
    arr = arr.T
    Z = arr[2]
    n = int(Val_Replicate.get())
    lst = []
    X_Y_Spline = []
    Y_Val = []
    for i in range(n):
        lst.append(Z[i*61:(i+1)*61])
        X_Y_Spline.append(make_interp_spline(Freq, lst[-1]))
        Y_Val.append(X_Y_Spline[-1](X_))
    Z = np.array(lst)
    # Create multiple points on the graph
    # Returns evenly spaced numbers
    # over a specified interval.
    
    # Y0_ = X_Y_Spline1(X_)
    # Y1_ = X_Y_Spline2(X_)
    # Y2_ = X_Y_Spline3(X_)
    return Y_Val


def display_graphs(res):
    global Y_mi, Y_ma
    if check_setlim==1:
        Y_mi = int(Y_mini.get())
        Y_ma = int(Y_maxi.get())
    fig, axs = plt.subplots(2, 2)
    # For each quadrant
    # Plot Z1, Z2, Z3 
    # Sets the scale in Logarithmic
    # Set axis limits
    n = int(Val_Replicate.get())
    for i in range(n):
        axs[0,0].plot(res[0], res[i+1])
    axs[0,0].set_yscale("log")
    axs[0,0].set_xscale("log")
    axs[0,0].set_xlim(10, 1000000)
    axs[0,0].set_ylim(Y_mi, Y_ma)
    axs[0,0].grid()
    axs[0,0].set_title(file[0])
    for i in range(n):
        axs[0,1].plot(res[0], res[n+(i+1)])
    axs[0,1].set_yscale("log")
    axs[0,1].set_xscale("log")
    axs[0,1].set_xlim(10, 1000000)
    axs[0,1].set_ylim(Y_mi, Y_ma)
    axs[0,1].grid()
    axs[0,1].set_title(file[1])
    for i in range(n):
        axs[1,0].plot(res[0], res[2*n+(i+1)])
    axs[1,0].set_yscale("log")
    axs[1,0].set_xscale("log")
    axs[1,0].set_xlim(10, 1000000)
    axs[1,0].set_ylim(Y_mi, Y_ma)
    axs[1,0].grid()
    axs[1,0].set_title(file[2])
    for i in range(n):
        axs[1,1].plot(res[0], res[3*n+(i+1)])
    axs[1,1].set_yscale("log")
    axs[1,1].set_xscale("log")
    axs[1,1].set_xlim(10, 1000000)
    axs[1,1].set_ylim(Y_mi, Y_ma)
    axs[1,1].grid()
    axs[1,1].set_title(file[3])
    for ax in axs.flat:
        ax.set(xlabel='Frequency', ylabel='Z1,Z2,Z3')
    for ax in fig.get_axes():
        ax.label_outer()


def set_limits():
    # Create new frame and labels within 
    # Set the limit variables to global for mutability
    global Y_mini, Y_maxi, frame_setlim, check_setlim
    frame_setlim = LabelFrame(frame, text="Create Graphs", height = 200,width=200, padx = 5, pady = 5)
    frame_setlim.grid(row = 1, column = 1, columnspan=10,rowspan=6, padx = 2, pady = 2)
    Y_min = Label(frame_setlim, text = "Y_min").grid(row=0,column=0)
    Y_max = Label(frame_setlim, text = "Y_max").grid(row=1,column=0)
    Y_mini = Entry(frame_setlim, width=25)
    Y_mini.insert(0, 100)
    Y_mini.grid(row=0,column=1)
    Y_maxi = Entry(frame_setlim, width=25)
    Y_maxi.insert(0, 10000000)
    Y_maxi.grid(row=1,column=1)
    check_setlim = 1


def set_Freq():
    # Creates new field for input 
    # set global variable when pressed
    global Freq_mini, Freq_maxi, Freq_min, Freq_max, check_setfreq
    Freq_min = Label(frame1, text = "Freq_min")
    Freq_min.grid(row=5,column=10, rowspan=2)
    Freq_max = Label(frame1, text = "Freq_max")
    Freq_max.grid(row=7,column=10, rowspan=2)
    Freq_mini = Entry(frame1, width=5)
    Freq_mini.insert(0, 120)
    Freq_mini.grid(row=5,column=11, rowspan=2)
    Freq_maxi = Entry(frame1, width=5)
    Freq_maxi.insert(0, 1400)
    Freq_maxi.grid(row=7,column=11, rowspan=2)
    check_setfreq = 1


def Add_file():
   global top, btn_clear,btn_create, file_names, file_add, excel_name, btn_setfreq, check_setfreq, msg, msg_check
   dir_file = filedialog.askopenfilename()
   sheet_name = os.path.basename(dir_file) # file name with extension
   excel_name.append(os.path.splitext(sheet_name)[0]) # File name without extension
   file_names.append(dir_file)
   label = Label(frame_file, text=sheet_name, font= ('Helvetica', 10), justify="left")
   # file_add is a pointer array
   file_add.append(id(label)) # Stores the address to label
   ctypes.cast(file_add[top], ctypes.py_object).value.grid(row = top, column = 1) # gets value from address
   # Enable the buttons
   top+=1
   if top == 1:
    btn_clear = ttk.Button(frame1, text="CLEAR", command=Clear_file)
    btn_clear.grid(row = 3, column = 10)
    btn_clear['state']=NORMAL
    btn_setfreq['state']=NORMAL
    btn_create = Button(frame1, text="CREATE", height = 5, width = 10, command=Create_Matrix)
    btn_create.grid(row = 9, column = 10)
    if msg_check==1:
        msg.destroy()
        msg_check = 0


def Clear_file():
    global top, Freq_mini, Freq_maxi, Freq_min, Freq_max, btn_setfreq, check_setfreq
    top-=1
    # pop values from all the stacks
    file_names.pop()
    excel_name.pop()
    ctypes.cast(file_add[top], ctypes.py_object).value.destroy() # Destroy label by obtaining value from the address
    file_add.pop()
    if top==0:
        btn_create['state']=DISABLED
        btn_clear['state']=DISABLED
        btn_setfreq['state']=DISABLED
    if top==0 and check_setfreq==1:
        Freq_min.grid_forget()
        Freq_max.grid_forget()
        Freq_mini.grid_forget()
        Freq_maxi.grid_forget()
        btn_setfreq['state']=DISABLED
        check_setfreq = 0


def Get_Freq(arr, mi, ma):
    arr = arr.T
    Freq = arr[0][0:61]
    a=-1
    b = len(Freq)-1
    for i in range(len(Freq)):
        if Freq[i]>=mi:
            if a!=-1:
                if Freq[i]>=ma:
                    b = i
                    break
                else:
                    continue
            else:
                a = i
        else:
            continue
    return [a,b] # Returns upper and lower index of range


def Get_Z(arr, a, b):
    arr = arr.T
    Z = arr[2]
    n = int(Val_Replicate.get())
    lst = []
    for i in range(n):
        lst.append(Z[i*61:(i+1)*61])
    Z = np.array(lst)
    Z = Z.T
    return Z[a:b+1] # Returns Z Matrix between range


def Get_Matrix(dict,file, fmin, fmax):
    df = dict.get(file[0])
    arr = df.to_numpy()
    for i in range(4):
        df = dict.get(file[i])
        arr = df.to_numpy()
        Z = Get_Z(arr, fmin, fmax)
        if i==0:
            Mat = Z
        else:
            Mat = np.concatenate((Mat,Z),axis=0) 
    return Mat # Returns matrix of entire workbook


def Create_Matrix():
    # Browses to save file
    global top, name, freqmin, freqmax, check_setfreq, msg, msg_check
    dir_file = filedialog.askdirectory()
    if dir_file=="":
        return
    for j in range(len(file_names)+1):
        i = j-1
        if i == -1:
            # Adds the frequency values to the top of the matrix
            dir = file_names[i]
            dict = pd.read_excel(dir, None)
            file = list(dict.keys())
            df = dict.get(file[0])
            arr = df.to_numpy()
            if check_setfreq==1:
                freqmin = int(Freq_mini.get())
                freqmax = int(Freq_maxi.get())
            a, b = Get_Freq(arr, freqmin, freqmax)
            arr = arr.T
            Freq = np.array(arr[0][0:61][a:b+1])
            Mat = np.tile(Freq,4)
            Mat = Mat[np.newaxis, :]
        else:
            # Appends triplicates of each workbook and pops values from stack simultaneously
            dir = file_names[i]
            top-=1
            ctypes.cast(file_add[top], ctypes.py_object).value.destroy()
            file_add.pop()
            dict = pd.read_excel(dir, None)
            file = list(dict.keys())
            Mat = np.concatenate((Mat,(Get_Matrix(dict,file, a, b)).T),axis=0)
    # Adds Header to the database
    # Adds a column providing row header
    # Creates a workbook and upodates values
    column = list(np.repeat(["IDE1","IDE2","IDE3","IDE4"],b-a+1))
    row = [" "]
    row = row + list(np.repeat(excel_name,int(Val_Replicate.get())))
    df = pd.DataFrame (Mat, columns = column)
    df.insert(0, " ", row, True)
    filepath = str(dir_file) + "/" + name.get() + ".xlsx"
    workbook = Workbook()
    workbook.save(filepath) 
    df.to_excel(filepath, index=False)
    top = 0
    file_names.clear()
    name.delete(0,END)
    name.insert(0,"File_Sample")
    excel_name.clear()
    msg = Label(frame1, text = "Matrix Created\U0001F600")
    msg.grid(row = 10, column = 10)
    msg_check = 1
    if top==0 and check_setfreq==1:
        Freq_min.grid_forget()
        Freq_max.grid_forget()
        Freq_mini.grid_forget()
        Freq_maxi.grid_forget()
        check_setfreq = 0
    btn_create['state']=DISABLED
    btn_clear['state']=DISABLED
    btn_setfreq['state']=DISABLED

# Creates a Terminal
root = Tk()
root.title('E-Tongue Sample Analysis')
root.geometry("460x510") # width*Height 
root.wm_attributes('-toolwindow', 'True')

# Initializing variables
excel_name = []
file_names = []
file_add = []
Y_mi = 100
Y_ma = 10000000
freqmin = 120
freqmax = 1400
check_setlim = 0
check_setfreq = 0
msg_check = 0
top = 0

# Create frame for graphing
frame = LabelFrame(root,text="Graph Analysis", height = 700, padx = 5, pady = 5)
frame.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = EW)
root.grid_columnconfigure(0,weight=1)

# Create frame for matrix formation
frame1 = LabelFrame(root, text = "PCA Matrix Formation", height = 380, padx = 5, pady = 5)
frame1.grid(row = 1, column = 0, padx = 5, pady = 5, sticky = 'ew')
root.grid_columnconfigure(0,weight=1)
frame1.grid_propagate(0)

# Create frame for adding file names
frame_file = LabelFrame(frame1, height = 320, width = 300, padx = 2, pady = 2, relief=SUNKEN)
frame_file.grid(row = 0, column = 0, rowspan=10, columnspan=10, padx = 5, pady = 5)
frame_file.grid_propagate(0)

# Create buttons for Matrix Softare
Num_Replicate = Label(frame1, text = "Replicates:").grid(row=1,column=10)
Val_Replicate = Entry(frame1, width=4)
Val_Replicate.grid(row=1,column=11)
Val_Replicate.insert(6,"6")
btn_add = ttk.Button(frame1, text="ADD", command=Add_file).grid(row = 2, column = 10)
btn_clear = ttk.Button(frame1, text="CLEAR", command=Clear_file,state=DISABLED)
btn_setfreq = ttk.Button(frame1, text="SET FREQ", command=set_Freq,state=DISABLED)
btn_clear.grid(row = 3, column = 10)
btn_setfreq.grid(row = 4, column = 10)
btn_create = Button(frame1, text="CREATE", height = 5, width = 10, command=Create_Matrix,state=DISABLED)
btn_create.grid(row = 9, column = 10)

# Create Label to input file save name
file_n = Label(frame1, text = "Save As").grid(row=10,column=0)
name = Entry(frame1, width=42)
name.grid(row=10,column=1)
name.insert(0,"File_Sample")

# Create buttons for Graphing Software
btn_browse = ttk.Button(frame, text="Browse", command=open_file).grid(row = 0, column = 0)
btn_setlim = ttk.Button(frame, text="Set Limits", command=set_limits,state=DISABLED)
btn_setlim.grid(row = 1, column = 0)
btn_getgraph = ttk.Button(frame, text="Get Graphs", command=lambda: build_graph(my_dir),state=DISABLED)
btn_getgraph.grid(row = 2, column = 0)

root.mainloop()
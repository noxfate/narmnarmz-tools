import common
from Inspection_plan import simple
from Validate_Insp import validate
from Validate_TaskList import validateTL
from  Validate_Recipe import validateR
import sys
from Tkinter import *

sys.dont_write_bytecode = True

# validate.run()


def sel():
    getVarB = varB.get()
    selection = "You selected the option " + str(varP.get())
    label.config(text = selection)
    if varP.get() == 1:
        simple.run()
        root.quit()
        root.destroy()
    elif varP.get() == 2:
        Run.pack_forget()
        ValidateInsp()
    elif varP.get() == 3:
        validateTL.run()
        root.quit()
        root.destroy()
    elif varP.get() == 4:
        validateR.run(getVarB)
        root.quit()
        root.destroy() 	

def ValidateInsp():

    SheetFrame = LabelFrame(root, text="Select sheet to validate")
    C1 = Checkbutton(SheetFrame, text = "Header", variable = ws_header, onvalue = 1, offvalue = 0 )
    C1.pack( anchor = W )
    C2 = Checkbutton(SheetFrame, text = "Operation", variable = ws_operation, onvalue = 1, offvalue = 0)
    C2.pack( anchor = W )
    C3 = Checkbutton(SheetFrame, text = "MIC", variable = ws_mic, onvalue = 1, offvalue = 0)
    C3.pack( anchor = W )
    C4 = Checkbutton(SheetFrame, text = "Mat. Assign.", variable = ws_matassign, onvalue = 1, offvalue = 0)
    C4.pack( anchor = W )
    #C5 = Checkbutton(SheetFrame, text = "Dep. Char.", variable = ws_depchar, onvalue = 1, offvalue = 0)
    #C5.pack( anchor = W )
    C6 = Button(SheetFrame, text ="Run", command = runValInsp)
    C6.pack( anchor = CENTER )
    SheetFrame.pack(fill="both", expand="yes")

def runValInsp():
    
    getVarB = varB.get()
    getVarAdd = varAdd.get()

    if ws_header.get() == 1:
        varSheet.append(1)
    if ws_operation.get() == 1:
        varSheet.append(2)
    if ws_mic.get() == 1:
        varSheet.append(3)
    if ws_matassign.get() == 1:
        varSheet.append(4)
    #if ws_depchar.get() == 1:
        #varSheet.append(5)
    validate.run(varSheet,getVarB,getVarAdd)
    root.quit() 
    root.destroy()

root = Tk()
varP = IntVar()
varB = StringVar()
varAdd = BooleanVar()
varSheet = []
ws_header = IntVar()
ws_operation = IntVar()
ws_mic = IntVar()
ws_matassign = IntVar()
ws_depchar = IntVar()

SelectB = LabelFrame(root, text="Select Business")
SelectB.pack(fill="both", expand="yes")

Business = [
    ("Factory", "Factory"),
    ("Farm", "Farm")
]

varB.set("Factory")

for text, mode in Business:
    B = Radiobutton(SelectB, text=text, variable=varB, value=mode)
    B.pack(anchor=W)

SelectP = LabelFrame(root, text="Select Program")
SelectP.pack(fill="both", expand="yes")

Program = [
    ("Convert Inspection Plan", 1),
    ("Validate Inspection Plan", 2),
    ("Validate Task List", 3),
    ("Validate Recipe", 4)
]

for text, mode in Program:
    P = Radiobutton(SelectP, text=text, variable=varP, value=mode)
    P.pack(anchor=W)

varAdd.set(True)
CAdd = Checkbutton(SelectP, text = "MIC Additional Condition", variable = varAdd, onvalue = True, offvalue = False )
CAdd.pack(anchor = W)

Run = Button(SelectP, text ="Run", command = sel)
Run.pack( anchor = CENTER)

'''
R1 = Radiobutton(SelectP, text="Convert Inspection Plan", variable=var, value=1, command=sel)
R1.pack( anchor = W )

R2 = Radiobutton(SelectP, text="Validate Inspection Plan", variable=var, value=2, command=ValidateInsp)
R2.pack( anchor = W )

#R3 = Radiobutton(SelectP, text="Validate Task List", variable=var, value=3, command=sel)
#R3.pack( anchor = W )

R4 = Radiobutton(SelectP, text="Validate Recipe", variable=var, value=4, command=sel)
R4.pack( anchor = W )
'''

label = Label(root)
label.pack()
root.quit()
root.mainloop()
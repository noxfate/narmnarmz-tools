import common
from Inspection_plan import simple
from Validate_Insp import validate
#from Validate_TaskList import validateTL
from  Validate_Recipe import validateR
import sys
from Tkinter import *

sys.dont_write_bytecode = True

# validate.run()


def sel():
    selection = "You selected the option " + str(var.get())
    label.config(text = selection)
    if var.get() == 1:
        root.quit()
        root.destroy()
        simple.run()
    elif var.get() == 2:
        if ws_header.get() == 1:
            varSheet.append(1)
        if ws_operation.get() == 1:
            varSheet.append(2)
        if ws_mic.get() == 1:
            varSheet.append(3)
        if ws_matassign.get() == 1:
            varSheet.append(4)
        if ws_depchar.get() == 1:
            varSheet.append(5)
        validate.run(varSheet)
        root.quit() 
        root.destroy()
    elif var.get() == 3:
        root.quit()
        root.destroy()
    	validateTL.run()
    elif var.get() == 4:
        root.quit()
        validateR.run()
        root.quit()
        root.destroy()
    	
    	

def ValidateInsp():

    label.config(text="Select sheet to validate")
    C1 = Checkbutton(root, text = "Header", variable = ws_header, onvalue = 1, offvalue = 0 )
    C1.pack( anchor = W )
    C2 = Checkbutton(root, text = "Operation", variable = ws_operation, onvalue = 1, offvalue = 0)
    C2.pack( anchor = W )
    C3 = Checkbutton(root, text = "MIC", variable = ws_mic, onvalue = 1, offvalue = 0)
    C3.pack( anchor = W )
    C4 = Checkbutton(root, text = "Mat. Assign.", variable = ws_matassign, onvalue = 1, offvalue = 0)
    C4.pack( anchor = W )
    C5 = Checkbutton(root, text = "Dep. Char.", variable = ws_depchar, onvalue = 1, offvalue = 0)
    C5.pack( anchor = W )
    C6 = Button(root, text ="Run", command = sel)
    C6.pack( anchor = CENTER )

root = Tk()
var = IntVar()
varSheet = []
ws_header = IntVar()
ws_operation = IntVar()
ws_mic = IntVar()
ws_matassign = IntVar()
ws_depchar = IntVar()

R1 = Radiobutton(root, text="Convert Inspection Plan", variable=var, value=1, command=sel)
R1.pack( anchor = W )

R2 = Radiobutton(root, text="Validate Inspection Plan", variable=var, value=2, command=ValidateInsp)
R2.pack( anchor = W )

#R3 = Radiobutton(root, text="Validate Task List", variable=var, value=3, command=sel)
#R3.pack( anchor = W )

R4 = Radiobutton(root, text="Validate Recipe", variable=var, value=4, command=sel)
R4.pack( anchor = W )

label = Label(root)
label.pack()
root.quit()
root.mainloop()
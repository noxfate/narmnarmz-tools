import common
from Inspection_plan import simple
from Validate_Insp import validate
import sys
from Tkinter import *

sys.dont_write_bytecode = True

def sel():
    selection = "You selected the option " + str(var.get())
    label.config(text = selection)
    if var.get() == 1:
        simple.run()
    elif var.get() == 2:
        validate.run()

root = Tk()
var = IntVar()
R1 = Radiobutton(root, text="Run Inspection Plan", variable=var, value=1, command=sel)
R1.pack( anchor = W )

R2 = Radiobutton(root, text="Run Validate Inspection Plan", variable=var, value=2, command=sel)
R2.pack( anchor = W )

label = Label(root)
label.pack()
root.mainloop()
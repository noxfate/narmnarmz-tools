import common
from Inspection_plan import simple
from Validate_Insp import validate
#from Validate_TaskList import validateTL
import sys
from Tkinter import *

sys.dont_write_bytecode = True

# validate.run()

def sel():
    selection = "You selected the option " + str(var.get())
    label.config(text = selection)
    if var.get() == 1:
        simple.run()
        root.quit()
        root.destroy()
    elif var.get() == 2:
        validate.run()
        root.quit()
        root.destroy()
    elif var.get() == 3:
    	validateTL.run()
    	root.quit()
    	root.destroy()
    '''
    elif var.get() == 3:
    	validateR.run()
    	root.quit()
    	root.destroy()
    '''

root = Tk()
var = IntVar()
R1 = Radiobutton(root, text="Convert Inspection Plan", variable=var, value=1, command=sel)
R1.pack( anchor = W )

R2 = Radiobutton(root, text="Validate Inspection Plan", variable=var, value=2, command=sel)
R2.pack( anchor = W )

#R3 = Radiobutton(root, text="Validate Task List", variable=var, value=3, command=sel)
#R3.pack( anchor = W )
'''
R4 = Radiobutton(root, text="Validate Recipe", variable=var, value=4, command=sel)
R4.pack( anchor = W )
'''
label = Label(root)
label.pack()
root.quit()
root.mainloop()
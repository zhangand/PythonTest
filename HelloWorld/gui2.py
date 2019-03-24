from tkinter import *
from tkinter import filedialog
from tkinter import messagebox


class simpleapp_tk(Frame):
    def __init__(self,parent):
        Frame.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()

        self.entryVariable = StringVar()

        self.entry = Entry(self,textvariable=self.entryVariable)
        self.entry.grid(column=0,row=0,sticky='EW')
        self.entry.bind("<Return>", self.OnPressEnter)
        self.entryVariable.set(u"Enter text here.")

        button_browse = Button(self, text = u"Browse", command = lambda:self.entryVariable.set(filedialog.askopenfilename()))
        button = Button(self,text=u"Go", command=self.OnButtonClick)
        button_browse.grid(column=1,row=0)
        button.grid(column=2,row=0)

        self.labelVariable = StringVar()
        label = Label(self,textvariable=self.labelVariable,anchor="w",fg="black",bg="white")
        label.grid(column=0,row=1,columnspan=3,sticky='EW')
        self.labelVariable.set(u"Hello !")

        self.grid_columnconfigure(0,weight=1)
        #self.resizable(True,False)
        self.update()
        #self.geometry(self.geometry())
        self.entry.focus_set()
        self.entry.selection_range(0, END)


    def OnButtonClick(self):
        self.labelVariable.set( self.entryVariable.get()+" (You clicked the button)" )
        self.entry.focus_set()
        self.entry.selection_range(0, END)

    def OnPressEnter(self,event):
        self.labelVariable.set( self.entryVariable.get()+" (You pressed ENTER)" )
        self.entry.focus_set()
        self.entry.selection_range(0, END)

if __name__ == "__main__":
    root = Tk()
    app = simpleapp_tk(root)
    app.mainloop()

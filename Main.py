from tkinter import filedialog
from tkinter import *
import whatsappbot


root = Tk()
root.title("WhatsApp Bot")

# Gets the requested values of the height and width.
windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()
# Gets both half the screen width/height and window width/height
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)
root.geometry("500x200+{}+{}".format(positionRight, positionDown))
root.configure(background='#009688')
root.resizable(width=False, height=False)
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)
root.rowconfigure(2, weight=1)
root.columnconfigure(2, weight=1)

root.after(5000, lambda: root.focus_force())


contents = Frame(root)
contents.grid(row=1, column=1)

p1 = PhotoImage(file='whatsapp.png')
root.iconphoto(False, p1)


def browsefunc():
    filename = filedialog.askopenfilename(filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
    ent1.delete(0, 'end')
    ent1.insert(END, filename)

    if(ent1.get() == ''):
        b2.configure(state="disabled")
    else:
        b2.configure(state="normal")


def submit():
    bot = whatsappbot.WhatsAppBot(ent1.get())
    bot.main()


ent1 = Entry(root, font=40)
ent1.grid(row=1, column=1)

label1 = Label(root, text="Path", background='#009688', foreground="#FFFFFF")
label1.grid(row=1)

b1 = Button(root, text="Browse", font=40, command=browsefunc,
                background='#FFFFFF', foreground="#009688", highlightcolor='#009688', highlightbackground='#009688')
b1.grid(row=2, columnspan=2)

b2 = Button(root, text="Submit", font=40, command=submit,
                background='#FFFFFF', foreground="#009688", state="disabled", highlightcolor='#009688', highlightbackground='#009688' )
b2.grid(row=2, columnspan=3)

root.mainloop()

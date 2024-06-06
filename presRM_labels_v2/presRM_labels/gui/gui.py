from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from presRM_labels.labels.labels import LabelGenerator, login_preservica
from time import sleep
import os
import shlex
try:
    import secret
    s_flag = True
except ImportError:
    s_flag = False


def browse_button_output():
    file = filedialog.askdirectory()
    output.set(file)
    
def browse_button_boxref():
    file = filedialog.askopenfile(mode='r',filetypes=[('Excel Files','*.xlsx'),('CSV Files','*.csv')])
    boxref.set(file)
    
def try_login_pres(user,passwd,fail_count=0):
    try:
        content,entity = login_preservica(username=user,password=passwd)
    except Exception as e:
        if e.args[0] in {401}:    
            print('Error signing in... Please try again...')
            if fail_count == 2: print('This is your last chance...')
            elif fail_count == 3: print('YOU. ARE. DONE.'); sleep(10); raise SystemExit()
            fail_count += 1
            content, entity = try_login_pres(fail_count)
        elif e.args[0] in {404,501}:
            print('Preservica is having a moment, please try again later')
        else:
            print('An unknown error occured... Panic!')
            print(e)
            sleep(5)        
    return content,entity
    
def main():
    if s_flag:
        user = secret.username
        passwd = secret.password
    else:
        user = user_entry.get()
        passwd = pass_entry.get()
    content,entity = try_login_pres(user=user,passwd=passwd)
    boxref = shlex.split(boxref_entry.get())
    output=output_entry.get()
    combine=bool(combine_val.get())
    filesonly=bool(filesonly_val.get())
    boxonly = bool(boxonly_val.get())
    toplevel = toplevel_entry.get()
    print(boxref,output,combine,filesonly,boxonly,toplevel)
    LabelGenerator(content=content,entity=entity,boxref=boxref,output=output,boxonly=boxonly,combine=combine,filesonly=filesonly,toplevel=toplevel).main()
    
window = Tk()
window.title("UARM Label Generator")
for c in range(3):
    window.columnconfigure(c,weight=1)
for r in range(4):
    window.rowconfigure(r,weight=1)

master_frame = Frame(master=window,width=50,height=100)
master_frame.pack()

cred_frame = Frame(master=master_frame,width=50,height=50)
cred_frame.pack()
subframe1 = Frame(master=cred_frame,width=25,height=25)
subframe1.pack(side=LEFT)
subframe2 = Frame(master=cred_frame,width=25,height=25)
subframe2.pack(side=RIGHT)
if s_flag:
    user_label = Label(master=master_frame, text="Username & Password assigned by Secrets Files")
    user_label.pack()
else:
    user_label = Label(master=subframe1,text="Username")
    user_entry = Entry(master=subframe1,width=25)
    pass_label = Label(master=subframe2,text="Password")
    pass_entry = Entry(master=subframe2,width=25)
    user_label.pack()
    user_entry.pack()
    pass_label.pack()
    pass_entry.pack()
    
box_frame = Frame(master=master_frame,width=50,height=50)
box_frame.pack()
bsubframe1 = Frame(master=box_frame,width=25,height=25)
bsubframe1.pack(side=LEFT)
bsubframe2 = Frame(master=box_frame,width=25,height=25)
bsubframe2.pack(side=RIGHT)

boxref = StringVar(master=subframe1)
boxref_label = Label(master=bsubframe1,text="Box Reference")
boxref_entry = Entry(master=bsubframe1,textvariable=boxref,width=25)
boxref_label.pack()
boxref_entry.pack()
boxref_button = Button(master=bsubframe2,text="Browse",command=browse_button_boxref)
boxref_button.pack()

output_frame = Frame(master=master_frame,width=50,height=50)
output_frame.pack()
osubframe1 = Frame(master=output_frame,width=25,height=25)
osubframe1.pack(side=LEFT)
osubframe2 = Frame(master=output_frame,width=25,height=25)
osubframe2.pack(side=RIGHT)

output = StringVar(master=osubframe1,value=str(os.getcwd()))
output_label = Label(master=osubframe1,text="Output Directory")
output_entry = Entry(master=osubframe1,textvariable=output,width=25)
output_label.pack()
output_entry.pack()
output_button = Button(master=osubframe2,text="Browse",command=browse_button_output)
output_button.pack()

output_frame.pack()

toplevel_frame = Frame(master=window)
toplevel_val = StringVar(master=toplevel_frame,value="de7eaf74-7ad3-4c9b-bab6-ae21008185f0")
toplevel_label = Label(master=toplevel_frame,text="Top Level for Search")
toplevel_entry = Entry(master=toplevel_frame,textvariable=toplevel_val,state=DISABLED)
toplevel_label.pack()
toplevel_entry.pack()
toplevel_frame.pack()

checks_frame = Frame(master=window)

filesonly_val = IntVar(master=checks_frame, value=0)
filesonly_check = Checkbutton(master=checks_frame,text="Files Only",variable=filesonly_val)
filesonly_check.pack()
boxonly_val = IntVar(value=0)
boxonly_check = Checkbutton(master=checks_frame,text="Box Only",variable=boxonly_val)
boxonly_check.pack()
combine_val = IntVar(value=1)
combine_check = Checkbutton(master=checks_frame,text="Combine",variable=combine_val)
combine_check.pack()

checks_frame.pack()

run_frame = Frame(master=window)
run_button = Button(master=run_frame,text="Generate Label",command=main)
run_button.pack()
run_frame.pack()

window.mainloop()

raise SystemExit()
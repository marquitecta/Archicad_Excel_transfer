from tkinter import *
from tkinter import messagebox,filedialog
import re
from xlrd import *
import time
from archicad import ACConnection, handle_dependencies
import os, sys, uuid

handle_dependencies('xlrd')
conn = ACConnection.connect()

assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities


class MyDialog:

    def abt1(self):
        #tkinter.messagebox.showinfo("Welcome", "Welcome 2 second window")
        self.filelocation = filedialog.askopenfilename(initialdir="~/documents/pythonsnippets/", title="Select file",
                                                       filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))
        self.filename.set(os.path.basename(self.filelocation))
        self.label_2 = Label(self.window, textvariable=self.filename, relief="solid", font=("arial", 14, "bold")).grid(row=3,
                                                                                                                  column=0)
    def __init__(self,parent):
        self.filelocation=''
        self.filename =StringVar()
        self.filename.set('')
        window=self.window=Toplevel(parent)
        parent.withdraw()
        window.title("Select Your Excel File")
        window.geometry("400x500")
        self.label_1 = Label(window, text="File selection", relief="solid", font=("arial", 14, "bold")).grid(row=1,column=0)
        b1 = Button(window ,text="Click to Select File", width=30, bg='brown', fg='white', command=self.abt1,padx=5,pady=5).grid(row=2,column=0,pady=30)
        b2 = Button(window, text="Open Program", width=20, bg='brown', fg='white', command=self.ok,padx=5,pady=5).grid(row=4,column=0,pady=30)
        b3 = Button(window, text="Cancel", width=12, bg='brown', fg='white', command=self.cancel,padx=5,pady=5).grid(row=5, column=0,pady=30)

    def ok(self):
        global filelocation
        #print('got',self.filelocation)
        filelocation=self.filelocation
        self.window.destroy()

    def cancel(self):
        self.window.destroy()
def clearck():
    for i in range(len(var)):
        var[i].set(0)
    for i in range(len(varout)):
        varout[i].set(0)
def changexlsx(n,m,x):
    global frame, frame3
    if n=='sheets':
        #print('changed sheets')
        frame.destroy()
        frame3.destroy()
        frame = makeframe1(a)
        frame3 = makeframe3(a)
        getsheet(sheetclick.get())
        inputsection()
        outputsection()
    elif n=='spread':
        pass
        # print('changed spreadsheet')
    else:
        pass
        #print('Something Else')
def controlsection():
    pass
def findelements(objtype,columns,cellvalues):
   global filter_ok,foundelement,noelement,clicked
   noelement =0
   propertyid_search = []
   typeobj = clicked.get()
   typeobj = optiondict[typeobj]

   elemlist= acc.GetElementsByType(typeobj)  # Limits search to type specified in row

   for i in columns:
       buf = i
       if i.find('*',1)!= -1:
           buf= i.strip('*')
           propertyid_search.append(acu.GetBuiltInPropertyId(buf))
       else:
           buf = i.rsplit('_',1)
           propertyid_search.append(acu.GetUserDefinedPropertyId(buf[0],buf[1])) # 0 = Group, 1 = name
   result = acc.GetPropertyValuesOfElements(elemlist, propertyid_search)
   elements_found = []
   elem_position =0
   for xy in result:
       elem_hold = elemlist[elem_position]
       elem_position +=1
       loops = len(xy.propertyValues)
       propertyok = 0
       for index in range(loops):
           propertyValueObject = xy.propertyValues[index].propertyValue
           if propertyValueObject.status != 'normal':
               continue
           if hasattr(propertyValueObject, "value"):
               out1 =  xy.propertyValues[index].propertyValue.value
               if out1=='Office Workstation Solo 24':
                   print('Found Workstation')
                   print(rex_ok)
               if rex_ok.get() == 1:
                   if bool(re.match(cellvalues[index],out1)):
                       propertyok += 1
               elif out1 == cellvalues[index]: #If SS search field and plan property value are equal -- match
                   propertyok +=1
               else:
                   continue
       if propertyok == loops:
           elements_found.append(elem_hold)
   foundelement = len(elements_found)
   noelement = len(elemlist) - foundelement
   return [foundelement,noelement,elements_found]
   """.................. END OF FUNCTIOM...................."""
def getfile(outckcols,outcellvalues,elemlist):
    global objects_changed
    ''' elemlist = element guids to be modified, outckcols = selected prperties to be changed, outcellvalues = nev alues of properties'''
    elementIds =[]
    for ii in elemlist:
        buf = ii.elementId.guid
        elementIds.append(act.ElementId(buf))
    propertyid_search = []
    mark_list = []
    for i in outckcols:
        buf = i
        
        if i.find('#' , 1) != -1:
            mark_list.append(True)
            i = i.strip('#')
        else:
            mark_list.append(False)
        if i.find('*', 1) != -1:
            buf = i.strip('*')
            propertyid_search.append(acu.GetBuiltInPropertyId(buf))
        else:
            buf = i.split('_', 1)
            propertyid_search.append(acu.GetUserDefinedPropertyId(buf[0], buf[1]))  # 0 = Group, 1 = name
    property_propertyvalues = []
    index = 0
    for str in outcellvalues:
        if mark_list[index] == False:
            property_propertyvalues.append(act.PropertyValue('string','normal',str))
        else:
            disenum = act.DisplayValueEnumId(str,'displayValue')
            normalenum = act.NormalSingleEnumPropertyValue(disenum,'singleEnum','normal')
            property_propertyvalues.append(normalenum)
        index +=1
    elemPropertyValues = []
    for ii in range(len(elementIds)):
        for jj in range(len(outcellvalues)):
                propertyValue = property_propertyvalues[jj]
                test = type(propertyValue)
                if propertyValue.type != 'singleEnum':
                    propertyValue.value = outcellvalues[jj]
                elemPropertyValues.append(act.ElementPropertyValue(elementIds[ii], propertyid_search[jj], propertyValue))
                objects_changed += 1
    execution_results =acc.SetPropertyValuesOfElements(elemPropertyValues)
    return execution_results
    #print(execution_results)
    tkinter.messagebox.showinfo(('status',execution_results))
    for xxx in execution_results:
        #print('success   ',xxx.success)
        #tkinter.messagebox.showinfo(('status', xxx.success))
        if not(xxx.success) :
            print('error  ',xxx.error)

    """............  End of Function  ..........."""
def getsheet(selectedsheet):
    global sheetlist,maxrows,maxcols,arr,sheetclick
    with open_workbook(filelocation,on_demand=True) as workbook:
        if selectedsheet=='Pick Sheet' :
            sheet=workbook.sheet_by_index(0)
        else:
            sheet=workbook.sheet_by_name(selectedsheet)
        sheetlist=workbook.sheet_names()
        maxrows=sheet.nrows
        maxcols=sheet.ncols
        arr[:]=[]

        for c in range(skip,maxcols):
            arr.append(sheet.cell(0,c)) #Header values - row 1`
def getskipstatus():
    filelocation = '~/documents/pythonsnippets/archicad.xlsx'
    workbook = open_workbook(filelocation, on_demand=True)
    sheet = workbook.sheet_by_index(0)
    x=sheet.cell(7,3).value
    workbook.release_resources()
    del workbook
def getspreadvalues():   # not needed -- see mget()
    ss = []
def getxlsx():
    global filelocation,frame,frame3,sheetclick
    filelocation = filedialog.askopenfilename(initialdir="~/documents/pythonsnippets/", title="Select file",
                                                   filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))
    frame.destroy()
    frame3.destroy()
    frame = makeframe1(a)
    frame3 = makeframe3(a)
    sheetclick.set('Pick Sheet')
    getsheet('Pick Sheet')
    inputsection()
    outputsection()
def makeframe3(a):
    frame3=Frame(a,padx=20,pady=20)
    frame3.config(bd=3, relief=SUNKEN)
    frame3.grid(row=0,column=3,sticky=N,padx=30,pady=30)
    Label(frame3,text= 'Select Output',font='none 18 bold underline',bg='gray').grid(row = 0)
    return frame3
def makeframe1(a):
    frame=Frame(a,padx=20,pady=20)
    frame.config(bd=3, relief=SUNKEN)
    frame.grid(row=0,column=1,padx=30,pady=30)
    Label(frame,text= 'Search Properties',font='none 18 bold underline',bg='gray').grid(row = 0)
    return frame
def mget(maxrows):
    global foundelement, noelement,clicked,a,objects_changed
    foundelement=noelement=0
    #result = []
    workbook = open_workbook(filelocation, on_demand=True)
    sheet =sheet=workbook.sheet_by_name(sheetclick.get())
    start_time = time.time()
    reslist = [0]
    elements_matched = 0

    for ix in range(1,maxrows):
        print(f'Doing row {ix+1} of {maxrows}')
        count = ix  # starting  row to get values
        # rows and columns start at 0
        if sheet.cell_value(count,0)==0 :
            #Run box not checked
            print(f'row {ix+1} skipped')
            continue

        if sheet.cell_value(ix,1)!=clicked.get(): # check type Requested
            continue
        ckcols=[]
        cellvalues=[]
        outckcols = []
        outcellvalues = []
        for i in range(len(var)): #get selected search columns
            if (var[i].get()==1):
                ckcols.append(arr[i].value)
                if sheet.cell_type(count,i)== 3:   # DateTime
                    str = xldate_as_tuple(sheet.cell_value(count,i+skip), 0)
                    xx =f'{str[1]}/{str[2]}/{str[0]}'
                    cellvalues.append(xx)
                else:
                    cellvalues.append(sheet.cell(count,i+skip).value)
        for i in range(len(varout)): #get selected output columns
            if (varout[i].get()==1):
                #print('cell type  ',sheet.cell_type(count,i+skip))
                outckcols.append(arr[i].value)
                if sheet.cell_type(count, i + skip) == 2:
                    str = f'{sheet.cell(count,i+skip).value:9.2f}'
                    #print('Float  ',str,str.__class__)
                    outcellvalues.append(str)
                if sheet.cell_type(count,i+skip)== 3:   # DateTime
                    str = xldate_as_tuple(sheet.cell_value(count,i+skip), 0)
                    xx =f'{str[1]}/{str[2]}/{str[0]}'
                    outcellvalues.append(xx)
                if sheet.cell_type(count, i + skip) == 1:
                    str = f'{sheet.cell(count, i + skip).value}'
                    outcellvalues.append(str)
                if (sheet.cell_type(count, i + skip) == 0) or (sheet.cell_type(count, i + skip) == 6) :
                    outcellvalues.append('BLANK')
        """Passes the element type and the selected properties and their values to search on"""
        reslist=findelements(objtype, ckcols, cellvalues) #returns found elements from search criteria

        if reslist[0] > 0: #no of found elements
            elements_matched += reslist[0]
            #objects_changed += len(outckcols)
            execution_results = getfile(outckcols, outcellvalues, reslist[2]) # Sets propty value in selected properties
        else:
            messagebox.showinfo('Status','No Matches Found\n{ckcols}\n{cellvalues}')
            continue
        count =0
        for xyy in execution_results:
            count += 1
            if not xyy.success:
                #msg = xyy.error + 'count  ' + str(count)
                #tkinter.messagebox.showinfo('Status',msg)
                print(xyy.error)
            else:
                #tkinter.messagebox.showinfo('Status', xyy.success)
                #print(xyy.success)
                pass
        #tkinter.messagebox.showinfo('Status', 'Next Element')
    end_time =time.time()
    timetaken=round(end_time-start_time,1)
    tex='Update Complete - Quit "Element Transfer" to Return to Project'+ '\n' + f'time taken {timetaken} seconds'+'\n' \
    + f'Elements Affected {elements_matched}\nProperties Changed {objects_changed}'
    responce = messagebox.askyesno(' Complete Status',tex)
    if responce == True:
        a.destroy()
    clearck()
    workbook.release_resources()
   ##########    END of Founction and Program ############
def inputsection():
    var[:]=[]
    for i in range(maxcols - skip):
        var.append(IntVar(value=0))
        Checkbutton(frame, text=arr[i].value, variable=var[i]).grid(row=i + 1, sticky=W)
    Button(frame, text='Set Values', command=lambda: mget(maxrows), width=20,bg='light green').grid(row=maxcols + 1, sticky=W)
def outputsection():
    varout[:]=[]
    index =0
    for i in range(maxcols - skip):
        mark = 0
        #print(arr[i].value)
        if arr[i].value.find('*',1) > 0:
         mark=1
        varout.append(IntVar(value=0))
        if mark == 0:
            Checkbutton(frame3, text=arr[i].value, variable=varout[i]).grid(row=i + 1, sticky=W)
        else:
            if arr[i].value[-2] != '*':
                Checkbutton(frame3, text=arr[i].value, variable=varout[i],state=DISABLED).grid(row=i + 1, sticky=W)
            else:
                Checkbutton(frame3, text=arr[i].value, variable=varout[i]).grid(row=i + 1, sticky=W)
def quitprogram():
    a.destroy()

"""...............Global Variables ......................................."""
# comment test
skip=2
var=[]
varout=[]
maxrows=0
maxcols=0
arr=[]
workbook=''
sheet=''
objtype = 'None'    # API_ObjectItype name
cellvalues = []     # SS cell Values
filelocation=''
sheetlist=[]
foundelement=0
objects_changed = 0  # accumulate changed properties


"""............... Global Variables ....................................."""

a =  Tk()
filter_ok=StringVar(name='filter')
rex_ok=IntVar(name='re_exp')

remote=MyDialog(a)
a.wait_window(remote.window)
a.update()
a.deiconify()

"........... BYPASS ............"
#filelocation=filedialog.askopenfilename(initialdir="~/documents/pythonsnippets/",title="Select file",filetypes=(("Excel Files","*.xlsx"),("all files","*.*")))
"................... BYPASS................."

xlsxfilename = filelocation
#filelocation='~/documents/pythonsnippets/archicad.xlsx'
#filelocation='~/documents/pythonsnippets/'+ 'archicad-v 1_1.xlsx'
#filelocation='~/documents/pythonsnippets/'+ filelocation
#print(filelocation)

getsheet('Pick Sheet')

a.geometry('1450x800')
a.title('Element Properties Transfer')
a.configure(background='light gray')
a.iconbitmap('~/documents/pythonsnippets/python.ico')

#menu Section ...............
menu =Menu(a)
a.config(menu=menu)
submenu=Menu(menu)
editmenu=Menu(menu)
helpmenu=Menu(menu)
menu.add_cascade(label='File',menu=submenu)
menu.add_cascade(label='Edit',menu=editmenu)
menu.add_cascade(label='Help',menu=helpmenu)
submenu.add_command(label='Excel File',command=getxlsx)
submenu.add_command(label='New',command=getfile)
submenu.add_separator()
submenu.add_command(label='Exit',command=getfile)
editmenu.add_command(label='Selections',command=getfile)
helpmenu.add_command(label='Instructions',command=getfile)
#.......... Frame Section.......

frame=makeframe1(a)

frame2=Frame(a,padx=20,pady=20)
frame2.config(bd=3, relief=SUNKEN)
frame2.grid(row=0,column=2,sticky=N,padx=30,pady=30)
Label(frame2,text='Select Element Type',font='none 18 bold underline',bg='gray').grid(row=0,columnspan=2)


frame3=makeframe3(a)

'......Selection section.......'

inputsection()

Button(frame2,text='QUIT',command=a.destroy,width= 25,padx=5,pady=5,bg='light blue').grid(row=6,sticky=W,padx=10,pady=50)
'.......... Select Type Section..........'
optionlist = ['WallID','DoorID','WindowID','OpeningID','ColumnID','BeamID','SlabID','StairsID','RailingID',
              'RoofID','ShellID','SkylightID','Curtain_WallID','MorphID','ObjectID','ZoneID','MeshID']
optiondict = {'WallID':'Wall','DoorID':'Door','WindowID':'Window','OpeningID':'Opening','ColumnID':'Column','BeamID':'Beam',
             'SlabID':'Slab','StairsID':'Stars','RailingID':'Railing',
             'RoofID':'Roof','ShellID':'Shell','SkylightID':'Skylight','Curtain_WallID':'Curtain Wall','MorphID':'Morph',
             'ObjectID':'Object','ZoneID':'Zone','MeshID':'Mesh'}
clicked=StringVar()
clicked.set('ObjectID')
sheetclick=StringVar(name='sheets')
sheetclick.set('Pick Sheet')
clicked.set('Pick Type')
dList=OptionMenu(frame2,clicked,*optionlist)
dList.configure(font=('Arial',25),bg='light yellow')
dList.grid(row=1,column=1)
Label(frame2,text='Select Type ',padx=20,pady=30,font='none 18 bold').grid(row=1,column=0)
Label(frame2,text='Select Sheet ',padx=20,pady=30,font='none 18 bold').grid(row=2,column=0)
sheetselect=OptionMenu(frame2,sheetclick,*sheetlist)
sheettrace=sheetclick.trace('w',changexlsx)
sheetselect.configure(font=('Arial',25),bg='light yellow')
sheetselect.grid(row=2,column=1)
btn= Button(frame2,text='Select Xlsx File',command=getxlsx)
btn.configure(font=('Arial',25),bg='light yellow')
btn.grid(row=3,column=1)

Label(frame2,text='Excel File Used',padx=20,pady=30,font='none 18 bold').grid(row=3)
Label(frame2,text=os.path.basename(filelocation),padx=20,pady=2,font='none 18 bold').grid(row=4)
tx=Text(frame2,padx=2,pady=20,font='none 10 bold',width=30,height=5)
tx.grid(row=5)
instructions='Add instructions\nHere'
tx.insert(INSERT,instructions)
tx.config(state='disabled')
"""
filter_options=['UserDefined','FundamentalBuiltIn','UserLevelBuiltIn','All']
filter_ok.set('Set Filter Mode')
filter_options=OptionMenu(frame2,filter_ok,*filter_options)
filter_options.configure(font=('Arial',25),bg='light yellow')
filter_options.grid(row=5,column=1)
"""
Checkbutton(frame2, text='Use Regular Expressions', variable=rex_ok,font='none 18 bold').grid(row=6,column=1, sticky=W)

'..........Output Section.....'

outputsection()

a.mainloop()

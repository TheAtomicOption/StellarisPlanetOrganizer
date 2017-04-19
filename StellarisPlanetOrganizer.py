from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from zipfile import *
import ctypes.wintypes


class ListboxDragDrop(Listbox):
    """ A tk listbox with drag'n'drop reordering of entries. """
    def __init__(self, master, **kw):
        kw['selectmode'] = SINGLE
        kw['activestyle'] = 'none'
        Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.getState, add='+')
        self.bind('<Button-1>', self.setCurrent, add='+')
        self.bind('<B1-Motion>', self.shiftSelection)
        self.curIndex = None
        self.curState = None

    def setCurrent(self, event):
        """ gets the current index of the clicked item in the listbox """
        self.curIndex = self.nearest(event.y)

    def getState(self, event):
        """ checks if the clicked item in listbox is selected """
        i = self.nearest(event.y)
        self.curState = 1 #self.selection_includes(i)

    def shiftSelection(self, event):
        """ shifts item up or down in listbox """
        i = self.nearest(event.y)
        if self.curState == 1:
            self.selection_set(self.curIndex)
        else:
            self.selection_clear(self.curIndex)
        if i < self.curIndex:
            # Moves up
            x = self.get(i)
            selected = self.selection_includes(i)
            self.delete(i)
            self.insert(i+1, x)
            if selected:
                self.selection_set(i+1)
            self.curIndex = i
        elif i > self.curIndex:
            # Moves down
            x = self.get(i)
            selected = self.selection_includes(i)
            self.delete(i)
            self.insert(i-1, x)
            if selected:
                self.selection_set(i-1)
            self.curIndex = i


class StellarisSave:
    """A class for reading/writing Stellaris save files to change the order of player planets in the ingame outliner"""
    def __init__(self):

        self.filenameFullPath = StringVar()
        self.planetnumlistindex = IntVar()
        self.planetnumlist = []
        self.planetnumnamelist =[]
        self.gslines = []
        self.planetSaveIndexStart = IntVar()
        self.planetSaveIndexEnd = IntVar()

        #Start in Stellaris save game directory
        CSIDL_PERSONAL = 5  # My Documents
        SHGFP_TYPE_CURRENT = 0  # Get current, not default value
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
        self.lastpath = buf.value + '\Documents\Paradox Interactive\Stellaris\save games\\'
        self.lastpath = self.lastpath.replace("\\", "/")  # in python, \ in strings causes problems
        # print(self.lastpath) #Debug print

    def openSave(self):
        """Opens a Stellaris save and reads planet data"""
        print('Opening save... ')
        self.filenameFullPath = filedialog.askopenfilename(filetypes=[("Stellaris Saves", "*.sav")],
                                                           initialdir=self.lastpath,
                                                           title='Open Stellaris Save File')
        if self.filenameFullPath is None:
            return FALSE
        self.lastpath = self.filenameFullPath[:self.filenameFullPath.rindex("/")]

        print('Opening: ' + self.filenameFullPath)  # Debug file to open

        savezip = ZipFile(self.filenameFullPath, 'r')
        gamestate = ZipFile.open(savezip, 'gamestate', 'r')

        self.gslines = gamestate.readlines()
        # print(type(self.gslines[3])) # yup they're bytes. need to bytes.decode('UTF-8').
        gamestate.close()
        savezip.close()
        # find first 'owned_planets=' which should be the players as the player is the first empire.
        # for loop finds line indexes for planet list and owned planet line.
        # Ideally a parser than can just read the Stellaris save file properly should replace this, but
        # only reason to build that would be to make a full on trainer
#           # which would be mostly pointless since we have console commands.
        for i, val in enumerate(self.gslines, start=0):
            if (val.decode('UTF-8').find('planet={') > -1):
                self.planetSaveIndexStart = i
            if (val.decode('UTF-8').find('country={') > -1):
                self.planetSaveIndexEnd = i
            if (val.decode('UTF-8').find('owned_planets=') > -1):
                self.planetnumlistindex = i + 1
                break
        print("planet start index: " + str(self.planetSaveIndexStart))
        print("planet end index: " + str(self.planetSaveIndexEnd))
        print("owned planets index: " + str(self.planetnumlistindex))

        print(self.gslines[self.planetnumlistindex])
        a = self.gslines[self.planetnumlistindex].decode('UTF-8')[3:]  # remove 3 indenting tabs
        self.planetnumlist = a.split(" ")
        print(*self.planetnumlist, sep='\n') # Debug print list of planets

        # populate planet names
        planetlines = self.gslines[self.planetSaveIndexStart:self.planetSaveIndexEnd]
        self.planetnumnamelist.clear()
        for planetnum in self.planetnumlist:
            for i, val in enumerate(planetlines, start=0):
                if val.decode('UTF-8') == '\t' + planetnum + '={\n':
                    # print(planetlines[i+1])
                    self.planetnumnamelist.append([planetnum,planetlines[i+1].decode('UTF-8')[8:-2]])

        for l in self.planetnumnamelist:
            print(*l, sep=' ')

    def setPlanets(self, a: str):
        """Accept a triple tab prefixed, unicode, space separated, \n terminated, string to replace the line with planet numbers"""
        self.gslines = self.gslines[:self.planetnumlistindex] + self.gslines[self.planetnumlistindex + 1:]
        self.gslines.insert(self.planetnumlistindex, a.encode('UTF-8'))

    def saveFile(self):
        """Saves planet data to disk"""
        # savezip = self.filenameFullPath[:self.filenameFullPath.rindex("/")]+"/TestSave.sav.zip" # Save to test file for debug
        savezip = ZipFile(self.filenameFullPath, 'w')
        gamestate = ZipFile.open(savezip, 'gamestate', 'w')

        for line in self.gslines:
            gamestate.write(line)

        gamestate.close()
        savezip.close()
        print('File closed.')



class MainWindow:

    def __init__(self, master):
        self.planetList = ['no', 'save', 'loaded', 'yet']
        self.planetNamedList = []
        self.filenamePartialPath = StringVar()
        self.filenamePartialPath.set('No File Selected')
        self.activeSave = StellarisSave()

        self.master = master
        master.title('Stellaris Planet Organizer')

        mainframe = ttk.Frame(root, padding='3 3 3 3')
        mainframe.grid(column=0, row=0, sticky=(W, N, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)

        # may end up using Treeview object instead? Going to try with a listbox first.
        # treeplanets = ttk.Treeview(mainframe)
        # treeplanets.grid(column=1,row=1,sticky=(N,W,E))
        # for planet in planetList:
        #     treeplanets.insert('','end',text=planet)

        self.listplanets = ListboxDragDrop(mainframe, listvariable=self.planetNamedList, width=80, height=10)
        self.listplanets.grid(column=1, row=1, sticky=(N, W, E, S))
        self.listplanets.rowconfigure(0, weight=1)
        self.listplanets.columnconfigure(0, weight=1)

        mainframe_c2r1 = ttk.Frame(mainframe)
        mainframe_c2r1.grid(column=2, row=1, sticky=(W, N, E, S))
        bOpen = ttk.Button(mainframe_c2r1, text='Open Save...', command=self.chooseFile)
        bOpen.grid(column=1, row=1, sticky=(N, W, E))
        bOpen.focus()

        bSave = ttk.Button(mainframe_c2r1, text='Save Planet Order', command=self.saveFile)
        bSave.grid(column=1, row=2, sticky=(N, W, E))
        bSave.grid_configure(pady=5)
        ttk.Label(mainframe, textvariable=self.filenamePartialPath).grid(column=1, row=2, sticky=(S, W))

        for child in mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def chooseFile(self):
        """On click event for Open Save... button"""
        print('Picking a file...')
        self.activeSave.openSave()

        if self.activeSave.filenameFullPath:
            a = self.activeSave.filenameFullPath
            b = a[:a.rindex("/") - 1]
            self.filenamePartialPath.set(a[b.rindex("/"):])
            print(self.filenamePartialPath.get())

            self.listplanets.delete(0, END)
            self.listplanets.insert(END, *self.activeSave.planetnumnamelist)
            #for item in self.activeSave.planetnumnamelist:
            #    self.listplanets.insert(END, item)

    def saveFile(self):
        print('Saving data...')
        a = str()
        for item in self.listplanets.get(0, END):
            a = a + item[0] + ' '
        a = '\t\t\t' + a + '\n'

        print(a)
        print(a.encode('UTF-8'))
        print(*self.planetNamedList, sep='\n')
        self.activeSave.setPlanets(a)

        print(self.activeSave.gslines[self.activeSave.planetnumlistindex - 1])
        print(self.activeSave.gslines[self.activeSave.planetnumlistindex])
        print(self.activeSave.gslines[self.activeSave.planetnumlistindex + 1])

        self.activeSave.saveFile()

root = Tk()
pomainwindow = MainWindow(root)
root.update()
root.minsize(root.winfo_width(), root.winfo_height())
root.mainloop()

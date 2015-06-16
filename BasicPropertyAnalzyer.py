from Tkinter import *
import tkMessageBox as box
import tkFileDialog 
import stfio
import tkSimpleDialog 
rec = stfio.read("14122000.abf")
last_value = 0
complete_dataset = np.array([0,0])
rig = 0
abffile = "None"
class spontActiviy():
    def attempted_action_potential_found(complete_section):
        #Sucht nach versuchten Aktionspotentialen (Peak hoeher als -20 mV) im gesamten Trace
        if complete_dataset[complete_dataset.argmax()] < -20:
            return False
        else:
            return True

    def action_potential_found(complete_section):
        #Sucht nach Aktionspotentialen im gesamten Trace
        if complete_dataset[complete_dataset.argmax()] < 0:
            return False
        else:
            return True
        
    def mean_first_10s(self, complete_section):     
        global last_value, mean_array
        #Calculates the mean of the first 10 s
        first_value = 0
        # 10000 ms period divided by sampling rate
        # for example: 10.000/0.2 (0.2 ms interval) = 50.000 data points
        # for example 2: 10.000/0.05 (0.05 ms interval) = 50.000 data points
        last_value = 10000/rec.dt

        mean_array = complete_section[first_value:last_value]

        return mean_array.mean()

class Example(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent        
        self.initUI()
        
    def initUI(self):
      
        self.parent.title("Basic property analyzer for the PJK lab")
        self.pack(fill=BOTH, expand=1)
        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Open file...", command=self.onOpen)
        fileMenu.add_command(label="Exit", command=self.onExit)    
        menubar.add_cascade(label="File", menu=fileMenu)     
        
        helpMenu = Menu(menubar)
        helpMenu.add_command(label="Help", command=self.onHelp)
        helpMenu.add_command(label="About", command=self.onAbout)    
        menubar.add_cascade(label="?", menu=helpMenu)              

    
        #self.txt = Text(self)
        #self.txt.pack(fill=BOTH, expand=1)

    def onExit(self):
        root.destroy()
        
    def onAbout(self):
        box.showinfo("About Basic Properties Analyzer", "Version 0.0.1, June 2015\n\n Copyright: Christian Schnell (cschnell@schnell-thiessen.de)") 
    
    def onHelp(self):
        pass     
    

		
    def onOpen(self):
        global rec, complete_dataset, abffile
        ftypes = [('Axon binary files', '*.abf'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes)
        fl = dlg.show()
        
        if fl != '':
            rec = stfio.read(str(fl))
            path = str(fl)
            
            #Extract filename for Excel file
            singles = path.split("/")
            abffile = singles[-1]

            # Read sampling interval in ms -> 0.2 ms for testing file
            # 1 section mit 256 Datenpunkten entspricht 51 ms bei 0.2 sampling interval
            # print(rec.dt)
            
            #Daten aus Section1 korrektes Array res kopieren
            complete_dataset = rec[0][0].asarray()
            
            #Wie viele Sections gibt es in Channel 1?
            number_of_sections = len(rec[0])
  
            # Alle Section aus Channel 0 in eine Datenspur vereinen
            # Needs improvement, very slow, possibly too slow for comfortable working with it
            i = 1
            while i < number_of_sections:
                helper_arr = rec[0][i].asarray()    
                complete_dataset = np.hstack((complete_dataset,helper_arr))
                #print complete_dataset.shape
                i += 1
            #text = self.readFile(fl)
            #self.txt.insert(END, text)
            
            # noch nicht ganz sauber, muss in separate Klasse
            if spontActiviy.action_potential_found:
                print "no action potential"
            else:
                print "Action potentials found"
                
            if spontActiviy.attempted_action_potential_found:
                print "no attempted action potential"
            else:
                print "Attempted Action potentials found"

            print spontActiviy().mean_first_10s(complete_dataset)   
            
            print rig
            
            print abffile

    def readFile(self, filename):
        pass
        #f = open(filename, "r")
        #text = f.read()
        #return text
         

root = Tk()
ex = Example(root)
rig = tkSimpleDialog.askstring("Rig number", "Enter the number of your rig.")


root.geometry("300x250+300+300")
root.mainloop()

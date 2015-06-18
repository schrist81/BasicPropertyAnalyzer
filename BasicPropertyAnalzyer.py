from Tkinter import *
import tkMessageBox as box
import tkFileDialog 
import stfio
import tkSimpleDialog 

# Imports for Export to Excel
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter
from openpyxl.styles import Font, Fill

# Imports for Peak Detection
import numpy as np
from math import pi, log
import pylab
from scipy import fft, ifft
from scipy.optimize import curve_fit

rec = stfio.read("14122000.abf")
last_value = 0
complete_dataset = np.array([0,0])
rig = 0
abffile = "None"
iAP_file = "None"
input_resistance = 0

#Create Excel file
wb = Workbook()
dest_filename = 'F:\Programmierung\Python\empty_book.xlsx'
ws1 = wb.active
ws1.title = "Test"

# Filling of Column A with Descriptions
ws1['A5'] = "Filename of the first CC recording"
ws1['A6'] = "rig"
ws1['A7'] = "Vmembrane (mV)"
ws1['A8'] = "1-5 slope input res (Mohms)"
ws1['A9'] = "Capacitance (pf)"  
ws1['A10'] = "Series resistance (Mohm)"
ws1['A11'] = "Voltage Offset (mV)"
ws1['A12'] = "sAP"
ws1['A13'] = "n (1)"
ws1['A14'] = "a (2)"
ws1['A15'] = "y (3)"
ws1['A16'] = "sAP code"
ws1['A17'] = "sAP description (burst | single)"
ws1['A18'] = "iAP"
ws1['A19'] = "n (1)"
ws1['A20'] = "a (2)"
ws1['A21'] = "ys (3)"
ws1['A22'] = "at (4)"
ws1['A23'] = "yt (5)"
ws1['A24'] = "iAP code"


ws1['A27'] = "filename"
ws1['A28'] = "sweep number"

ws1['A30'] = "Overshoot (mv)"




ws1['A36'] = "Spike frequency (1 s)"
ws1['A37'] = "-10"
ws1['A38'] = "0"
ws1['A39'] = "10"
ws1['A40'] = "20"
ws1['A41'] = "30"
ws1['A42'] = "40"
ws1['A43'] = "50"
ws1['A44'] = "60"
ws1['A45'] = "70"
ws1['A46'] = "80"
ws1['A47'] = "90"
ws1['A48'] = "100"
ws1['A49'] = "110"
ws1['A50'] = "120"
ws1['A51'] = "130"
ws1['A52'] = "140"
ws1['A53'] = "150"
ws1['A54'] = "160"
ws1['A55'] = "170"
ws1['A56'] = "180"
dest_filename = 'F:\Programmierung\Python\empty_book.xlsx'
wb.save(filename = dest_filename)


def _datacheck_peakdetect(x_axis, y_axis):
    if x_axis is None:
        x_axis = range(len(y_axis))
    
    if len(y_axis) != len(x_axis):
        raise (ValueError, 
                'Input vectors y_axis and x_axis must have same length')
    
    #needs to be a numpy array
    y_axis = np.array(y_axis)
    x_axis = np.array(x_axis)
    return x_axis, y_axis


def peakdetect(y_axis, x_axis = None, lookahead = 100, delta=0):
    """
    Converted from/based on a MATLAB script at: 
    http://billauer.co.il/peakdet.html
    found at https://gist.github.com/sixtenbe/1178136
    
    function for detecting local maximas and minmias in a signal.
    Discovers peaks by searching for values which are surrounded by lower
    or larger values for maximas and minimas respectively
    
    keyword arguments:
    y_axis -- A list containg the signal over which to find peaks
    x_axis -- (optional) A x-axis whose values correspond to the y_axis list
        and is used in the return to specify the postion of the peaks. If
        omitted an index of the y_axis is used. (default: None)
    lookahead -- (optional) distance to look ahead from a peak candidate to
        determine if it is the actual peak (default: 200) 
        '(sample / period) / f' where '4 >= f >= 1.25' might be a good value
    delta -- (optional) this specifies a minimum difference between a peak and
        the following points, before a peak may be considered a peak. Useful
        to hinder the function from picking up false peaks towards to end of
        the signal. To work well delta should be set to delta >= RMSnoise * 5.
        (default: 0)
            delta function causes a 20% decrease in speed, when omitted
            Correctly used it can double the speed of the function
    
    return -- two lists [max_peaks, min_peaks] containing the positive and
        negative peaks respectively. Each cell of the lists contains a tupple
        of: (position, peak_value) 
        to get the average peak value do: np.mean(max_peaks, 0)[1] on the
        results to unpack one of the lists into x, y coordinates do: 
        x, y = zip(*tab)
    """
    max_peaks = []
    min_peaks = []
    dump = []   #Used to pop the first hit which almost always is false
       
    # check input data
    x_axis, y_axis = _datacheck_peakdetect(x_axis, y_axis)
    # store data length for later use
    length = len(y_axis)
    
    
    #perform some checks
    if lookahead < 1:
        raise ValueError, "Lookahead must be '1' or above in value"
    if not (np.isscalar(delta) and delta >= 0):
        raise ValueError, "delta must be a positive number"
    
    #maxima and minima candidates are temporarily stored in
    #mx and mn respectively
    mn, mx = np.Inf, -np.Inf
    
    #Only detect peak if there is 'lookahead' amount of points after it
    for index, (x, y) in enumerate(zip(x_axis[:-lookahead], 
                                        y_axis[:-lookahead])):
        if y > mx:
            mx = y
            mxpos = x
        if y < mn:
            mn = y
            mnpos = x
        
        ####look for max####
        if y < mx-delta and mx != np.Inf:
            #Maxima peak candidate found
            #look ahead in signal to ensure that this is a peak and not jitter
            if y_axis[index:index+lookahead].max() < mx:
                if y_axis[mxpos] > 0: 
                    max_peaks.append([mxpos, mx])
                    dump.append(True)
                #set algorithm to only find minima now
                mx = np.Inf
                mn = np.Inf
                if index+lookahead >= length:
                    #end is within lookahead no more peaks can be found
                    break
                continue
            #else:  #slows shit down this does
            #    mx = ahead
            #    mxpos = x_axis[np.where(y_axis[index:index+lookahead]==mx)]
        
        ####look for min####
        
        if y > mn+delta and mn != -np.Inf:
            #Minima peak candidate found 
            #look ahead in signal to ensure that this is a peak and not jitter
            if y_axis[index:index+lookahead].min() > mn:
                min_peaks.append([mnpos, mn])
                dump.append(False)
                #set algorithm to only find maxima now
                mn = -np.Inf
                mx = -np.Inf
                if index+lookahead >= length:
                    #end is within lookahead no more peaks can be found
                    break
            #else:  #slows shit down this does
            #    mn = ahead
            #    mnpos = x_axis[np.where(y_axis[index:index+lookahead]==mn)]
    return max_peaks    



class spontActivity():
    def attempted_action_potential_found(complete_section):
        #Sucht nach versuchten Aktionspotentialen (Peak hoeher als -20 mV) im gesamten Trace
        return complete_dataset[complete_dataset.argmax()] > -10


    def action_potential_found(complete_section):
        #Sucht nach Aktionspotentialen im gesamten Trace
        return complete_dataset[complete_dataset.argmax()] > 0

        
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
        
class inducedActivity():
    def calculateInputResistance(self, first_mean, second_mean):
        global input_resistance
        '''
        Receives means of two current steps in mV
        Assumes Delta of 10 pA
        Returns the input resistance in MOhm, calculated with Ohm's law U = R * I
        '''
        #transform pA in A
        deltaI = 1E-11
        
        #DeltaU and transform mV in V
        deltaU = (second_mean-first_mean)/1000
        
        # U = R * I => R = U/I
        resistance = deltaU/deltaI
        
        #tranform Ohm in MOhm
        input_resistance = resistance/1000000
        
        return input_resistance

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
        fileMenu.add_command(label="Open file gap free recording...", command=self.onOpenGapFree)
        fileMenu.add_command(label="Open current step recording...", command=self.onOpenCurrentStep)
        fileMenu.add_command(label="Exit", command=self.onExit)    
        menubar.add_cascade(label="File", menu=fileMenu)     
        
        helpMenu = Menu(menubar)
        helpMenu.add_command(label="Help", command=self.onHelp)
        helpMenu.add_command(label="About", command=self.onAbout)    
        menubar.add_cascade(label="?", menu=helpMenu)              

    
        #self.txt = Text(self)
        #self.txt.pack(fill=BOTH, expand=1)

    def onExit(self):
        wb.save(filename = dest_filename)
        root.destroy()
        
    def onAbout(self):
        box.showinfo("About Basic Properties Analyzer", "Version 0.0.1, June 2015\n\n Copyright: Christian Schnell (cschnell@schnell-thiessen.de)") 
    
    def onHelp(self):
        pass     
    
    def onOpenCurrentStep(self):
        global rec, complete_dataset, iAP_file, input_resistance, ws1
        ftypes = [('Axon binary files', '*.abf'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes)
        fl = dlg.show()
        
        if fl != '':
            rec = stfio.read(str(fl))
            path = str(fl)
            
            #Extract filename for Excel file
            singles = path.split("/")
            iAP_file = singles[-1]      
        

            i = 0
            #set sweep number to -1 to set it later to the first sweep with an AP
            sweep = -1
            print sweep
            while i < len(rec[0]):
                
                trace = rec[0][i].asarray()
                
                #sampling rate: rec.dt in ms, mean of interval between 0.5 and 1 s
                interval_begin = 500/rec.dt
                interval_end = 1000/rec.dt
            
                
                if (i == 0):
                    if (trace[trace.argmax()] < 0):
                        first_mean = trace[interval_begin:interval_end].mean()
                    else:
                        first_mean = "Error"
                elif (i == 1): 
                    if (trace[trace.argmax()] < 0):
                        second_mean = trace[interval_begin:interval_end].mean()
                    else:
                        second_mean = "Error"
                i = i + 1
                

                #extracts the 1s current step part of each sweep
                injected_trace = trace[1612:101612]                 
                coordinate = "B" + str(36+i)
                if (injected_trace[injected_trace.argmax()] > 0):
                    if sweep == -1:
                        ws1['B28'] = i
                        sweep = i
                        # Determine Overshoot
                        print peakdetect(injected_trace, None, lookahead = 300, delta=0)[0][1]
                        #ws1['B30'] = 
                    iAP_frequency = len(peakdetect(injected_trace, None, lookahead = 300, delta=0))
                    ws1[coordinate] = iAP_frequency
                else:
                    ws1[coordinate] = 0
            ws1['B8'] = inducedActivity().calculateInputResistance(first_mean, second_mean)  
            ws1['B27'] = iAP_file
            wb.save(filename = dest_filename)
            
    
        
    def onOpenGapFree(self):
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
            # Needs improvement, very slow, definetely too slow for comfortable working with it
            i = 1
            while i < number_of_sections:
                helper_arr = rec[0][i].asarray()    
                complete_dataset = np.hstack((complete_dataset,helper_arr))
                #print complete_dataset.shape
                i += 1
            #text = self.readFile(fl)
            #self.txt.insert(END, text)
            
            # noch nicht ganz sauber, muss in separate Klasse
            if spontActivity().action_potential_found():
                sAP =  3
                ws1['B15'] = sAP
                ws1['B13'] = ''
                ws1['B14'] = ''

            else:
                if spontActivity().attempted_action_potential_found():
                    sAP =  2
                    ws1['B14'] = sAP
                    ws1['B13'] = ''
                    ws1['B15'] = ''
                else:
                    sAP =  1
                    ws1['B13'] = sAP
                    ws1['B14'] = ''
                    ws1['B15'] = ''
                    #RMP only determined in sAP code 1 files (other cases more complicated, maybe later)
                    ws1['B7'] = spontActivity().mean_first_10s(complete_dataset)   

            ws1['B16'] = sAP
            ws1['B6'] = rig
            ws1['B5'] = abffile
            wb.save(filename = dest_filename)
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

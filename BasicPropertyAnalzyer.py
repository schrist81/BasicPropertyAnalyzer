import sys
# IO package für Pfadnamen
stflist = ['C:\\Program Files\\Stimfit 0.14\\wx-3.0-msw', 
           'C:\\Program Files\\Stimfit 0.14\\wx-3.0-msw', 
           'C:\\Program Files\\Stimfit 0.14', 
           'C:\\Program Files\\Stimfit 0.14\\stf-site-packages', 
           'C:\\WINDOWS\\SYSTEM32\\python27.zip', 
           'C:\\Users\\c-sch_000\\Anaconda\\Lib', 
           'C:\\Users\\c-sch_000\\Anaconda\\DLLs', 
           'C:\\Python27\\Lib', 
           'C:\\Python27\\DLLs', 
           'C:\\Python27\\Lib\\lib-tk', 
           'C:\\Program Files\\Stimfit 0.14', 
           'C:\\Python27', 
           'C:\\Python27\\lib\\site-packages']
sys.path = list(set(sys.path + stflist))
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

# Imports for Peak Detection and Calculations
import numpy as np
from math import pi, log
import pylab
from scipy import fft, ifft
from scipy.optimize import curve_fit

# Import for directory settings

import os
from time import strftime

rec = stfio.read("14122000.abf")
last_value = 0
complete_dataset = np.array([0,0])
rig = 0
abffile = "None"
iAP_file = "None"
input_resistance = 0
capacitance = 1000
openDirectory = "C:\\"
saveDirectory = "C:\\temp\\"

#Create Excel file
wb = Workbook()
#dest_filename = saveDirectory + "\\"  + strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"
dest_filename = ""
dest_directory = ""

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
ws1['A31'] = "Afterhyperpolarization (mv)"
ws1['A32'] = "(calculated) Spike Height (mV)"

ws1['A81'] = "I Na max (pA/pF)"
ws1['A82'] = "I K max (pA/pF)"

ws1['A84'] = "Na activation (pA)"

ws1['A122'] = "Na activation (pA/pF)"

ws1['A160'] = "U = R*I, (120-66.68)"
ws1['A161'] = "Na activation driving force corrected (66.68 mV)"

ws1['A201'] = "Na activation, normalized conductance G/Gmax"

ws1['A241'] = "Na inactivation (pA)"

ws1['A280'] = "Na inactivation conductance (driving force corrected (66.68 mV))"

ws1['A320'] = "Na inactivation, normalized conductance G/Gmax"

ws1['A359'] = "K currents (pA) last 50 ms averaged"

ws1['B24'] = "=SUM(B19:B23)"


def voltageStepInserter(first_column, last_column, voltage_begin, step):
    m = 0
    for k in range(first_column,last_column+1):
        coordinate = "A" + str(k)
        ws1[coordinate] = voltage_begin-(m*-step)
        m = m + 1   
# instert steps for current step protocol
voltageStepInserter(37, 56, -10, 10)

#wb.save(filename = dest_filename)


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


def peakdetect(y_axis, direction, x_axis = None, lookahead = 100, delta=0):
    global min_peak
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
    min_peak = 10000000
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
                    #dump.append(True)
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
                #print y_axis[mnpos] 
                min_peaks.append([mnpos, mn])
                if y_axis[mnpos] < min_peak:
                    min_peak = y_axis[mnpos]
                    
                #dump.append(False)
                #set algorithm to only find maxima now
                mn = -np.Inf
                mx = -np.Inf
                if index+lookahead >= length:
                    #end is within lookahead no more peaks can be found
                    break
            #else:  #slows shit down this does
            #    mn = ahead
            #    mnpos = x_axis[np.where(y_axis[index:index+lookahead]==mn)]
    if direction == "positive":
        return max_peaks 
        
    if direction == "negative":
        return min_peak
       



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
        
        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_command(label="Open file gap free recording...", command=self.onOpenGapFree)
        fileMenu.add_command(label="Open current step recording...", command=self.onOpenCurrentStep)
        fileMenu.add_command(label="Open voltage step recording...", command=self.onOpenVoltageStep)
        fileMenu.add_command(label="Exit", command=self.onExit)    
        menubar.add_cascade(label="File", menu=fileMenu)     
        
        optionMenu = Menu(menubar, tearoff=0)
        optionMenu.add_command(label="Set parent file directory for recordings", command=self.askdirectory)
        optionMenu.add_command(label="Set directory for generated excel files", command=self.askdirectorySave)
        menubar.add_cascade(label="Options", menu=optionMenu)           
        
        helpMenu = Menu(menubar, tearoff=0)
        helpMenu.add_command(label="Help", command=self.onHelp)
        helpMenu.add_command(label="About", command=self.onAbout)    
        menubar.add_cascade(label="?", menu=helpMenu)              

    
        #self.txt = Text(self)
        #self.txt.pack(fill=BOTH, expand=1)
        
    def askdirectory(self):
        global openDirectory, saveDirectory, dest_filename, dest_directory
        """Returns a selected directoryname."""
        openDirectory = tkFileDialog.askdirectory()
        saveDirectory = openDirectory
        dest_directory = openDirectory
        print dest_directory
        
    def askdirectorySave(self):
        global saveDirectory, dest_filename, dest_directory
        """Returns a selected directoryname."""
        saveDirectory = tkFileDialog.askdirectory()  
        dest_directory  = saveDirectory
        #dest_filename = saveDirectory +  "\\"  + abffile[0] + ".xlsx"
        print dest_directory

    def onExit(self):
        #wb.save(filename = dest_filename)
        root.destroy()
        
    def onAbout(self):
        box.showinfo("About Basic Properties Analyzer", "Version 1.1, July 2015\n\n Copyright: Christian Schnell (cschnell@schnell-thiessen.de)\n\n https://github.com/schrist81/BasicPropertyAnalyzer") 
    
    def onHelp(self):
        pass     
    
    def onOpenVoltageStep(self):
        global rec, complete_dataset, ws1
        ftypes = [('Axon binary files', '*.abf'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes, initialdir = openDirectory)
        fl = dlg.show()
        
        if fl != '':
            rec = stfio.read(str(fl))
            path = str(fl)
            
            #Extract filename for Excel file
            singles = path.split("/")
            iAP_file = singles[-1]      
        

            i = 0
            K_max_current = 0
            Na_act_max_current = 0
            # voltage step inserter for K mean
            voltageStepInserter(360, 396, -120, 5)
            # voltage step inserter for K current density
            voltageStepInserter(399, 435, -120, 5)        
            # voltage step inserter for Na activation currents
            voltageStepInserter(85, 121, -120, 5)         
            # voltage step inserter for Na activation current densities
            voltageStepInserter(123, 159, -120, 5)      
            # voltage step inserter for Na activation current densities driving force correction
            voltageStepInserter(162, 198, -120, 5)        
            # voltage step inserter for Na activation current densities driving force correction normalized conductance
            voltageStepInserter(202, 238, -120, 5)  
            
            # voltage step inserter for Na activation currents
            voltageStepInserter(242, 278, -120, 5)          
            # voltage step inserter for Na inactivation driving force correction
            voltageStepInserter(281, 317, -120, 5)         
            # voltage step inserter for Na inactivation driving force correction normalized conductance
            voltageStepInserter(321, 357, -120, 5)                
            
            while i < len(rec[0]):
                
                trace = rec[0][i].asarray()
                
                ''' Na activation currents 
                '''                
                # sampling rate: rec.dt in ms, mean of interval between 58 and 100 ms
                # look for negative peak later
                Na_activation_interval_begin = 58/rec.dt
                Na_activation_interval_end = 100/rec.dt
                
                # Na activation currents
                min_activation_peak = peakdetect(trace[Na_activation_interval_begin:Na_activation_interval_end], "negative", None, lookahead = 14 , delta=0)               
                #print str(i)+": "+str(min_activation_peak)
                coordinateNaAct = "B" + str(85+i)
                ws1[coordinateNaAct] = min_activation_peak         
                
                # Na activation current densities
                coordinate = "B" + str(123+i)
                field = "="+coordinateNaAct+"/B$9"
                ws1[coordinate] = field             
                
                # Na activation current densities - driving force correction
                coordinateNaActDriForce = "B" + str(162+i)
                ws1[coordinateNaActDriForce] = "="+str(coordinate)+"/($A"+str(162+i)+"-66.68)"
                
                # Na activation current densities - driving force correction - normalized conductance
                coordinateNaActDriForceConduct = "B" + str(202+i)
                ws1[coordinateNaActDriForceConduct] = "=B"+str(162+i)+"/MAX(B162:B198)"  
                
                # save maximum Na current
                if min_activation_peak < Na_act_max_current:
                    Na_act_max_current = min_activation_peak                
                Na_act_max_current_density = "="+str(Na_act_max_current)+"/B9"
                ws1['B81'] = Na_act_max_current_density
                
                
                ''' Na inactivation currents
                '''
                # sampling rate: rec.dt in ms, mean of interval between 50 and 100 ms
                # look for negative peak later
                Na_inactivation_interval_begin = 250/rec.dt
                Na_inactivation_interval_end = 270/rec.dt              
                
				# Na inactivation currents
                min_inactivation_peak = peakdetect(trace[Na_inactivation_interval_begin:Na_inactivation_interval_end], "negative", None, lookahead = 10 , delta=0)               
                #print str(i)+": "+str(min_activation_peak)
                coordinateNaInact = "B" + str(242+i)
                ws1[coordinateNaInact] = min_inactivation_peak         
                
                # Na inactivation driving force correction
                coordinateNaInactDriForce = "B" + str(281+i)
                ws1[coordinateNaInactDriForce] = "="+str(coordinateNaInact)+"/($A"+str(281+i)+"-66.68)"
                
                # Na inactivation - driving force correction - normalized conductance
                coordinateNaInactDriForceConduct = "B" + str(321+i)
                ws1[coordinateNaInactDriForceConduct] = "=B"+str(281+i)+"/MAX(B281:B317)"                  
                #=B281/MAX(B$281:B$317)
                ''' K currents 
                '''
                # sampling rate: rec.dt in ms, mean of interval between 208 and 258 ms
                K_mean_interval_begin = 208/rec.dt
                K_mean_interval_end = 258/rec.dt                    
                
                # determine mean K currents and fill in table
                coordinateA = "B" + str(360+i)
                ws1[coordinateA] = trace[K_mean_interval_begin:K_mean_interval_end].mean()
                # determine mean K current density and fill in table
                coordinate = "B" + str(399+i)
                field = "="+coordinateA+"/B9"
                ws1[coordinate] = field
                # save maximum potassium current
                mean = trace[K_mean_interval_begin:K_mean_interval_end].mean()
                if mean > K_max_current:
                    K_max_current = mean

                i = i + 1
                
        
            # save maximum K+ current
            K_max_current_density = "="+str(K_max_current)+"/B9"
            ws1['B82'] = K_max_current_density
            print "Inactivation protocol"
        wb.save(filename = dest_filename)
    
    def onOpenCurrentStep(self):
        global rec, complete_dataset, iAP_file, input_resistance, ws1
        ftypes = [('Axon binary files', '*.abf'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes, initialdir = openDirectory)
        fl = dlg.show()
        
        if fl != '':
            rec = stfio.read(str(fl))
            path = str(fl)
            
            #Extract filename for Excel file
            singles = path.split("/")
            iAP_file = singles[-1]      
        
            #Delete overshoot, AHP, sweep number and spike height in case
            #a file with no spikes is opened after one with spikes
            ws1['B27'] = ""            
            ws1['B28'] = ""
            ws1['B30'] = ""
            ws1['B31'] = ""
            ws1['B32'] = ""
            #set sweep number to -1 to set it later to the first sweep with an AP
            sweep = -1

            for i in range(0,len(rec[0])):
                #print str(i)+": "
                #trace = rec[0][i].asarray()
                trace = np.array(rec[0][i])  
                
                
                #sampling rate: rec.dt in ms, mean of interval between 0.5 and 1 s
                interval_begin = 500/rec.dt
                interval_end = 1000/rec.dt
            
                #extracts the 1s current step part of each sweep
                injected_trace = trace[1612:51612]    

                if (i == 0):
                    if (injected_trace[injected_trace.argmax()] < 0):
                        first_mean = trace[interval_begin:interval_end].mean()
                    else:
                        first_mean = "Error"
                elif (i == 1): 
                    if (injected_trace[injected_trace.argmax()] < 0):
                        second_mean = trace[interval_begin:interval_end].mean()
                    else:
                        second_mean = "Error"
                i = i + 1
             
                coordinate = "B" + str(36+i)
                if (injected_trace[injected_trace.argmax()] > 0):
                    if sweep == -1:
                        ws1['B28'] = i
                        sweep = i
                        # Determine Overshoot
                        overshoot = 0
                        afterhyperpolarization = 0
                        firstPeak = peakdetect(injected_trace, "positive", None, lookahead = 300, delta=0)[0]
                        ws1['B30'] = firstPeak[1]

                        # Determine Afterhyperpolarization
                        afterhyperpolarization_trace = injected_trace[firstPeak[0]:firstPeak[0]+(80/rec.dt)]
                        afterhyperpolarization = peakdetect(afterhyperpolarization_trace, "negative", None, lookahead = 300, delta=0) 

                        ws1['B31'] = afterhyperpolarization     
                        # Determine Spike Height
                        ws1['B32'] = "=B30-B31"
                        #ws1['B32'] = firstPeak[1] - afterhyperpolarization
                    ws1[coordinate] = len(peakdetect(injected_trace, "positive", None, lookahead = 300, delta=5))
                else:
                    ws1[coordinate] = 0
        
        ws1['B8'] = '' # leert input resistance, falls der Wert nicht korrekt berechnet werden kann, steht ansonsten noch der input resistance der letzten Zelle drin.
        ws1['B8'] = inducedActivity().calculateInputResistance(first_mean, second_mean)  
        ws1['B27'] = iAP_file
        print "Current step"
        wb.save(filename = dest_filename)
            
    
        
    def onOpenGapFree(self):
        completeList = []
        global rec, complete_dataset, abffile, dest_filename
        ftypes = [('Axon binary files', '*.abf'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes, initialdir = openDirectory)
        fl = dlg.show()
        
        if fl != '':
            rec = stfio.read(str(fl))
            path = str(fl)
            
            #Extract filename for Excel file
            singles = path.split("/")
            abffile = singles[-1]
            abffile = abffile.split(".")
            ws1['B3'] = singles[-3]+"/"+singles[-2]
            

            # Read sampling interval in ms -> 0.2 ms for testing file
            # 1 section mit 256 Datenpunkten entspricht 51 ms bei 0.2 sampling interval
            # print(rec.dt)
            
            # constructs single numpy array from recording
            for i in range(0,len(rec[0])):
                completeList.extend(rec[0][i])
            complete_dataset = np.array(completeList)
            
            
            # noch nicht ganz sauber, muss in separate Klasse
            ws1['B7'] = '' #Feld für RMP auf leer setzen, damit nicht der Wert der vorherigen Zelle eingetragen wird, falls Zelle sAP oder attempted sAP zeigt.
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

            ws1['B16'] = "=SUM(B13:B15)"
            ws1['B6'] = rig
            ws1['B5'] = abffile[0] + "." + abffile[1]
            dest_filename = dest_directory + "\\"  + abffile[0] + ".xlsx"
            wb.save(filename = dest_filename)
            print "Gap free"
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

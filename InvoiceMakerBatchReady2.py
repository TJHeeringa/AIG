import subprocess
import csv
import os
import re
import sys

from Tkinter import *
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory

######
# DUMP the sys.exits
# add the paymentPeriod
# add help
#
######


# ###
# Setup GUI
# ###
class Application(Frame):
    def create_pdf(self,input_filename, output_filename):
        process = subprocess.Popen([
            'latex',   # Or maybe 'C:\\Program Files\\MikTex\\miktex\\bin\\latex.exe
            '-output-format=pdf',
            '-job-name=' + output_filename,
            input_filename])
        process.wait()

    def numberCheckOn2Decimal(self, number, rowCount):
        try:
            number = "{0:.2f}".format(float(number))
        except:
            print "There is something wrong with the numbers on row: "+str(rowCount)+"."
        finally:
            return number

    def openCSV(self):
        self.csvFile = askopenfilename(filetypes=[("CSV", "*.csv")])
        self.EntryText.set(str(self.csvFile))
        if (self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != ""):
            self.button4.config(state='normal')

    def openTemplate(self):
        self.template = askopenfilename(filetypes=[("Tex","*.tex")])
        self.EntryText2.set(str(self.template))
        if (self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != ""):
            self.button4.config(state='normal')

    def openDestination(self):
        self.destinationFolder = askdirectory()
        self.EntryText3.set(str(self.destinationFolder))
        if (self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != ""):
            self.button4.config(state='normal')

    def setDefaultTemplate(self):
            if self.CheckVar8.get() == 1:
                self.EntryText2.set("Default Invoice Used")
                self.button2.config(state='disabled')
                self.EntryBox2.config(state='disabled')
            else:
                self.EntryText2.set("")
                self.button2.config(state='normal')
                self.EntryBox2.config(state='normal')

    def denyOpeningLogFile(self):
        if self.CheckVar.get() == 0:
            self.CheckVar2.set(0) 

    def generateInvoices(self):
        logFile = open(self.destinationFolder+"/logFile.txt", "w")
        if self.CheckVar8.get() == 1:
            try:
                templateContent = open(os.path.dirname(sys.argv[0])+"\standaard_factuur.tex").read()
            except:
                try:
                    templateContent = open(os.path.dirname(sys.argv[0])+"\default_invoice.tex").read()
                except:
                    print "FATAL ERROR: No default invoice present in the application folder."
                    sys.exit("FATAL ERROR: No default invoice present in the application folder.")
        else:
            templateContent = open(self.template).read()
        rowCount = self.rowCount
        if self.CheckVar3.get() == 0:
            with open(self.csvFile, 'rb') as csvfile:
                spamreader = csv.reader(csvfile, delimiter=self.delimiterString, quotechar='|')
                for row in spamreader:
                    '''
                    In the first round the loop will gather the location of the keywords in the csvfile.
                    '''        
                    if rowCount == 0:
                        # add something for the payment period
                        try:
                            voornaamIndex = row.index("Voornaam")
                            tussenvoegselIndex =  row.index("Tussenvoegsel")
                            achternaamIndex = row.index("Achternaam")
                            initialenIndex = row.index("Initialen")
                            adresIndex = row.index("Adres")
                            postcodeIndex = row.index("Postcode")
                            plaatsIndex = row.index("Plaats")
                        except:
                            print "FATAL ERROR: There is an error in the names for the columns for the userdata."
                            sys.exit("FATAL ERROR: There is an error in the names for the columns for the userdata.")
                        try:
                            factuurnummerIndex = row.index("Factuurnummer")
                        except:
                            print "FATAL ERROR: You forgot to add the column with invoice numbers called: Factuurnummer."
                            sys.exit("FATAL ERROR: You forgot to add the column with invoice numbers called: Factuurnummer.")
                        try:
                            onderwerpIndex = row.index("Onderwerp")
                        except:
                            print "FATAL ERROR: You forgot to add the column with the subject called: Onderwerp."
                            sys.exit("FATAL ERROR: You forgot to add the column with the subject called: Onderwerp.")
                        try:
                            zaakIndex = row.index("Zaak")
                        except:
                            print "FATAL ERROR: You forgot to add the column with description for the invoice called: Zaak."
                            sys.exit("FATAL ERROR: You forgot to add the column with description for the invoice called: Zaak.")
                        try:
                            bedragIndex = row.index("Bedrag")
                        except:
                            print "FATAL ERROR: You forgot to add the column with the amounts of the bills called: Bedrag."
                            sys.exit("FATAL ERROR: You forgot to add the column with the amounts of the bills called: Bedrag.")
                        try:
                            postIndex = row.index("Post")
                            postIndexPresent = True
                        except:
                            postIndex = onderwerpIndex
                            postIndexPresent = False
                            print "NOTICE:  You did not add a column for the entries, the subject will be used as the entry."
                            logFile.write("NOTICE:  You did not add a column for the entries, the subject will be used as the entry. \n")
                        rowCount = rowCount+1
                        '''
                        For each row with data this part will get the data from the proper columns and place them in the correct part of the selected template.
                        '''
                    else:
                        '''
                        Reset step
                        '''
                        totaalbedrag = 0
                        bedragLine = ""


                        '''
                        Mail loop
                        '''
                        try:
                            voornaam = row[voornaamIndex]
                            tussenvoegsel = row[tussenvoegselIndex]
                            achternaam = row[achternaamIndex]
                            initialen = row[initialenIndex]
                            adres = row[adresIndex]
                            postcode = row[postcodeIndex]
                            plaats = row[plaatsIndex]
                        except:
                            print "ERROR: There is something wrong with the userdata on row: "+str(rowCount)+"."
                            logFile("ERROR: There is something wrong with the userdata on row: "+str(rowCount)+". \n")
                        factuurnummer = row[factuurnummerIndex]
                        onderwerp = row[onderwerpIndex]
                        zaak = row[zaakIndex]
                        bedrag = row[bedragIndex].split(",")
                        # add something for the Payment period
                        logFile.write("hoi")
                        newFileContent = templateContent.replace("adres}{adres","adres}{"+initialen+" "+achternaam+r'\\'+adres+r'\\'+postcode+" "+plaats)
                        try:
                            if self.CheckVar4.get() == 1:
                                if len(factuurnummer)+1 != int(self.EntryVar.get()): # An invoice number has a default length of 16 digits. 
                                    print "WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this."
                                    logFile.write("WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this. \n")
                                    try:
                                        numberFormat = "{0:0"+self.EntryVar.get()+"d}"
                                        factuurnummer = numberFormat.format(int(factuurnummer))
                                        print "NOTICE:  The invoice number on row "+str(rowCount)+" has been corrected."
                                        logFile.write("NOTICE:  The invoice number on row "+str(rowCount)+" has been corrected. \n")
                                    except:
                                        print "ERROR:   Attempt to correct invoice number has failed."
                                        logFile.write("ERROR:   Attempt to correct invoice number has failed. \n")
                        except:
                            print "WARNING: You made an error in writing down the invoice number on row: "+str(rowCount)+"."
                            logFile.write("WARNING: You made an error in writing down the invoice number on row: "+str(rowCount)+".")
                        newFileContent = newFileContent.replace("factuurnummer}{factuurnummer","factuurnummer}{"+factuurnummer)
                        newFileContent = newFileContent.replace("betreftregel}{onderwerp","betreftregel}{"+onderwerp)
                        newFileContent = newFileContent.replace("voornaam}{voornaam","voornaam}{"+voornaam)
                        newFileContent = newFileContent.replace("zaak}{zaak","zaak}{"+zaak)
                        if postIndexPresent == True:
                            try:
                                post = row[postIndex].split(self.blockDelimiterString)
                                if len(bedrag)!=len(post):
                                    print "WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+"."
                                    logFile.write("WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+". \n")
                                for i in range(0,len(bedrag)):
                                    if post[i]=='':
                                        print "WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry."
                                        logFile.write("WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry. \n")
                                        post[i]=onderwerp
                                    bedragLine = bedragLine + post[i] +r" & \euro & "+self.numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline "
                                    totaalbedrag = totaalbedrag + float(bedrag[i])
                                newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                            except:
                                try:
                                    for i in range(0,len(bedrag)):
                                        bedragLine = bedragLine + r"\betreftregel & \euro & "+self.numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline"
                                        totaalbedrag = totaalbedrag + float(bedrag[i])
                                    newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                                    print "ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+"."
                                    logFile.write("ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+". \n")
                                except:
                                    print "ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+"."
                                    logFile.write("ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+". \n")
                        else:
                            try:
                                bedragLine = bedragLine + r"\betreftregel & \euro & "+self.numberCheckOn2Decimal(bedrag[0],rowCount)+r" \\[6pt]\hline"
                                totaalbedrag = totaalbedrag + float(bedrag[0])
                                newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                                if len(bedrag)!=1:
                                    print "WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used."
                                    logFile.write("WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used. \n")
                            except:
                                print "ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+"."
                                logFile.write("ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+". \n")
                        if self.CheckVar2.get() == 1:
                            newFileContent = newFileContent.replace("paymentperiod}{14","paymentperiod}{"+self.EntryVar2.get())
                        totaalbedrag = self.numberCheckOn2Decimal(totaalbedrag,rowCount)
                        newFileContent = newFileContent.replace("totaalbedrag",str(totaalbedrag))

                        '''
                        Create temporary .tex file to store the edited template The temporary file will be made to an .pdf and can be deleted afterwards.
                        '''
                        # change this to something more general and usefull as this doenst really work with Mailmerge
                        if tussenvoegsel == "":
                            newFile = open(self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+achternaam+".tex","a+")
                            newFile.write(newFileContent)
                            newFile.close()
                            self.create_pdf(self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+achternaam+".tex",self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+achternaam)
                        else:
                            newFile = open(self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+tussenvoegsel+" "+achternaam+".tex","a+")
                            newFile.write(newFileContent)
                            newFile.close()
                            self.create_pdf(self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+tussenvoegsel+" "+achternaam+".tex",self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+tussenvoegsel+" "+achternaam)
                        if self.CheckVar6.get() == 0:
                            os.remove(self.destinationFolder+"/"+onderwerp+" "+voornaam+" "+tussenvoegsel+" "+achternaam+".tex")
                        rowCount = rowCount+1
            logFile.close()
            if self.CheckVar2.get() == 1:
                os.startfile(self.destinationFolder+"/logFile.txt")
            if self.CheckVar.get() == 0:
                os.remove(self.destinationFolder+"/logFile.txt")
    
    def createWidgets(self):
        '''
        Help label
        '''
        self.helpFrame = LabelFrame(root,width=350, text="HELP")
        self.helpFrame.grid(row=0,column=9, columnspan=2, rowspan=8, sticky='NSE', padx=5, pady=5, ipadx=5, ipady=5)
        label = Label(self.helpFrame, text = self.helpText).grid(row=0, column=1, sticky="E")

        '''
        Top-left label
        '''
        self.labelFrame = LabelFrame(root, width=600,height=200, text="Enter File Details:")
        self.labelFrame.grid(row=0, column=0, sticky='WE', padx=5, pady=5, ipadx=5, ipady=5)
        Label(self.labelFrame, text=self.labelText).grid(row=0, column=1, sticky="W")
        Label(self.labelFrame, text=self.labelText2).grid(row=1, column=1, sticky="W")
        Label(self.labelFrame, text=self.labelText3).grid(row=2, column=1, sticky="W")
        Entry(self.labelFrame, text=self.EntryText,width=60).grid(row=0, column=9, columnspan=100, sticky="WE", pady=3)
        self.EntryBox2 = Entry(self.labelFrame, text = self.EntryText2,state=DISABLED)
        self.EntryBox2.grid(row=1, column=9, columnspan=100, sticky="WE", pady=3)
        Entry(self.labelFrame, text = self.EntryText3).grid(row=2, column =9,columnspan=100, sticky="WE", pady=3)
        Button(self.labelFrame, text="Browse File", command=self.openCSV).grid(row=0, column =140,sticky='WE', padx=5, pady=2)
        self.button2 = Button(self.labelFrame, text="Browse File", command=self.openTemplate,state=DISABLED)
        self.button2.grid(row=1, column=140, sticky='WE', padx=5, pady=2)
        Button(self.labelFrame, text="Browse Folder", command=self.openDestination).grid(row=2, column=140,sticky='WE', padx=5, pady=2)

        '''
        Middle-left label
        '''
        self.labelFrame2 = LabelFrame(root,width=600,height=150, text="Enter Options:")
        self.labelFrame2.grid_propagate(False)
        self.labelFrame2.grid(row=1, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)
        Checkbutton(self.labelFrame2, text="Create Logfile",variable=self.CheckVar,command=self.denyOpeningLogFile).grid(row=0,sticky="W")
        Checkbutton(self.labelFrame2, text="Open Logfile on finish",variable=self.CheckVar2).grid(row=1,sticky="W")
        Checkbutton(self.labelFrame2, text="Don't do anything",variable=self.CheckVar3).grid(row=2,sticky="W")
        Checkbutton(self.labelFrame2, text="Check whether the invoice number has",variable=self.CheckVar4).grid(row=3,sticky="W")
        Entry(self.labelFrame2, text=self.EntryVar).grid(sticky="W",row=3,column=2,columnspan=1)
        Label(self.labelFrame2, text = "digits and if not correct the number").grid(row=3, column =90,sticky="WE")
        Checkbutton(self.labelFrame2, text="Set the payment period to",variable=self.CheckVar5).grid(row=4,sticky="W")
        Entry(self.labelFrame2, text=self.EntryVar2).grid(sticky="W",row=4,column=2,columnspan=1)
        Label(self.labelFrame2, text = "days").grid(row=4, column =90,sticky="W")
        Checkbutton(self.labelFrame2, text="Keep .tex Files",variable=self.CheckVar6).grid(row=0,column=2,sticky="W")
        Checkbutton(self.labelFrame2, text="Remove subsidiary files",variable=self.CheckVar7).grid(row=1,column=2,sticky="W")
        Checkbutton(self.labelFrame2, text="Use default invoice",variable=self.CheckVar8,command=self.setDefaultTemplate).grid(row=2,column=2,sticky="W")

        '''
        Bottom-left label
        '''
        self.labelFrame3 = LabelFrame(root,width=600,height=50, text="Run:")
        self.labelFrame3.grid_propagate(False)
        self.labelFrame3.grid(row=2, column=0, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)
        self.button4 = Button(self.labelFrame3, text="Generate Invoices", command=self.generateInvoices, state=DISABLED)
        self.button4.grid(padx=5, pady=5,row=30, rowspan=400, column=40, sticky="W")

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid_propagate(False)
        self.grid()
        '''
        General config
        '''
        self.delimiterString = ";"
        self.blockDelimiterString = ","
        self.totaalbedrag = 0
        self.bedragLine = ""
        self.rowCount = 0
        '''
        Static labels
        '''
        self.helpText = "To those who seek help and to those who are bored"
        self.labelText = "Select .csv File With Data:"
        self.labelText2 = "Import Template File:"
        self.labelText3 = "Save Files to:"
        '''
        Variable boxes
        '''
        self.EntryText = StringVar()
        self.EntryText2 = StringVar()
        self.EntryText3 = StringVar()
        self.CheckVar = IntVar()
        self.CheckVar2 = IntVar()
        self.CheckVar3 = IntVar()
        self.CheckVar4 = IntVar()
        self.CheckVar5 = IntVar()
        self.CheckVar6 = IntVar()
        self.CheckVar7 = IntVar()
        self.CheckVar8 = IntVar()
        self.EntryVar = StringVar()
        self.EntryVar2 = StringVar()
        '''
        Variable boxes default values
        '''
        self.EntryText2.set("Default Invoice Used")
        self.EntryVar.set("16")
        self.EntryVar2.set("14")
        self.CheckVar.set(1)
        self.CheckVar2.set(1)
        self.CheckVar3.set(0)
        self.CheckVar4.set(1)
        self.CheckVar5.set(0)
        self.CheckVar6.set(1)
        self.CheckVar7.set(1)
        self.CheckVar8.set(1)
        self.createWidgets()

        
root = Tk()
root.title("Automatic Invoice Generator")
root.geometry("925x370")
app = Application(master=root)
app.mainloop()

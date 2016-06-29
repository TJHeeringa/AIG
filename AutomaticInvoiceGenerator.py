import subprocess
import csv
import os
import re
import sys

from Tkinter import *
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory
import tkMessageBox

######
# add help
######


# ###
# Setup GUI
# ###
class Application(Frame):

    def create_pdf(self, input_filename, output_directory):
        process = subprocess.Popen([
            'pdflatex',   # Or maybe 'C:\\Program Files\\MikTex\\miktex\\bin\\latex.exe
            '-interaction=nonstopmode',
            '-output-directory='+output_directory,
            input_filename])
        process.wait()

    def numberCheckOn2Decimal(self, number, rowCount, logFile):
        try:
            number = "{0:.2f}".format(float(number))
        except:
            logFile.write("There is something wrong with the numbers on row: " + str(rowCount) + ".")
        finally:
            return number

    def openCSV(self):
        self.csvFile = askopenfilename(filetypes=[("CSV", "*.csv")])
        self.EntryText.set(str(self.csvFile))
        if self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != "":
            self.button4.config(state='normal')

    def openTemplate(self):
        self.template = askopenfilename(filetypes=[("Tex", "*.tex")])
        self.EntryText2.set(str(self.template))
        if self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != "":
            self.button4.config(state='normal')

    def openDestination(self):
        self.destinationFolder = askdirectory()
        self.EntryText3.set(str(self.destinationFolder))
        if self.EntryText.get() != "" and self.EntryText2.get() != "" and self.EntryText3.get() != "":
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

    def getAdresContent(self, userData):
        titel = userData["titel"]
        initialen = userData["initialen"]
        tussenvoegsel = userData["tussenvoegsel"]
        achternaam = userData["achternaam"]
        bedrijsnaam = userData["bedrijfsnaam"]
        adres = userData["adres"]
        postcode = userData["postcode"]
        plaats = userData["plaats"]

        if bedrijsnaam != "":
            adresContent = bedrijsnaam + r'\\' + "t.a.v. "
            if initialen != "" and achternaam != "":
                if titel != "":
                    if tussenvoegsel != "":
                        adresContent += titel + " " + initialen + " " + tussenvoegsel + " " + achternaam
                    else:
                        adresContent += titel + " " + initialen + " " + achternaam
                else:
                    if tussenvoegsel != "":
                        adresContent += initialen + " " + tussenvoegsel + " " + achternaam
                    else:
                        adresContent += initialen + " " + achternaam
            else:
                adresContent += "de crediteurenadministratie"
        else:
            if titel != "":
                if tussenvoegsel != "":
                    adresContent = titel + " " + initialen + " " + tussenvoegsel + " " + achternaam
                else:
                    adresContent = titel + " " + initialen + " " + achternaam
            else:
                if tussenvoegsel != "":
                    adresContent = initialen + " " + tussenvoegsel + " " + achternaam
                else:
                    adresContent = initialen + " " + achternaam
        adresContent += r'\\' + adres + r'\\' + postcode + " " + plaats
        return adresContent

    def getAanhefContent(self, userData):
        titel = userData["titel"]
        initialen = userData["initialen"]
        voornaam = userData["voornaam"]
        tussenvoegsel = userData["tussenvoegsel"]
        achternaam = userData["achternaam"]
        if voornaam != "":
            aanhefContent = voornaam
        else:
            if titel != "":
                if tussenvoegsel != "":
                    aanhefContent = titel + " " + initialen + " " + tussenvoegsel + " " + achternaam
                else:
                    aanhefContent = titel + " " + initialen + " " + achternaam
            else:
                if achternaam == "":
                    aanhefContent = "heer, mevrouw"
        return aanhefContent

    def getBetalingstermijnContent(self, betalingstermijn):
        if self.CheckVar5.get() == 1:
            betalingstermijnContent = self.EntryVar2.get()
        else:
            if betalingstermijn != "":
                betalingstermijnContent = betalingstermijn
            else:
                betalingstermijnContent = self.EntryVar2.get()
        return betalingstermijnContent

    def getFactuurnummerContent(self, factuurnummer, rowCount, logFile):
        if self.CheckVar4.get() == 1:
            try:
                """
                An invoice number has a default length of 16
                digits. Self.EntryVar contains the lenght of the invoice number
                This part will check and correct the length of the number
                """
                if len(factuurnummer) + 1 != int(self.EntryVar.get()):
                    logFile.write("WARNING: The invoice number has not the proper amount of digits on row: " + str(
                        rowCount) + ". An attempt will be made to correct this. \n")
                    try:
                        numberFormat = "{0:0" + self.EntryVar.get() + "d}"
                        factuurnummer = numberFormat.format(int(factuurnummer))
                        logFile.write(
                            "NOTICE:  The invoice number on row " + str(rowCount) + " has been corrected. \n")
                    except:
                        logFile.write(
                            "ERROR:   Attempt to correct invoice number on row" + str(rowCount) + "has failed. \n")
            except:
                logFile.write(
                    "WARNING: You made an error in writing down the invoice number on row: " + str(rowCount) + ".")
        return factuurnummer

    def getTabelContent(self, bedrag, post, onderwerp, rowCount, logFile):
        bedragLine = ""
        totaalbedrag = 0
        if post != [""]:
            try:
                if len(bedrag) != len(post):
                    logFile.write(
                        "WARNING: The number of entries and amounts are unequal on row: " + str(rowCount) + ". \n")
                for i in range(0, len(bedrag)):
                    if post[i] == '':
                        logFile.write("WARNING: The " + str(i + 1) + "th entry is empty on row: " + str(
                            rowCount) + ". Subject will be used as entry. \n")
                        post[i] = onderwerp
                    bedragLine += post[i] + r" & \euro & " + self.numberCheckOn2Decimal(
                        bedrag[i], rowCount, logFile) + r" \\[6pt]\hline "
                    totaalbedrag += float(bedrag[i])
            except:
                try:
                    for i in range(0, len(bedrag)):
                        bedragLine += r"\betreftregel & \euro & " + self.numberCheckOn2Decimal(
                            bedrag[i], rowCount, logFile) + r" \\[6pt]\hline"
                        totaalbedrag += float(bedrag[i])
                    logFile.write(
                        "ERROR:   You made an error in writing down the entries on row: " + str(rowCount) + ". \n")
                except:
                    logFile.write(
                        "ERROR:   You made an error in writing down the amounts on row: " + str(rowCount) + ". \n")
        else:
            try:
                bedragLine += r"\betreftregel & \euro & " + \
                              self.numberCheckOn2Decimal(
                                  bedrag[0], rowCount, logFile) + r" \\[6pt]\hline"
                totaalbedrag += float(bedrag[0])
                if len(bedrag) != 1:
                    logFile.write(
                        "WARNING: You wrote multiple amounts eventhough there are no different entries on row: " + str(
                            rowCount) + ". Only the first number will be used. \n")
            except:
                logFile.write(
                    "ERROR:   You made an error in writing down the amount on row: " + str(rowCount) + ". \n")
        tabelContent = {
            "bedragLine": bedragLine,
            "totaalbedrag": str(self.numberCheckOn2Decimal(totaalbedrag, rowCount, logFile))
        }
        return tabelContent

    def getPonummerContent(self, ponummer):
        if ponummer != "":
            ponummerContent = "\uwkenmerk{" + ponummer + "}"
        else:
            ponummerContent = "%\uwkenmerk{}"
        return ponummerContent

    def generateInvoices(self):
        """
        Set log destination
        """
        logFile = open(self.destinationFolder + "/logFile.txt", "w")
        """
        Set the template to use
        """
        if self.CheckVar8.get() == 1:
            try:
                templateContent = open(os.path.dirname(
                    sys.argv[0]) + "\standaard_factuur.tex").read()
            except:
                try:
                    templateContent = open(os.path.dirname(
                        sys.argv[0]) + "\default_invoice.tex").read()
                except:
                    tkMessageBox.showerror("Fatal Error",
                        "FATAL ERROR: No default invoice present in the application folder.")
        else:
            templateContent = open(self.template).read()
        """
        Look if the code should do anything
        """
        if self.CheckVar3.get() == 0:
            rowCount = 1
            with open(self.csvFile, 'rb') as csvfile:
                spamreader = csv.reader(
                    csvfile, delimiter=self.delimiterString)
                for row in spamreader:
                    """
                    In the first round the loop will gather the location of the keywords in the csvfile.
                    """
                    if rowCount == 1:
                        try:
                            titelIndex = row.index("Titel")
                            voornaamIndex = row.index("Voornaam")
                            initialenIndex = row.index("Initialen")
                            tussenvoegselIndex = row.index("Tussenvoegsel")
                            achternaamIndex = row.index("Achternaam")
                            bedrijfsnaamIndex = row.index("Bedrijfsnaam")
                            adresIndex = row.index("Adres")
                            postcodeIndex = row.index("Postcode")
                            plaatsIndex = row.index("Plaats")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: There is an error in the names for the columns for the userdata.")
                        try:
                            betalingstermijnIndex = row.index("Betalingstermijn")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with invoice numbers called: Betalingstermijn.")
                        try:
                            ponummerIndex = row.index("PO-nummer")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with invoice numbers called: PO-nummer.")
                        try:
                            factuurnummerIndex = row.index("Factuurnummer")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with invoice numbers called: Factuurnummer.")
                        try:
                            onderwerpIndex = row.index("Onderwerp")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with the subject called: Onderwerp.")
                        try:
                            zaakIndex = row.index("Zaak")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with description for the invoice called: Zaak.")
                        try:
                            bedragIndex = row.index("Bedrag")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with the amounts of the bills called: Bedrag.")
                        try:
                            postIndex = row.index("Post")
                        except:
                            tkMessageBox.showerror("Fatal Error",
                                "FATAL ERROR: You forgot to add the column with the entries for the bills: Post.")
                        rowCount += 1
                        """
                        For each row with data this part will get the data from the proper columns and place them in the correct part of the selected template.
                        """
                    else:
                        """
                        Retrieve the data from the .csv
                        """
                        try:
                            titel = row[titelIndex].strip()
                            voornaam = row[voornaamIndex].strip()
                            initialen = row[initialenIndex].strip()
                            tussenvoegsel = row[tussenvoegselIndex].strip()
                            achternaam = row[achternaamIndex].strip()
                            bedrijfsnaam = row[bedrijfsnaamIndex].strip()
                            adres = row[adresIndex].strip()
                            postcode = row[postcodeIndex].strip()
                            plaats = row[plaatsIndex].strip()
                        except:
                            logFile.write(
                                "ERROR: There is something wrong with the userdata on row: " + str(rowCount) + ". \n")
                        factuurnummer = row[factuurnummerIndex].strip()
                        onderwerp = row[onderwerpIndex].strip()
                        zaak = row[zaakIndex].strip()
                        bedrag = row[bedragIndex].split(self.blockDelimiterString)
                        post = row[postIndex].split(self.blockDelimiterString)
                        betalingstermijn = row[betalingstermijnIndex].strip()
                        ponummer = row[ponummerIndex].strip()
                        """
                        Transform the gathered data to something that can be entered in the template
                        """
                        userData = {
                            "titel": titel,
                            "voornaam": voornaam,
                            "initialen": initialen,
                            "tussenvoegsel":tussenvoegsel,
                            "achternaam": achternaam,
                            "bedrijfsnaam": bedrijfsnaam,
                            "adres": adres,
                            "postcode": postcode,
                            "plaats": plaats,
                        }
                        adresContent = self.getAdresContent(userData)
                        aanhefContent = self.getAanhefContent(userData)
                        betalingstermijnContent = self.getBetalingstermijnContent(betalingstermijn)
                        factuurnummerContent = self.getFactuurnummerContent(factuurnummer, rowCount, logFile)
                        tabelContent = self.getTabelContent(bedrag, post, onderwerp, rowCount, logFile)
                        ponummerContent = self.getPonummerContent(ponummer)
                        """
                        Replace the content of the data in the invoice template
                        """
                        newFileContent = templateContent.replace(
                            "adres}{adres", "adres}{" + adresContent)
                        newFileContent = newFileContent.replace(
                            "betreftregel}{onderwerp", "betreftregel}{" + onderwerp)
                        newFileContent = newFileContent.replace(
                            "voornaam}{voornaam", "voornaam}{" + aanhefContent)
                        newFileContent = newFileContent.replace(
                            "zaak}{zaak", "zaak}{" + zaak)
                        newFileContent = newFileContent.replace(
                            "paymentperiod}{14", "paymentperiod}{" + betalingstermijnContent)
                        newFileContent = newFileContent.replace(
                            "factuurnummer}{factuurnummer", "factuurnummer}{" + factuurnummerContent)
                        newFileContent = newFileContent.replace(
                            r"\betreftregel & \euro & bedrag \\[6pt]\hline", tabelContent["bedragLine"])
                        newFileContent = newFileContent.replace(
                            "totaalbedrag", tabelContent["totaalbedrag"])
                        newFileContent = newFileContent.replace(
                            "%\uwkenmerk{}", ponummerContent)

                        """
                        Create temporary .tex file to store the edited template.
                        The temporary file will be made to an .pdf and can be deleted afterwards.
                        """
                        newFile = open(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".tex", "a+")
                        newFile.write(newFileContent)
                        newFile.close()
                        self.create_pdf(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".tex",
                                        self.destinationFolder + "/")
                        if self.CheckVar6.get() == 0:
                            os.remove(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".tex")
                        rowCount += 1
                        if self.CheckVar7.get() == 1:
                            os.remove(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".aux")
                            os.remove(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".log")
                            os.remove(self.destinationFolder + "/" + "F." + factuurnummerContent + " " + onderwerp + ".out")
            logFile.close()
            if self.CheckVar2.get() == 1:
                os.startfile(self.destinationFolder + "/logFile.txt")
            if self.CheckVar.get() == 0:
                os.remove(self.destinationFolder + "/logFile.txt")
            tkMessageBox.showinfo("AIG", "Invoices have been generated")

    def createWidgets(self):
        """
        Help label
        """
        self.helpFrame = LabelFrame(root, width=350, text="HELP")
        self.helpFrame.grid(row=0, column=9, columnspan=2, rowspan=8,
                            sticky='NSE', padx=5, pady=5, ipadx=5, ipady=5)
        label = Label(self.helpFrame, text=self.helpText).grid(
            row=0, column=1, sticky="E")

        '''
        Top-left label
        '''
        self.labelFrame = LabelFrame(
            root, width=600, height=200, text="Enter File Details:")
        self.labelFrame.grid(row=0, column=0, sticky='WE',
                             padx=5, pady=5, ipadx=5, ipady=5)
        Label(self.labelFrame, text=self.labelText).grid(
            row=0, column=1, sticky="W")
        Label(self.labelFrame, text=self.labelText2).grid(
            row=1, column=1, sticky="W")
        Label(self.labelFrame, text=self.labelText3).grid(
            row=2, column=1, sticky="W")
        Entry(self.labelFrame, text=self.EntryText, width=60).grid(
            row=0, column=9, columnspan=100, sticky="WE", pady=3)
        self.EntryBox2 = Entry(
            self.labelFrame, text=self.EntryText2, state=DISABLED)
        self.EntryBox2.grid(row=1, column=9, columnspan=100,
                            sticky="WE", pady=3)
        Entry(self.labelFrame, text=self.EntryText3).grid(
            row=2, column=9, columnspan=100, sticky="WE", pady=3)
        Button(self.labelFrame, text="Browse File", command=self.openCSV).grid(
            row=0, column=140, sticky='WE', padx=5, pady=2)
        self.button2 = Button(
            self.labelFrame, text="Browse File", command=self.openTemplate, state=DISABLED)
        self.button2.grid(row=1, column=140, sticky='WE', padx=5, pady=2)
        Button(self.labelFrame, text="Browse Folder", command=self.openDestination).grid(
            row=2, column=140, sticky='WE', padx=5, pady=2)

        '''
        Middle-left label
        '''
        self.labelFrame2 = LabelFrame(
            root, width=600, height=150, text="Enter Options:")
        self.labelFrame2.grid_propagate(False)
        self.labelFrame2.grid(row=1, column=0, sticky='W',
                              padx=5, pady=5, ipadx=5, ipady=5)
        Checkbutton(self.labelFrame2, text="Create Logfile", variable=self.CheckVar,
                    command=self.denyOpeningLogFile).grid(row=0, sticky="W")
        Checkbutton(self.labelFrame2, text="Open Logfile on finish",
                    variable=self.CheckVar2).grid(row=1, sticky="W")
        Checkbutton(self.labelFrame2, text="Don't do anything",
                    variable=self.CheckVar3).grid(row=2, sticky="W")
        Checkbutton(self.labelFrame2, text="Check whether the invoice number has",
                    variable=self.CheckVar4).grid(row=3, sticky="W")
        Entry(self.labelFrame2, text=self.EntryVar).grid(
            sticky="W", row=3, column=2, columnspan=1)
        Label(self.labelFrame2, text="digits and if not correct the number").grid(
            row=3, column=90, sticky="WE")
        Checkbutton(self.labelFrame2, text="Set the payment period to",
                    variable=self.CheckVar5).grid(row=4, sticky="W")
        Entry(self.labelFrame2, text=self.EntryVar2).grid(
            sticky="W", row=4, column=2, columnspan=1)
        Label(self.labelFrame2, text="days").grid(row=4, column=90, sticky="W")
        Checkbutton(self.labelFrame2, text="Keep .tex Files",
                    variable=self.CheckVar6).grid(row=0, column=2, sticky="W")
        Checkbutton(self.labelFrame2, text="Remove subsidiary files",
                    variable=self.CheckVar7).grid(row=1, column=2, sticky="W")
        Checkbutton(self.labelFrame2, text="Use default invoice", variable=self.CheckVar8,
                    command=self.setDefaultTemplate).grid(row=2, column=2, sticky="W")

        '''
        Bottom-left label
        '''
        self.labelFrame3 = LabelFrame(root, width=600, height=50, text="Run:")
        self.labelFrame3.grid_propagate(False)
        self.labelFrame3.grid(row=2, column=0, sticky='W',
                              padx=5, pady=5, ipadx=5, ipady=5)
        self.button4 = Button(self.labelFrame3, text="Generate Invoices",
                              command=self.generateInvoices, state=DISABLED)
        self.button4.grid(padx=5, pady=5, row=30,
                          rowspan=400, column=40, sticky="W")

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid_propagate(False)
        self.grid()
        '''
        General config
        '''
        self.delimiterString = ";"
        self.blockDelimiterString = "|"
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

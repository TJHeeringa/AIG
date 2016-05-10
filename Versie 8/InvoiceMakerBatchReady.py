import subprocess
import csv
import os
import re
import sys

from Tkinter import *
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory

def create_pdf(input_filename, output_filename):
    process = subprocess.Popen([
        'latex',   # Or maybe 'C:\\Program Files\\MikTex\\miktex\\bin\\latex.exe
        '-output-format=pdf',
        '-job-name=' + output_filename,
        input_filename])
    process.wait()

def numberCheckOn2Decimal(number,rowCount):
    try:
        number = "{0:.2f}".format(float(number))
    except:
        print "There is something wrong with the numbers on row: "+str(rowCount)+"."
    finally:
        return number

'''
Constructor
'''
delimiterString = ","
blockDelimiterString = ";"
totaalbedrag = 0
bedragLine = ""
rowCount = 0

'''
Setup GUI
'''
class Application(Frame):

    def openCSV(self):
        self.csvFile = askopenfilename(filetypes=[("CSV","*.csv")])

    def openTemplate(self):
        self.template = open(askopenfilename(filetypes=[("Tex","*.tex")]))

    def openDestination(self):
        self.destinationFolder = open(askopenfilename(filetypes=[("Tex","*.tex")]))

    def generateInvoices(self):
        logFile = open(destinationFolder+"/logFile.txt","w")
        with open(csvFile, 'rb') as csvfile:
            spamreader = csv.reader(csvfile, delimiter=delimiterString, quotechar='|')
            for row in spamreader:
                '''
                In the first round the loop will gather the location of the keywords in the csvfile.
                '''        
                if rowCount == 0:
                    try:
                        voornaamIndex = row.index("Voornaam")
                        tussenvoegselIndex =  row.index("Tussenvoegsel")
                        achternaamIndex = row.index("Achternaam")
                        initialenIndex = row.index("Initialen")
                        adresIndex = row.index("Adres")
                        postcodeIndex = row.index("Postcode")
                        plaatsIndex = row.index("Stad")
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
                    try:
                        termijnIndex = row.index("Termijn")
                        termijnIndexPresent = True
                    except:
                        termijnIndexPresent = False
                        print "NOTICE:  You did not add a column for the payment periods, the default payment period specified in the template will be used."
                        logFile.write("NOTICE:  You did not add a column for the payment periods, the default payment period specified in the template will be used. \n")
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
                    bedrag = row[bedragIndex].split(";")
                    newFileContent = templateContent.replace("adres}{adres","adres}{"+initialen+" "+achternaam+r'\\'+adres+r'\\'+postcode+" "+plaats)
                    try:
                        if len(factuurnummer)+1 != 16: # An invoice number has a default length of 16 digits. 
                            print "WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this."
                            logFile.write("WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this. \n")
                            try:
                                factuurnummer = "{0:016d}".format(int(factuurnummer))
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
                            post = row[postIndex].split(";")
                            if len(bedrag)!=len(post):
                                print "WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+"."
                                logFile.write("WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+". \n")
                            for i in range(0,len(bedrag)):
                                if post[i]=='':
                                    print "WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry."
                                    logFile.write("WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry. \n")
                                    post[i]=onderwerp
                                bedragLine = bedragLine + post[i] +r" & \euro & "+numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline "
                                totaalbedrag = totaalbedrag + float(bedrag[i])
                            newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                        except:
                            try:
                                for i in range(0,len(bedrag)):
                                    bedragLine = bedragLine + r"\betreftregel & \euro & "+numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline"
                                    totaalbedrag = totaalbedrag + float(bedrag[i])
                                newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                                print "ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+"."
                                logFile.write("ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+". \n")
                            except:
                                print "ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+"."
                                logFile.write("ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+". \n")
                    else:
                        try:
                            bedragLine = bedragLine + r"\betreftregel & \euro & "+numberCheckOn2Decimal(bedrag[0],rowCount)+r" \\[6pt]\hline"
                            totaalbedrag = totaalbedrag + float(bedrag[0])
                            newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                            if len(bedrag)!=1:
                                print "WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used."
                                logFile.write("WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used. \n")
                        except:
                            print "ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+"."
                            logFile.write("ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+". \n")
                    totaalbedrag = numberCheckOn2Decimal(totaalbedrag,rowCount)
                    newFileContent = newFileContent.replace("totaalbedrag",str(totaalbedrag))

                    '''
                    Create temporary .tex file to store the edited template. The temporary file will be made to an .pdf and will be deleted afterwards.
                    '''
                    newFile = open(destinationFolder+"/newFile.tex","a+")
                    newFile.write(newFileContent)
                    newFile.close()
                    create_pdf(destinationFolder+"/newFile.tex","Testen/"+onderwerp+" "+voornaam+" "+achternaam)
                    os.remove(destinationFolder+"/newFile.tex")
                    rowCount = rowCount+1
        logFile.close()
        os.startfile(destinationFolder+"/logFile.txt")

    def createWidgets(self):
        self.button = Button(root,text="Csv File",command=self.openCSV)
        self.button2 = Button(root,text="Invoice Template",command=self.openTemplate)
        self.button3 = Button(root,text="Destinations Folder",command=self.openDestination)
        self.button4 = Button(root,text="Generate Invoices",command=self.generateInvoices)
        self.button.pack()
        self.button2.pack()
        self.button3.pack()
        self.button4.pack()
        
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.filename = None
        self.pack()
        self.createWidgets()
        
root = Tk()
root.title("Image Manipulation Program")
app = Application(master=root)
app.mainloop()

'''
Gather files and directories
'''
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
template = open(filename)
#template = open("standaard_factuur.tex")
templateContent = template.read()
csvFile = askopenfilename()
destinationFolder = askdirectory()
#destinationFolder = "Testen"
logFile = open(destinationFolder+"/logFile.txt","w")
#with open("test.csv", 'rb') as csvfile:
with open(csvFile, 'rb') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=delimiterString, quotechar='|')
    for row in spamreader:
        '''
        In the first round the loop will gather the location of the keywords in the csvfile.
        '''        
        if rowCount == 0:
            try:
                voornaamIndex = row.index("Voornaam")
                tussenvoegselIndex =  row.index("Tussenvoegsel")
                achternaamIndex = row.index("Achternaam")
                initialenIndex = row.index("Initialen")
                adresIndex = row.index("Adres")
                postcodeIndex = row.index("Postcode")
                plaatsIndex = row.index("Stad")
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
            try:
                termijnIndex = row.index("Termijn")
                termijnIndexPresent = True
            except:
                termijnIndexPresent = False
                print "NOTICE:  You did not add a column for the payment periods, the default payment period specified in the template will be used."
                logFile.write("NOTICE:  You did not add a column for the payment periods, the default payment period specified in the template will be used. \n")
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
            bedrag = row[bedragIndex].split(";")
            newFileContent = templateContent.replace("adres}{adres","adres}{"+initialen+" "+achternaam+r'\\'+adres+r'\\'+postcode+" "+plaats)
            try:
                if len(factuurnummer)+1 != 16: # An invoice number has a default length of 16 digits. 
                    print "WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this."
                    logFile.write("WARNING: The invoice number has not the proper amount of digits on row: "+str(rowCount)+". An attempt will be made to correct this. \n")
                    try:
                        factuurnummer = "{0:016d}".format(int(factuurnummer))
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
                    post = row[postIndex].split(";")
                    if len(bedrag)!=len(post):
                        print "WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+"."
                        logFile.write("WARNING: The number of entries and amounts are unequal on row: "+str(rowCount)+". \n")
                    for i in range(0,len(bedrag)):
                        if post[i]=='':
                            print "WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry."
                            logFile.write("WARNING: The "+str(i+1)+"th entry is empty on row: "+str(rowCount)+". Subject will be used as entry. \n")
                            post[i]=onderwerp
                        bedragLine = bedragLine + post[i] +r" & \euro & "+numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline "
                        totaalbedrag = totaalbedrag + float(bedrag[i])
                    newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                except:
                    try:
                        for i in range(0,len(bedrag)):
                            bedragLine = bedragLine + r"\betreftregel & \euro & "+numberCheckOn2Decimal(bedrag[i],rowCount)+r" \\[6pt]\hline"
                            totaalbedrag = totaalbedrag + float(bedrag[i])
                        newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                        print "ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+"."
                        logFile.write("ERROR:   You made an error in writing down the entries on row: "+str(rowCount)+". \n")
                    except:
                        print "ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+"."
                        logFile.write("ERROR:   You made an error in writing down the amounts on row: "+str(rowCount)+". \n")
            else:
                try:
                    bedragLine = bedragLine + r"\betreftregel & \euro & "+numberCheckOn2Decimal(bedrag[0],rowCount)+r" \\[6pt]\hline"
                    totaalbedrag = totaalbedrag + float(bedrag[0])
                    newFileContent = newFileContent.replace(r"\betreftregel & \euro & bedrag \\[6pt]\hline",bedragLine)
                    if len(bedrag)!=1:
                        print "WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used."
                        logFile.write("WARNING: You wrote multiple amounts eventhough there are no different entries on row: "+str(rowCount)+". Only the first number will be used. \n")
                except:
                    print "ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+"."
                    logFile.write("ERROR:   You made an error in writing down the amount on row: "+str(rowCount)+". \n")
            totaalbedrag = numberCheckOn2Decimal(totaalbedrag,rowCount)
            newFileContent = newFileContent.replace("totaalbedrag",str(totaalbedrag))

            '''
            Create temporary .tex file to store the edited template. The temporary file will be made to an .pdf and will be deleted afterwards.
            '''
            newFile = open(destinationFolder+"/newFile.tex","a+")
            newFile.write(newFileContent)
            newFile.close()
            create_pdf(destinationFolder+"/newFile.tex","Testen/"+onderwerp+" "+voornaam+" "+achternaam)
            os.remove(destinationFolder+"/newFile.tex")
            rowCount = rowCount+1
logFile.close()
os.startfile(destinationFolder+"/logFile.txt")

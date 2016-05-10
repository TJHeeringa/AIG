import subprocess
import csv
import os

from Tkinter import Tk
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory

def create_pdf(input_filename, output_filename):
    process = subprocess.Popen([
        'latex',   # Or maybe 'C:\\Program Files\\MikTex\\miktex\\bin\\latex.exe
        '-output-format=pdf',
        '-job-name=' + output_filename,
        input_filename])
    process.wait()

#def generateTableContent(bedragen)

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
template = open(filename)
#template = open("standaard_factuur.tex")
templateContent = template.read()
csvFile = askopenfilename()
destinationFolder = askdirectory()
#with open("test.csv", 'rb') as csvfile:
with open(csvFile, 'rb') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=';', quotechar='"')
    rowCount = 0
    for row in spamreader:
        if rowCount == 0:
            adresIndex = row.index("Adres")
            plaatsIndex = row.index("Stad")
            postcodeIndex = row.index("Postcode")
            factuurnummerIndex = row.index("Factuurnummer")
            onderwerpIndex = row.index("Onderwerp")
            zaakIndex = row.index("Zaak")
            voornaamIndex = row.index("Voornaam")
            tussenvoegselIndex =  row.index("Tussenvoegsel")
            achternaamIndex = row.index("Achternaam")
            initialenIndex = row.index("Initialen")
            bedragIndex = row.index("Bedrag")
            ibanIndex = row.index("IBAN")
            rowCount = rowCount+1
        else:
            adres = row[adresIndex]
            plaats = row[plaatsIndex]
            postcode = row[postcodeIndex]
            factuurnummer = row[factuurnummerIndex]
            onderwerp = row[onderwerpIndex]
            zaak = row[zaakIndex]
            voornaam = row[voornaamIndex]
            tussenvoegsel = row[tussenvoegselIndex]
            achternaam = row[achternaamIndex]
            initialen = row[initialenIndex]
            iban = row[ibanIndex]
            bedrag = "{0:.2f}".format(float(row[bedragIndex]))
            totaalbedrag = bedrag
    #        tableContent = generateTableContent(bedragen)
            newFileContent = templateContent.replace("adres}{adres","adres}{"+initialen+" "+achternaam+r'\\'+adres+r'\\'+postcode+" "+plaats)
            newFileContent = newFileContent.replace("factuurnummer}{factuurnummer","factuurnummer}{"+factuurnummer)
            newFile2 = open(destinationFolder+"/newFile2.txt","a+")
            newFile2.write(factuurnummer+"/n"+"\n")
            newFile2.close()
            newFileContent = newFileContent.replace("betreftregel}{onderwerp","betreftregel}{"+onderwerp)
            newFileContent = newFileContent.replace("voornaam}{voornaam","voornaam}{"+voornaam)
            newFileContent = newFileContent.replace("zaak}{zaak","zaak}{"+zaak)
            newFileContent = newFileContent.replace("IBAN}{IBAN","IBAN}{"+iban)
            newFileContent = newFileContent.replace("rekeninghouder}{rekeninghouder","rekeninghouder}{"+initialen+" "+achternaam)
            newFileContent = newFileContent.replace("bedrag",bedrag,1)
            newFileContent = newFileContent.replace("totaalbedrag",totaalbedrag)
            newFile = open(destinationFolder+"/newFile.tex","a+")
            newFile.write(newFileContent)
            newFile.close()
            create_pdf(destinationFolder+"/newFile.tex",destinationFolder+"/"+onderwerp+" "+voornaam+" "+achternaam)
            os.remove(destinationFolder+"/newFile.tex")

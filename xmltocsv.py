import xml.etree.ElementTree as et
import csv
import sys
from datetime import datetime

class XMLtoCSV(object):
    
    def __init__(self):
        print("Finvoice XML -> ProCountor CSV-muunnin, Miika Ruissalo")
        inputfile = input("Anna muunnettavan XML-tiedoston nimi: ")

        try:
            self.inputfile = inputfile
            self.tree = et.ElementTree(file=self.inputfile)
            self.root = self.tree.getroot()
            self.convert(inputfile)
        except IOError:
            print("Tiedostoa ei löytynyt. Tarkista, että tiedosto on samassa kansiossa ja että kirjoitit tiedoston nimen tiedostopääte mukaanlukien (.xml)!")
            sys.exit()
        except et.ParseError:
            print("Tiedosto ei ole kelvollista XML:ää!")
            sys.exit()
        
             
    
    def convert(self, inputfile):
        try:
            with open("{0}.csv".format(self.inputfile[:-4]), 'w', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(self.makeInfoRow())
                for invoiceRow in self.root.findall('InvoiceRow'):
                    writer.writerow(self.invoiceRowIterator(invoiceRow))   
                     
            print("---\nMuunto valmis, tallennettu nimellä {0}.csv.".format(inputfile[:-4]))

        except AttributeError:
            print("Virhe muunnossa, XML-tietueen laskuelementtejä puuttuu?")
            sys.exit()
        except Exception as ex:
            print("Virhe muunnossa:", ex)
            sys.exit()

    
    def makeAddress(self, name, address):
        row = name.text
        for i in range(len(address)):
            row += '\\' + address[i].text
        
        return row


    def makeInfoRow(self):
        buyerDet = self.root.find('BuyerPartyDetails')
        dlvrDet = self.root.find('DeliveryPartyDetails')
        invDet = self.root.find('InvoiceDetails')
            
        row = [invDet.find('InvoiceTypeCode').text]
        row += [invDet.find('InvoiceTotalVatIncludedAmount').get('AmountCurrencyIdentifier')]
        row += [''] * 2     # viite ja pankkitili
        row += [buyerDet.find('BuyerPartyIdentifier').text]
        row += ['']         # maksutapa
        row += [buyerDet.find('BuyerOrganisationName').text]
        row += [''] * 4     # toimitustapa, alennus, sis.alv, hyvitys
        row += [invDet.find('PaymentTermsDetails').find('PaymentOverDueFineDetails').find('PaymentOverDueFinePercent').text]
        row += [self.formatTime(invDet.find('InvoiceDate').text)]
        row += [''] * 3     # toimituspvm, erapvm, liik.kump.os.
        row += [self.makeAddress(buyerDet.find('BuyerOrganisationName'), buyerDet.find('BuyerPostalAddressDetails'))]
        row += [self.makeAddress(dlvrDet.find('DeliveryOrganisationName'), dlvrDet.find('DeliveryPostalAddressDetails'))]
        row += [invDet.find('InvoiceFreeText').text]
        row += [''] * 7     # muistiinpanot, sposti, maksupvm, kurssi, loppusumma, alvpros, laskukanava

        return row
        

    def invoiceRowIterator(self, invoiceRow):
        row = ['']
        row += [invoiceRow.find('ArticleName').text]
        row += [invoiceRow.find('ArticleIdentifier').text]
        row += [invoiceRow.find('OrderedQuantity').text]
        row += [invoiceRow.find('UnitPriceAmount').get('UnitPriceUnitCode')]
        row += [invoiceRow.find('UnitPriceAmount').text]
        row += ['']       # alennus
        row += [invoiceRow.find('RowVatRatePercent').text]
        row += [''] * 18  # loput
        
        return row

    
    def formatTime(self, time):
        return datetime.strptime(time, '%Y%m%d').strftime('%d.%m.%Y')
    
    
if __name__ == "__main__":
    XMLtoCSV()

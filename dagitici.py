import datetime
import random
import locale
locale.setlocale(locale.LC_ALL, '')
import math
# Define the start and end dates
baslangictarihi = datetime.datetime(2023, 7, 17)
bitistarihi = datetime.datetime(2023, 8, 4)
toplamgunsayisi=(bitistarihi-baslangictarihi).days
# Define the days of the week
#define a dict for every day between start and end dates and also name of weekday in turkish
gunler={}
while baslangictarihi <= bitistarihi:
    gunler[baslangictarihi.strftime("%d-%m-%Y")] = baslangictarihi.strftime("%A")
    baslangictarihi += datetime.timedelta(days=1)

# Define the list of people and the days they cannot work
kisilergunler = {
    "Ali Karakaya" : "Pazartesi",
    "Veli Eşme" : "Salı",
    "Ayşe Yılmaz" : "Çarşamba",
    "Fatma Demir" : "Perşembe",
    "Hayriye Kaya" : "Cuma",
    "Hüseyin Şahin" : "Cumartesi",
    "Mehmet Yıldız" : "Pazar",
    "Mustafa Demir" : "Pazartesi",
    "Necati Yılmaz" : "Salı",
    "Nuray Kaya" : "Çarşamba",
    "Ömer Şahin" : "Perşembe",
    "Süleyman Yıldız" : "Cuma",
    "Şerife Demir" : "Cumartesi",
    "Zeynep Yılmaz" : "Pazar",
    "Zeki Şahin" : "Pazartesi",
    "Zeynep Yıldız" : "Salı",
    "Zeynep Şahin" : "Çarşamba",
    
}
kisilergunlerilk=kisilergunler.copy()
yeniliste={}

eslestirilmisliste={}
for date,day in gunler.items():
    #get someone from list randomly
    kisi1=random.choice(list(kisilergunler.keys()))
    #if the person is not available on that day, get another person
    while kisilergunler[kisi1]==day:
        kisi1=random.choice(list(kisilergunler.keys()))
    #remove this person from list
    yeniliste[kisi1]=kisilergunler[kisi1]
    kisilergunler.pop(kisi1)
    #get another person from list randomly
    kisi2=random.choice(list(kisilergunler.keys()))
    #if the person is not available on that day, get another person
    while kisilergunler[kisi2]==day:
        kisi2=random.choice(list(kisilergunler.keys()))
    #remove this person from list
    yeniliste[kisi2]=kisilergunler[kisi2]
    kisilergunler.pop(kisi2)
    #add the pair to the list
    eslestirilmisliste[date + " " +day]=[kisi1,kisi2]
    #if there is only one person left, add him/her to the list
    if len(kisilergunler)==1:
        kisison=list(kisilergunler.keys())[0]
        yeniliste[kisison]=kisilergunler[kisison]        
        #make the yeniliste the kisilergunler
        kisilergunler=yeniliste
        yeniliste={}
#print eslestirilmis liste pretty
for date,day in gunler.items():
    print(date,day,eslestirilmisliste[date + " " +day])
#print how many days each person works in eslestirilmisliste
print("\n")
for kisi in kisilergunlerilk:
    #create a new list from eslestirilmisliste.values() list items
    kisiler=[item for sublist in list(eslestirilmisliste.values()) for item in sublist]
    #count how many times the person is in the list
    kisisayisi=kisiler.count(kisi)
    #print the person and the number of days he/she works
    print(kisi,kisisayisi)

#write the dates and people into an excel sheet
import xlsxwriter
workbook = xlsxwriter.Workbook('dagitici.xlsx')
worksheet = workbook.add_worksheet()

#headers bold
bold = workbook.add_format({'bold': True})
worksheet.write('A1', "Tarih", bold)
worksheet.write('B1', "Gün", bold)
worksheet.write('C1', "Kişi 1", bold)
worksheet.write('D1', "Kişi 2", bold)

#make the columns larger
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20)

row = 1
col = 0
for date,day in gunler.items():
    worksheet.write(row, col, date)
    worksheet.write(row, col + 1, day)
    worksheet.write(row, col + 2, eslestirilmisliste[date + " " +day][0])
    worksheet.write(row, col + 3, eslestirilmisliste[date + " " +day][1])
    row += 1
workbook.close()

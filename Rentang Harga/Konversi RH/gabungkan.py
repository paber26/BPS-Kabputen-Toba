from openpyxl import load_workbook

fileKosong = open("D:\BPS-Kabputen-Toba\Rentang Harga\Konversi RH\kosong.txt", "r").readlines()
fileBerisi = open("D:\BPS-Kabputen-Toba\Rentang Harga\Konversi RH\isi.txt", "r").readlines()


wb = load_workbook(filename="D:\BPS-Kabputen-Toba\Rentang Harga\Konversi RH\Template.xlsx")
sheetRange = wb['Sheet1']

prov = str(sheetRange['B1'].value)
kab = str(sheetRange['B2'].value)
kec = str(sheetRange['B3'].value)
desa = str(sheetRange['B4'].value)
klas = str(sheetRange['B5'].value)
nks = str(sheetRange['B6'].value)
pengentri = str(sheetRange['B7'].value)

kode = prov + kab + kec + desa + klas + nks

fileHasil = open("hasil.txt","w")
fileHasil.write(kode + '2' + pengentri + '\n')


start = 11
finish = 217

barisfile = 1
panjang = 0
for x in range(start, finish+1):
    
    temp = len(fileKosong[barisfile])
    if(panjang < temp):
        panjang = temp
#     nomor = sheetRange['A' + str(x)].value

#     min = str(sheetRange['D' + str(x)].value)
#     maks = str(sheetRange['E' + str(x)].value)


#     if(min == "None" and maks == "None"):
#         fileHasil.write(kode + '1' + fileKosong[barisfile][18:])
#     else:
#         cmin = str(' ' * (9 - len(min)) + min)
#         cmaks = str(' ' * (9 - len(maks)) + maks)
#         komoditas = fileBerisi[barisfile][18:]
#         fileHasil.write(kode + '1' + komoditas[:-19] + cmin + cmaks + '\n')

    barisfile += 1
    print(temp)
        
print('-----')
print(panjang)

# fileHasil.close()
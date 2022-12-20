import os, sys
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from docx2pdf import convert


wb = load_workbook(filename="E:\BPS TOBA\Regsosek\Data Toba.xlsx")
sheetRange = wb['Sheet1']

os.chdir(sys.path[0])

doc = DocxTemplate('LABEL.docx')

i = 2
while i <= 904:
    nmkec1 = sheetRange['I'+str(i)].value
    nmdesa1 = sheetRange['K'+str(i)].value
    nmsls1 = sheetRange['O'+str(i)].value
    kdkec1 = sheetRange['H'+str(i)].value
    kddesa1 = sheetRange['J'+str(i)].value
    kdsls1 = sheetRange['M'+str(i)].value
    kdsubsls1 = sheetRange['N'+str(i)].value

    nmkec2 = sheetRange['I'+str(i+1)].value
    nmdesa2 = sheetRange['K'+str(i+1)].value
    nmsls2 = sheetRange['O'+str(i+1)].value
    kdkec2 = sheetRange['H'+str(i+1)].value
    kddesa2 = sheetRange['J'+str(i+1)].value
    kdsls2 = sheetRange['M'+str(i+1)].value
    kdsubsls2 = sheetRange['N'+str(i+1)].value


    nmkec3 = sheetRange['I'+str(i+2)].value
    nmdesa3 = sheetRange['K'+str(i+2)].value
    nmsls3 = sheetRange['O'+str(i+2)].value
    kdkec3 = sheetRange['H'+str(i+2)].value
    kddesa3 = sheetRange['J'+str(i+2)].value
    kdsls3 = sheetRange['M'+str(i+2)].value
    kdsubsls3 = sheetRange['N'+str(i+2)].value

    context = {'nmkec1': nmkec1, 'nmdesa1': nmdesa1, 'nmsls1': nmsls1, 'kdkec1': kdkec1, 'kddesa1': kddesa1, 'kdsls1': kdsls1, 'kdsubsls1': kdsubsls1,
    'nmkec2': nmkec2, 'nmdesa2': nmdesa2, 'nmsls2': nmsls2, 'kdkec2': kdkec2, 'kddesa2': kddesa2, 'kdsls2': kdsls2, 'kdsubsls2': kdsubsls2,
    'nmkec3': nmkec3, 'nmdesa3': nmdesa3, 'nmsls3': nmsls3, 'kdkec3': kdkec3, 'kddesa3': kddesa3, 'kdsls3': kdsls3, 'kdsubsls3': kdsubsls3}
    
    doc.render(context)
    
    doc.save('Label_rendered.docx')
    output = 'E:\BPS TOBA\Regsosek\Hasil PDF\Hasil ' + str(i) + '-' + str(i+2) + '.pdf'
    convert("E:\BPS TOBA\Regsosek\Label_rendered.docx", output)

    i += 3
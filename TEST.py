import openpyxl
from openpyxl import workbook

data_report = openpyxl.load_workbook('baocao1.xlsx')
sheet_name = data_report.sheetnames[0]
sheet1 = data_report[sheet_name]

rows = sheet1.max_row #29
cols = sheet1.max_column #69
vi_tri = str(sheet1[7][cols-23])[15:17]
print(vi_tri)
for cot in range(0, rows-7):
    x = cot + 8
    sheet1[x][cols-22].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"a")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,a5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,a5")/2'
    sheet1[x][cols-21].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"d")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,d5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,d5")/2'
    sheet1[x][cols-20].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"b")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,b5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,b5")/2'
    sheet1[x][cols-19].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"c")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"r0,c5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,c5")/2'
    sheet1[x][cols-18].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"KH")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"TL")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"L")'
    sheet1[x][cols-17].value = ''
    sheet1[x][cols-16].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"hh")'
    sheet1[x][cols-15].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"nt7")'
    sheet1[x][cols-14].value = '=COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"P")+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,a5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,b5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,c5")/2+COUNTIF($G'+str(x)+':'+ vi_tri+str(x)+',"p5,d5")/2'
    sheet1[x][cols-13].value = '=SUBTOTAL(9,BQ7:BY7)'
    sheet1[x][cols-12].value = '=SUM(J7,L7,N7,P7,R7,T7,X7,Z7,AB7,AD7,AF7,AH7,AL7,AN7,AP7,AR7,AT7,AV7,AZ7,BB7,BD7,BF7,BH7,BP7)'
    sheet1[x][cols-11].value = '=SUM(H7,V7,AJ7,AX7,BL7)'
    sheet1[x][cols-10].value = '=SUM(BJ7,BN7)'
    sheet1[x][cols-9].value = '=IF(BZ7>25,4,IF(BZ7>18,3,IF(BZ7>12,2,IF(BZ7>4,1,0))))-BX7-2+1'
    sheet1[x][cols-8].value = ''
    sheet1[x][cols-7].value = '=CD7*CI7'
    sheet1[x][cols-6].value = '=F7-BY7+1'
    sheet1[x][cols-5].value = '=CE7*16'
    sheet1[x][cols-4].value = ''
    sheet1[x][cols-3].value = ''
    sheet1[x][cols-2].value = ''
    sheet1[x][cols-1].value = ''

data_report.save('baocao1.xlsx')
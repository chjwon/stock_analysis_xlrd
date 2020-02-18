import xlrd


#file location
loc = ("naver_stock.xlsx")


#open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)



#sheet.cell_value(0,0) #call (0,0)value
#print(sheet.cell_value(0,0))


#size of excel
#print(sheet.nrows) #4062
#print(sheet.ncols) #2


#columns
#for i in range(sheet.ncols):
#    print(sheet.cell_value(0,i))

global year_num
year_num = 0
global year_start
year_start = 2002
global total
total = 0
global row_num
row_num = 0
global avg
avg = 0
global date_num
date_num = 0
global month_num
month_num = 0
global quarter
quarter = 0


#print by year
for i in range(sheet.nrows-1):
    #print(int(sheet.cell_value(i+1,0)))
    year_num = int(sheet.cell_value(i+1,0) / 10000)
    #print(year_num)
    if(year_num == year_start):
        total = total +sheet.cell_value(i+1,1)
        row_num = row_num +1
    else:
        avg = float(total/row_num )
        avg = round(avg,5)
        print(str(year_start) + "년도 평균 주가 : "+str(avg))
        year_start = year_start + 1
        avg = 0
        total = 0
        row_num = 0
        

print("===========================")

global total_1
total_1 = 0
global total_2
total_2 = 0
global total_3
total_3 = 0
global total_4
total_4 = 0
global row_num_1
row_num_1 = 1
global row_num_2
row_num_2 = 1
global row_num_3
row_num_3 = 1
global row_num_4
row_num_4 = 1


year_start = 2002



#print by quarter
for i in range(sheet.nrows-1):
    year_num = int(sheet.cell_value(i+1,0)/10000)
    date_num = int(sheet.cell_value(i+1,0)%10000)
    month_num = int(date_num/100)
    quarter = int((month_num - 1)/3)+1
    #print(quarter)
    if(year_num == year_start):
        if(quarter == 1):
            total_1 = total_1 + sheet.cell_value(i+1,1)
            row_num_1 = row_num_1 + 1
        if(quarter == 2):
            total_2 = total_2 + sheet.cell_value(i+1,1)
            row_num_2 = row_num_2 + 1
        if(quarter == 3):
            total_3 = total_3 + sheet.cell_value(i+1,1)
            row_num_3 = row_num_3 + 1
        if(quarter == 4):
            total_4 = total_4 + sheet.cell_value(i+1,1)
            row_num_4 = row_num_4 + 1
    else:
        avg = float(total_1 / row_num_1 )
        avg = round(avg,5)
        print(str(year_start) + "년도 1분기 평균 주가 : "+str(avg))
        avg = float(total_2 / row_num_2 )
        avg = round(avg,5)
        print(str(year_start) + "년도 2분기 평균 주가 : "+str(avg))
        avg = float(total_3 / row_num_3 )
        avg = round(avg,5)
        print(str(year_start) + "년도 3분기 평균 주가 : "+str(avg))
        avg = float(total_4 / row_num_4 )
        avg = round(avg,5)
        print(str(year_start) + "년도 4분기 평균 주가 : "+str(avg))
        print('\n')
        year_start = year_start + 1
        avg = 0
        total = 0
        row_num = 0 
      


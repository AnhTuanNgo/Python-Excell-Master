# -*- coding: utf-8 -*-
"""
Created on Sat Aug 20 14:14:45 2022

@author: Admin
"""

import xlwings as xw
import pandas as pd
import numpy as np




#Viết hàm mở file excel 

def readExcel(path_input): 
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    wb = app.books.open(path_input) 
    return wb
# Hàm kiểm tra last_row, last_columns
           

def main(): 
    
    print('**CHÀO MỪNG ĐẾN VỚI PROJECT AUTOMATION EXCEL**')
    # Mở file Excell
    wb=readExcel(r'C:\Users\Admin\Dropbox\PC\Desktop\Python for Excel\Example\Add _data\Test.xlsx')
    # Đến sheet cần làm việc --> đặt tên biên sheet
    sh=wb.sheets['data']
    # Nhập xác nhận chạy trương trình
    polling_active = True
    while polling_active:
        keys = input("\nBạn chắc chắn muốn chạy chương trình không (yes/press any in keybord)? ")
        if keys == 'yes':
            print('Chương trình tiếp tục chạy:')
            polling_active = False
        else:  
            return  wb.app.quit()
            
          

    
       
        
         
            

    # THÊM DỮ LIỆU
    # Phương thước sheets.range('ô excel').options(transpose = True).value--> điền theo hàng dọc
    
    lr_b=sh.range("B1").end('down').row
    print('Lr_b:',lr_b)
    data=[['Speedy Express','(503) 555-9833',3,100000],
         ['Speedy Express','(503) 555-9833',2,100000]]
    dt=pd.DataFrame(data=data)   
                
                      
    # Điền giá trị  trong list vào file excell
    # Loại bỏ Index, header của dataframe bằng thuộc tính : options(index=False, header=0)
    sh.range(f'B{lr_b+1}').options(index=False,header=0).value= dt   
    
   #TÍNH TOÁN ADD CÔNG THỨC VÀO CỘT F
    # Tìm dòng cuối từ dưới lên trên ở cột D HOẶC E, C
    lr_data=sh.range('B'+str(sh.cells.last_cell.row)).end('up').row
    print('Lr_data:',lr_data)
    # Tạo list chứa dữ liệu công thức cần điền
    c_i=[]
    # Thêm các giá trị cần điền vào list mới tạo
    for i in range(2,lr_data+1):
     ct_total= f'=D{i}*E{2}'
     c_i.append(ct_total)
    # Điền giá trị  trong list vào file excell     
          
    sh.range('F2').options(transpose =True).value= c_i
    
    # ĐÁNH SÔ THỨ TỰ Ở CỘT A
    #Tìm dòng cuối từ trên xuống dưới ở dòng A (ghi thêm dữ liệu )
    lr_a=sh.range("A1").end('down').row
    
    print('Lr_a:',lr_a)
    
    stt=[]
    
    for i in range(lr_a,lr_data):
     stt.append(i)
        
    sh.range(f'A{lr_a+1}').options(transpose =True).value= stt
    #Lưu file
    wb.save(r'C:\Users\Admin\Dropbox\PC\Desktop\Python for Excel\Example\Add _data\Test.xlsx')
    #Đóng workbooK
    #wb.close()
    #Đóng ứng dụng
    #wb.app.quit()
          
    
    
    
    
    
if __name__=='__main__':
    main() 
         



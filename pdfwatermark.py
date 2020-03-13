######## 基础准备 ########
import os
os.getcwd() 
os.chdir('D:\\sourcecode\\python\\test')  ##设置需要进行读取的操作目录
os.getcwd()    #获取当前工作目录

sourcefile = 'SBS1单词汇总.pdf'           ##需要加上水印的原文件  
filename = '分析报告'                     ##生成后命名的文件内容
#password = 'bestpay'       ## PDF的密码,设置为空则不设密码访问
password = ''
listfile = 'watermark.xlsx'    ##批量名单获取的文件，清单在Sheet1，从第2行开始的第2列内容

from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 
pdfmetrics.registerFont(TTFont('song', 'C:/Windows/Fonts/simsun.ttc'))#宋体
from PyPDF2 import PdfFileWriter,PdfFileReader
import xlrd
  
######## 1.生成水印pdf的函数 ########
def create_watermark(content):
    #默认大小为21cm*29.7cm
    c = canvas.Canvas('mark.pdf', pagesize = (30*cm, 30*cm))   
    c.translate(10*cm, 10*cm) #移动坐标原点(坐标系左下为(0,0)))                                                                                                                             
    c.setFont('song',22)#设置字体为宋体，大小22号
    c.setFillColorRGB(0.5,0.5,0.5)#灰色                                                                                                                         
    c.rotate(45)#旋转45度，坐标系被旋转
    c.drawString(-7*cm, 0*cm, content)
    c.drawString(7*cm, 0*cm, content)
    c.drawString(0*cm, 7*cm, content)
    c.drawString(0*cm, -7*cm, content)                                                                                                                              
    c.save()#关闭并保存pdf文件

######## 2.为pdf文件加水印的函数 ########
def add_watermark2pdf(input_pdf,output_pdf,watermark_pdf):
    watermark = PdfFileReader(watermark_pdf)
    watermark_page = watermark.getPage(0)
    pdf = PdfFileReader(input_pdf,strict=False)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(watermark_page)
        pdf_writer.addPage(pdf_page)
    pdfOutputFile = open(output_pdf,'wb')   
    if password !='' :
        pdf_writer.encrypt(password)                       #设置pdf密码
    pdf_writer.write(pdfOutputFile)
    pdfOutputFile.close()

######## 3.导入excel：excel ########
ExcelFile = xlrd.open_workbook(listfile)
sheet=ExcelFile.sheet_by_name('Sheet1')                 #打开有名单那个sheet
print('———————已导入名单———————')
col = sheet.col_values(1)                               #第2列内容为名称
id = sheet.col_values(0)                                #第1列内容为ID
del col[0];del id[0]                                    #去掉第1行标题
id2 = [str(int(i)) for i in id]
merchant_as_mark_content =[(i+'  ')*4 if len(i)<=5 else i for i in col]#如果名称太短则重复4个为一行

######## 4.调用前面的函数制作商家水印pdf ########
if __name__=='__main__':
    for i,j,k in zip(merchant_as_mark_content,col,id2):#i制作水印，j文件名，k对应ID
        create_watermark(i)#创造了一个水印pdf：mark.pdf
        add_watermark2pdf(sourcefile ,k+filename+'('+j+').pdf','mark.pdf')              ##需要加水印
        print('———————已制作好第'+k+'个pdf，正在准备下一个———————')
    print('———————所有文件已转化完毕———————')

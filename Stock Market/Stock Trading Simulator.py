import time,xlrd,matplotlib
from tkinter import *
from PIL import ImageTk,Image 
from matplotlib.pyplot import *
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter.ttk import Combobox
from matplotlib import rcParams
from datetime import datetime

root=Tk()
root.title('Stock Trading Simulator')
root.geometry('1920x1080')
frame1=Frame(root,width=1087,height=480,cursor='circle')
frame1.place(x=0,y=0)
canvas1=Canvas(frame1,width=1093,height=480,scrollregion=(0,0,1093,1325))
stock_ref_tab=ImageTk.PhotoImage(Image.open('Stock Refresh Tab.png'))
canvas1.create_image(546.5,612.5,image=stock_ref_tab)
scroll1=Scrollbar(frame1,orient=VERTICAL,width=10,cursor='dot')
scroll1.pack(side=RIGHT,fill=Y)
scroll1.configure(command=canvas1.yview)
canvas1.configure(yscrollcommand=scroll1.set)
canvas1.pack()
canvas4=Canvas(root,width=608,height=600)
cart_tab=ImageTk.PhotoImage(Image.open('Cart.png'))
canvas4.create_image(306,180,image=cart_tab)
canvas4.place(x=485,y=480)
frame2=Frame(root,width=608,height=339,cursor='circle',bg='light gray')
frame2.place(x=485,y=505)
frame3=Frame(root,width=485,height=560,cursor='circle')
frame3.place(x=0,y=520)
frame4=Frame(root,width=810,height=1080,cursor='circle')
frame4.place(x=1110,y=0)
frame3a=Frame(root,width=485,height=40,cursor='circle',bg='Black')
frame3a.place(x=0,y=480)
canvas2=Canvas(frame4,width=440,height=842,bg='black')
trans_tab=ImageTk.PhotoImage(Image.open('Transaction.png'))
canvas2.create_image(221,424,image=trans_tab)
canvas2.pack()
rcParams.update({'figure.autolayout':True})

canvas1.create_text(30,37.5,text='S.No.',fill='orange',font='Calibri 14')
canvas1.create_text(86,37.5,text='Code',fill='orange',font='Calibri 14')
canvas1.create_text(155,37.5,text='Exchange',fil='orange',font='Calibri 14')
canvas1.create_text(349,37.5,text='Name',fil='orange',font='Calibri 14')
canvas1.create_text(555,37.5,text='High',fil='orange',font='Calibri 14')
canvas1.create_text(620,37.5,text='Low',fil='orange',font='Calibri 14')
canvas1.create_text(683,37.5,text='Price',fil='orange',font='Calibri 14')
canvas1.create_text(746,37.5,text='Close',fil='orange',font='Calibri 14')
canvas1.create_text(832.5,37.5,text='Volume',fil='orange',font='Calibri 14')
canvas1.create_text(927.5,37.5,text='Adj Close',fil='orange',font='Calibri 14')
canvas1.create_text(1027,37.5,text='Change',fill='orange',font='Calibri 14')
canvas4.create_text(15,12.5,text='No.',fill='orange',font='Calibri 14')
canvas4.create_text(85,12.5,text='Date',fill='orange',font='Calibri 14')
canvas4.create_text(175,12.5,text='Exchange',fill='orange',font='Calibri 14')
canvas4.create_text(240,12.5,text='Code',fill='orange',font='Calibri 14')
canvas4.create_text(360,12.5,text='Name',fill='orange',font='Calibri 14')
canvas4.create_text(510,12.5,text='Price',fill='orange',font='Calibri 14')
canvas4.create_text(560,12.5,text='Qu.',fill='orange',font='Calibri 14')
canvas4.create_text(589,12.5,text='State',fill='orange',font='Calibri 14')

workbook=xlrd.open_workbook('Stocks.xlsx')
worksheet=workbook.sheet_by_index(0)

def stock(Number,Code,Exchange,Name,number,code,exchange,name,open,high,low,close,volume,adj,vchange,row,y):
    change=worksheet.cell(row,3).value-worksheet.cell(row-1,3).value
    canvas1.delete(Number,Code,Exchange,Name,open,high,low,close,volume,adj,vchange)
    Number=canvas1.create_text(30,y,text=number,fill='white',font='Calibri 14')
    Code=canvas1.create_text(86,y,text=code,fill='white',font='Calibri 14')
    Exchange=canvas1.create_text(155,y,text=exchange,fill='white',font='Calibri 14')
    Name=canvas1.create_text(349,y,text=name,fill='white',font='Calibri 14')
    if worksheet.cell(row,1).value>worksheet.cell((row-1),1).value:
        high=canvas1.create_text(555,y,text=worksheet.cell(row,1).value,fill='light green',font='Calibri 14')
    else:
        high=canvas1.create_text(555,y,text=worksheet.cell(row,1).value,fill='red',font='Calibri 14')
    if worksheet.cell(row,2).value>worksheet.cell(row-1,2).value:
        low=canvas1.create_text(620,y,text=worksheet.cell(row,2).value,fill='light green',font='Calibri 14')
    else:
        low=canvas1.create_text(620,y,text=worksheet.cell(row,2).value,fill='red',font='Calibri 14')
    if worksheet.cell(row,3).value>worksheet.cell(row-1,3).value:
        open=canvas1.create_text(683,y,text=worksheet.cell(row,3).value,fill='light green',font='Calibri 14')
    else:
        open=canvas1.create_text(683,y,text=worksheet.cell(row,3).value,fill='red',font='Calibri 14')   
    if worksheet.cell(row,4).value>worksheet.cell(row-1,4).value:
        close=canvas1.create_text(746,y,text=worksheet.cell(row,4).value,fill='light green',font='Calibri 14')
    else:
        close=canvas1.create_text(746,y,text=worksheet.cell(row,4).value,fill='red',font='Calibri 14')
    if worksheet.cell(row,5).value>worksheet.cell(row-1,5).value:
        volume=canvas1.create_text(832.5,y,text=worksheet.cell(row,5).value,fill='light green',font='Calibri 14')
    else:
        volume=canvas1.create_text(832.5,y,text=worksheet.cell(row,5).value,fill='red',font='Calibri 14')
    if worksheet.cell(row,6).value>worksheet.cell(row-1,6).value:
        adj=canvas1.create_text(927.5,y,text=worksheet.cell(row,6).value,fill='light green',font='Calibri 14')
    else:
        adj=canvas1.create_text(927.5,y,text=worksheet.cell(row,6).value,fill='red',font='Calibri 14')
    if change>0:
        vchange=canvas1.create_text(1027,y,text=round(change,2),fill='light green',font='Calibri 14')
    else:
        vchange=canvas1.create_text(1027,y,text=round(change,2),fill='red',font='Calibri 14')
    row+=1
    root.after(4000,stock,Number,Code,Exchange,Name,number,code,exchange,name,open,high,low,close,volume,adj,vchange,row,y)

displaytime=canvas2.create_text(20,20,text=worksheet.cell(1,6).value,fill='white',font='Calibri 14')
z=1
def time():
    global displaytime
    global z
    canvas2.delete(displaytime)
    displaytime=canvas2.create_text(320,20,text=worksheet.cell(int(z),0).value,fill='white',font='Calibri 14')
    root.after(4000,time)
    z+=1
    return z
time()

stock('num1', 'code1', 'exchange1', 'name1', '1.', 'AAPL', 'NASDAQ', 'Apple Inc.', 'open1', 'high1', 'low1', 'close1', 'volume1', 'adj1', 'vchange1', 2, 62.5)
stock('num2', 'code2', 'exchange2', 'name2', '2.', 'AMT', 'NYSE', 'American Tower', 'open2', 'high2', 'low2', 'close2', 'volume2', 'adj2', 'vchange2', 302, 87.5)
stock('num3', 'code3', 'exchange3', 'name3', '3.', 'AMZN', 'NASDAQ', 'Amazon.com Inc.', 'open3', 'high3', 'low3', 'close3', 'volume3', 'adj3', 'vchange3', 602, 112.5)
stock('num4', 'code4', 'exchange4', 'name4', '4.', 'AMAT', 'NASDAQ', 'Applied Materials', 'open4', 'high4', 'low4', 'close4', 'volume4', 'adj4', 'vchange4', 902, 137.5)
stock('num5', 'code5', 'exchange5', 'name5', '5.', 'AVGO', 'NASDAQ', 'Broadcom', 'open5', 'high5', 'low5', 'close5', 'volume5', 'adj5', 'vchange5', 1202, 162.5)
stock('num6', 'code6', 'exchange6', 'name6', '6.', 'AXP', 'NYSE', 'American Express Co', 'open6', 'high6', 'low6', 'close6', 'volume6', 'adj6', 'vchange6', 1502, 187.5)
stock('num7', 'code7', 'exchange7', 'name7', '7.', 'BA', 'NYSE', 'Boeing Company', 'open7', 'high7', 'low7', 'close7', 'volume7', 'adj7', 'vchange7', 1802, 212.5)
stock('num8', 'code8', 'exchange8', 'name8', '8.', 'CTL', 'NYSE', 'CenturyLink Inc', 'open8', 'high8', 'low8', 'close8', 'volume8', 'adj8', 'vchange8', 2102, 237.5)
stock('num9', 'code9', 'exchange9', 'name9', '9.', 'C', 'NYSE', 'Citigroup', 'open9', 'high9', 'low9', 'close9', 'volume9', 'adj9', 'vchange9', 2402, 262.5)
stock('num10', 'code10', 'exchange10', 'name10', '10.', 'CPRT', 'NASDAQ', 'Copart Inc', 'open10', 'high10', 'low10', 'close10', 'volume10', 'adj10', 'vchange10', 2702, 287.5)
stock('num11', 'code11', 'exchange11', 'name11', '11.', 'DIS', 'NYSE', 'The Walt Disney Company', 'open11', 'high11', 'low11', 'close11', 'volume11', 'adj11', 'vchange11', 3002, 312.5)
stock('num12', 'code12', 'exchange12', 'name12', '12.', 'EQIX', 'NASDAQ', 'Equinix', 'open12', 'high12', 'low12', 'close12', 'volume12', 'adj12', 'vchange12', 3302, 337.5)
stock('num13', 'code13', 'exchange13', 'name13', '13.', 'EFX', 'NYSE', 'Equifax', 'open13', 'high13', 'low13', 'close13', 'volume13', 'adj13', 'vchange13', 3602, 362.5)
stock('num14', 'code14', 'exchange14', 'name14', '14.', 'FB', 'NASDAQ', 'Facebook', 'open14', 'high14', 'low14', 'close14', 'volume14', 'adj14', 'vchange14', 3902, 387.5)
stock('num15', 'code15', 'exchange15', 'name15', '15.', 'FDX', 'NYSE', 'FedEx Corporation', 'open15', 'high15', 'low15', 'close15', 'volume15', 'adj15', 'vchange15', 4202, 412.5)
stock('num16', 'code16', 'exchange16', 'name16', '16.', 'FIS', 'NYSE', 'Fidelity National Information Services', 'open16', 'high16', 'low16', 'close16', 'volume16', 'adj16', 'vchange16', 4502, 437.5)
stock('num17', 'code17', 'exchange17', 'name17', '17.', 'FISV', 'NASDAQ', 'Fiserv Inc', 'open17', 'high17', 'low17', 'close17', 'volume17', 'adj17', 'vchange17', 4802, 462.5)
stock('num18', 'code18', 'exchange18', 'name18', '18.', 'GOOG', 'NASDAQ', 'Alphabet Inc Class C', 'open18', 'high18', 'low18', 'close18', 'volume18', 'adj18', 'vchange18', 5102, 487.5)
stock('num19', 'code19', 'exchange19', 'name19', '19.', 'GS', 'NYSE', 'Goldman Sachs Group', 'open19', 'high19', 'low19', 'close19', 'volume19', 'adj19', 'vchange19', 5402, 512.5)
stock('num20', 'code20', 'exchange20', 'name20', '20.', 'HON', 'NYSE', "Honeywell Int'l", 'open20', 'high20', 'low20', 'close20', 'volume20', 'adj20', 'vchange20', 5702, 537.5)
stock('num21', 'code21', 'exchange21', 'name21', '21.', 'HII', 'NYSE', 'Huntington Ingalls Industries', 'open21', 'high21', 'low21', 'close21', 'volume21', 'adj21', 'vchange21', 6002, 562.5)
stock('num22', 'code22', 'exchange22', 'name22', '22.', 'INTC', 'NYSE', 'Intel Corp.', 'open22', 'high22', 'low22', 'close22', 'volume22', 'adj22', 'vchange22', 6302, 587.5)
stock('num23', 'code23', 'exchange23', 'name23', '23.', 'IRM', 'NYSE', 'Iron Mountain Incorporated', 'open23', 'high23', 'low23', 'close23', 'volume23', 'adj23', 'vchange23', 6602, 612.5)
stock('num24', 'code24', 'exchange24', 'name24', '24.', 'ISRG', 'NASDAQ', 'Intuitive Surgical Inc.', 'open24', 'high24', 'low24', 'close24', 'volume24', 'adj24', 'vchange24', 6902, 637.5)
stock('num25', 'code25', 'exchange25', 'name25', '25.', 'JNJ', 'NYSE', 'Johnson & Johnson', 'open25', 'high25', 'low25', 'close25', 'volume25', 'adj25', 'vchange25', 7202, 662.5)
stock('num26', 'code26', 'exchange26', 'name26', '26.', 'KLAC', 'NASDAQ', 'KLA-Tencor', 'open26', 'high26', 'low26', 'close26', 'volume26', 'adj26', 'vchange26', 7502, 687.5)
stock('num27', 'code27', 'exchange27', 'name27', '27.', 'KHC', 'NASDAQ', 'Kraft Heinz Co', 'open27', 'high27', 'low27', 'close27', 'volume27', 'adj27', 'vchange27', 7802, 712.5)
stock('num28', 'code28', 'exchange28', 'name28', '28.', 'LH', 'NYSE', 'Laboratory of America Holdings', 'open28', 'high28', 'low28', 'close28', 'volume28', 'adj28', 'vchange28', 8102, 737.5)
stock('num29', 'code29', 'exchange29', 'name29', '29.', 'MA', 'NYSE', 'Mastercard Inc.', 'open29', 'high29', 'low29', 'close29', 'volume29', 'adj29', 'vchange29', 8402, 762.5)
stock('num30', 'code30', 'exchange30', 'name30', '30.', 'MCD', 'NYSE', "McDonald's Corp.", 'open30', 'high30', 'low30', 'close30', 'volume30', 'adj30', 'vchange30', 8702, 787.5)
stock('num31', 'code31', 'exchange31', 'name31', '31.', 'MU', 'NASDAQ', 'Micron Technology', 'open31', 'high31', 'low31', 'close31', 'volume31', 'adj31', 'vchange31', 9002, 812.5)
stock('num32', 'code32', 'exchange32', 'name32', '32.', 'MSFT', 'NASDAQ', 'Microsoft', 'open32', 'high32', 'low32', 'close32', 'volume32', 'adj32', 'vchange32', 9302, 837.5)
stock('num33', 'code33', 'exchange33', 'name33', '33.', 'NTAP', 'NASDAQ', 'NetApp', 'open33', 'high33', 'low33', 'close33', 'volume33', 'adj33', 'vchange33', 9602, 862.5)
stock('num34', 'code34', 'exchange34', 'name34', '34.', 'NFLX', 'NASDAQ', 'Netflix Inc.', 'open34', 'high34', 'low34', 'close34', 'volume34', 'adj34', 'vchange34', 9902, 887.5)
stock('num35', 'code35', 'exchange35', 'name35', '35.', 'NKE', 'NYSE', 'Nike', 'open35', 'high35', 'low35', 'close35', 'volume35', 'adj35', 'vchange35', 10202, 912.5)
stock('num36', 'code36', 'exchange36', 'name36', '36.', 'NVDA', 'NASDAQ', 'Nvidia Corporation', 'open36', 'high36', 'low36', 'close36', 'volume36', 'adj36', 'vchange36', 10502, 937.5)
stock('num37', 'code37', 'exchange37', 'name37', '37.', 'ORCL', 'NYSE', 'Oracle Corp.', 'open37', 'high37', 'low37', 'close37', 'volume37', 'adj37', 'vchange37', 10802, 962.5)
stock('num38', 'code38', 'exchange38', 'name38', '38.', 'PAYX', 'NASDAQ', 'Paychex', 'open38', 'high38', 'low38', 'close38', 'volume38', 'adj38', 'vchange38', 11102, 987.5)
stock('num39', 'code39', 'exchange39', 'name39', '39.', 'PYPL', 'NASDAQ', 'PayPal', 'open39', 'high39', 'low39', 'close39', 'volume39', 'adj39', 'vchange39', 11402, 1012.5)
stock('num40', 'code40', 'exchange40', 'name40', '40.', 'SPGI', 'NYSE', 'S&P Global', 'open40', 'high40', 'low40', 'close40', 'volume40', 'adj40', 'vchange40', 11702, 1037.5)
stock('num41', 'code41', 'exchange41', 'name41', '41.', 'SBUX', 'NASDAQ', 'Starbucks Corp.', 'open41', 'high41', 'low41', 'close41', 'volume41', 'adj41', 'vchange41', 12002, 1062.5)
stock('num42', 'code42', 'exchange42', 'name42', '42.', 'SNPS', 'NASDAQ', 'Synopsys', 'open42', 'high42', 'low42', 'close42', 'volume42', 'adj42', 'vchange42', 12302, 1087.5)
stock('num43', 'code43', 'exchange43', 'name43', '43.', 'TSLA', 'NASDAQ', 'Tesla Inc', 'open43', 'high43', 'low43', 'close43', 'volume43', 'adj43', 'vchange43', 12602, 1112.5)
stock('num44', 'code44', 'exchange44', 'name44', '44.', 'UPS', 'NYSE', 'United Parcel Service', 'open44', 'high44', 'low44', 'close44', 'volume44', 'adj44', 'vchange44', 12902, 1137.5)
stock('num45', 'code45', 'exchange45', 'name45', '45.', 'V', 'NYSE', 'Visa Inc.', 'open45', 'high45', 'low45', 'close45', 'volume45', 'adj45', 'vchange45', 13202, 1162.5)
stock('num46', 'code46', 'exchange46', 'name46', '46.', 'VRTX', 'NASDAQ', 'Vertex Pharmaceuticals Inc', 'open46', 'high46', 'low46', 'close46', 'volume46', 'adj46', 'vchange46', 13502, 1187.5)
stock('num47', 'code47', 'exchange47', 'name47', '47.', 'WFC', 'NYSE', 'Wells Fargo', 'open47', 'high47', 'low47', 'close47', 'volume47', 'adj47', 'vchange47', 13802, 1212.5)
stock('num48', 'code48', 'exchange48', 'name48', '48.', 'WDAY', 'NASDAQ', 'Workday', 'open48', 'high48', 'low48', 'close48', 'volume48', 'adj48', 'vchange48', 14102, 1237.5)
stock('num49', 'code49', 'exchange49', 'name49', '49.', 'ZBRA', 'NASDAQ', 'Zebra Technologies', 'open49', 'high49', 'low49', 'close49', 'volume49', 'adj49', 'vchange49', 14402, 1262.5)
stock('num50', 'code50', 'exchange50', 'name50', '50.', 'ZM', 'NASDAQ', 'Zoom', 'open50', 'high50', 'low50', 'close50', 'volume50', 'adj50', 'vchange50', 14702, 1287.5)

num_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=3,cursor='circle',selectbackground='black',activestyle='none')
date_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=10,cursor='circle',selectbackground='black',activestyle='none')
xchng_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=7,cursor='circle',selectbackground='black',activestyle='none')
code_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=5,cursor='circle',selectbackground='black',activestyle='none')
name_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=20,cursor='circle',selectbackground='black',activestyle='none')
price_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=7,cursor='circle',selectbackground='black',activestyle='none')
quantity_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=2,cursor='circle',selectbackground='black',activestyle='none')
status_lb=Listbox(frame2,bg='light gray',fg='black',bd=1,font='Calibri 14',height=14,width=4,cursor='circle',selectbackground='black',activestyle='none')

def scroll2(x,y):
    num_lb.yview(x,y)
    date_lb.yview(x,y)
    xchng_lb.yview(x,y)
    code_lb.yview(x,y)
    name_lb.yview(x,y)
    price_lb.yview(x,y)
    quantity_lb.yview(x,y)
    status_lb.yview(x,y)

n,scrollbar2=1,Scrollbar(frame2,orient=VERTICAL)
def trade_cart(date,exchange,code,name,price,quantity,status):
    global n,scrollbar2
    num_lb.insert(0,str(n)+'.')
    n+=1
    date_lb.insert(0,date)
    xchng_lb.insert(0,exchange)
    code_lb.insert(0,code)
    name_lb.insert(0,name)
    price_lb.insert(0,price)
    quantity_lb.insert(0,quantity)
    status_lb.insert(0,status)
    scrollbar2.destroy()
    scrollbar2=Scrollbar(frame2,orient=VERTICAL,width=10,cursor='dot')
    scrollbar2.pack(side=RIGHT,fill=Y)
    scrollbar2.configure(command=scroll2)
    num_lb.configure(yscrollcommand=scrollbar2.set)
    num_lb.pack(side=LEFT,fill=X,expand=True)
    date_lb.configure(yscrollcommand=scrollbar2.set)
    date_lb.pack(side=LEFT,fill=X,expand=True)
    xchng_lb.configure(yscrollcommand=scrollbar2.set)
    xchng_lb.pack(side=LEFT,fill=X,expand=True)
    code_lb.configure(yscrollcommand=scrollbar2.set)
    code_lb.pack(side=LEFT,fill=X,expand=True)
    name_lb.configure(yscrollcommand=scrollbar2.set)
    name_lb.pack(side=LEFT,fill=X,expand=True)
    price_lb.configure(yscrollcommand=scrollbar2.set)
    price_lb.pack(side=LEFT,fill=X,expand=True)
    quantity_lb.configure(yscrollcommand=scrollbar2.set)
    quantity_lb.pack(side=LEFT,fill=X,expand=True)
    status_lb.configure(yscrollcommand=scrollbar2.set)
    status_lb.pack(side=LEFT,fill=X,expand=True)

fig=Figure(figsize=(4.85,3.26),dpi=100)
plt=fig.add_subplot(111)
plt.set_title(' ',font='Calibri')
plt.set_facecolor('black')
fig.set_facecolor('black')
plt.grid(True,linestyle='-')
plt.tick_params(labelcolor='w',grid_color='gray')
fig.set_edgecolor='w'
subplots_adjust(bottom=0.15)
plt.set_xlabel('Date',color='w')
plt.plot([0],[0],marker='o',color='red',markerfacecolor='white')
for plt in fig.axes:
    plt.tick_params(axis='x',rotation=90)
canvas4=FigureCanvasTkAgg(fig,frame3)
canvas4.draw()
canvas4.get_tk_widget().pack(side=BOTTOM)

def graph(b,c,name):
    x=list()
    y=list()
    def subgraph(b,c,name):
        global canvas4
        if worksheet.cell(b,0).value==c:
            canvas4.get_tk_widget().pack_forget()
            fig=Figure(figsize=(4.85,3.26))
            plt=fig.add_subplot(111)
            plt.set_title(name,font='Calibri')
            plt.set_facecolor('black')
            fig.set_facecolor('black')
            plt.grid(True,linestyle='-')
            plt.tick_params(labelcolor='w',grid_color='gray')
            fig.set_edgecolor='w'
            subplots_adjust(bottom=0.15)
            plt.set_xlabel('Date',color='w')
            plt.plot(x,y,marker='o',color='red',markerfacecolor='white')
            for plt in fig.axes:
                plt.tick_params(axis='x',rotation=90)
            canvas4=FigureCanvasTkAgg(fig,frame3)
            canvas4.draw()
            canvas4.get_tk_widget().pack(side=RIGHT)
        else:
            b+=1
            v=worksheet.cell(b,0).value
            u=worksheet.cell(b,3).value
            x.append(v)
            y.append(u)
            subgraph(b,c,name)
    subgraph(b,c,name)

var=StringVar()
namegraph=Combobox(frame3a,width=20,textvariable=var,state='readonly',font='Calibri 14',value=('Apple Inc.', 'American Tower', 'Amazon.com Inc.', 'Applied Materials', 'Broadcom', 'American Express Co', 'Boeing Company', 'CenturyLink Inc', 'Citigroup', 'Copart Inc', 'The Walt Disney Company', 'Equinix', 'Equifax', 'Facebook', 'FedEx Corporation', 'Fidelity National Information Services', 'Fiserv Inc', 'Alphabet Inc Class C', 'Goldman Sachs Group', "Honeywell Int'l", 'Huntington Ingalls Industries', 'Intel Corp.', 'Iron Mountain Incorporated', 'Intuitive Surgical Inc.', 'Johnson & Johnson', 'KLA-Tencor', 'Kraft Heinz Co', 'Laboratory of America Holdings', 'Mastercard Inc.', "McDonald's Corp.", 'Micron Technology', 'Microsoft', 'NetApp', 'Netflix Inc.', 'Nike', 'Nvidia Corporation', 'Oracle Corp.', 'Paychex', 'PayPal', 'S&P Global', 'Starbucks Corp.', 'Synopsys', 'Tesla Inc', 'United Parcel Service', 'Visa Inc.', 'Vertex Pharmaceuticals Inc', 'Wells Fargo', 'Workday', 'Zebra Technologies', 'Zoom'))
namegraph.place(x=20,y=10)
namegraph.current()
namegraph.master.option_add('*TCombobox*Listbox.Background','gray')

x=0
start_time=datetime.now()
namedict={'Apple Inc.': [2, 'NASDAQ', 'AAPL'], 'American Tower': [302, 'NYSE', 'AMT'], 'Amazon.com Inc.': [602, 'NASDAQ', 'AMZN'], 'Applied Materials': [902, 'NASDAQ', 'AMAT'], 'Broadcom': [1202, 'NASDAQ', 'AVGO'], 'American Express Co': [1502, 'NYSE', 'AXP'], 'Boeing Company': [1802, 'NYSE', 'BA'], 'CenturyLink Inc': [2102, 'NYSE', 'CTL'], 'Citigroup': [2402, 'NYSE', 'C'], 'Copart Inc': [2702, 'NASDAQ', 'CPRT'], 'The Walt Disney Company': [3002, 'NYSE', 'DIS'], 'Equinix': [3302, 'NASDAQ', 'EQIX'], 'Equifax': [3602, 'NYSE', 'EFX'], 'Facebook': [3902, 'NASDAQ', 'FB'], 'FedEx Corporation': [4202, 'NYSE', 'FDX'], 'Fidelity National Information Services': [4502, 'NYSE', 'FIS'], 'Fiserv Inc': [4802, 'NASDAQ', 'FISV'], 'Alphabet Inc Class C': [5102, 'NASDAQ', 'GOOG'], 'Goldman Sachs Group': [5402, 'NYSE', 'GS'], "Honeywell Int'l": [5702, 'NYSE', 'HON'], 'Huntington Ingalls Industries': [6002, 'NYSE', 'HII'], 'Intel Corp.': [6302, 'NYSE', 'INTC'], 'Iron Mountain Incorporated': [6602, 'NYSE', 'IRM'], 'Intuitive Surgical Inc.': [6902, 'NASDAQ', 'ISRG'], 'Johnson & Johnson': [7202, 'NYSE', 'JNJ'], 'KLA-Tencor': [7502, 'NASDAQ', 'KLAC'], 'Kraft Heinz Co': [7802, 'NASDAQ', 'KHC'], 'Laboratory of America Holdings': [8102, 'NYSE', 'LH'], 'Mastercard Inc.': [8402, 'NYSE', 'MA'], "McDonald's Corp.": [8702, 'NYSE', 'MCD'], 'Micron Technology': [9002, 'NASDAQ', 'MU'], 'Microsoft': [9302, 'NASDAQ', 'MSFT'], 'NetApp': [9602, 'NASDAQ', 'NTAP'], 'Netflix Inc.': [9902, 'NASDAQ', 'NFLX'], 'Nike': [10202, 'NYSE', 'NKE'], 'Nvidia Corporation': [10502, 'NASDAQ', 'NVDA'], 'Oracle Corp.': [10802, 'NYSE', 'ORCL'], 'Paychex': [11102, 'NASDAQ', 'PAYX'], 'PayPal': [11402, 'NASDAQ', 'PYPL'], 'S&P Global': [11702, 'NYSE', 'SPGI'], 'Starbucks Corp.': [12002, 'NASDAQ', 'SBUX'], 'Synopsys': [12302, 'NASDAQ', 'SNPS'], 'Tesla Inc': [12602, 'NASDAQ', 'TSLA'], 'United Parcel Service': [12902, 'NYSE', 'UPS'], 'Visa Inc.': [13202, 'NYSE', 'V'], 'Vertex Pharmaceuticals Inc': [13502, 'NASDAQ', 'VRTX'], 'Wells Fargo': [13802, 'NYSE', 'WFC'], 'Workday': [14102, 'NASDAQ', 'WDAY'], 'Zebra Technologies': [14402, 'NASDAQ', 'ZBRA'], 'Zoom': [14702, 'NASDAQ', 'ZM']}
def loopgraph():
    global x
    global start_time
    selected=namegraph.get()
    endtime=datetime.now()
    timmy=(endtime-start_time).total_seconds()
    x=timmy//4
    if selected!='':
        b=namedict[selected][0]
        if x>14:
            x=int(x)-14+b
            graph(x,str(worksheet.cell(int(x+14),0).value),selected)
        else:
            root3=Tk()
            root3.title('Error!')
            root3.geometry('160x50')
            ga=Label(root3,text='Waiting for data...').pack(side=TOP)

    else:
        root2=Tk()
        root2.title('Error!')
        root2.geometry('160x50')
        ha=Label(root2,text='Select an Option!').pack(side=TOP)

graphbtn=tk.Button(frame3a,text='Plot Graph',background='light gray',activebackground='gray',font='Calibri 14',command=lambda:loopgraph())
graphbtn.place(x=300,y=0)

canvas2.create_text(80,515,text='Stock:',fill='orange',font='Calibri 14')
name_select=Combobox(frame4,width=12,state='readonly',textvariable=var,font='Calibri 14',value=('Apple Inc.', 'American Tower', 'Amazon.com Inc.', 'Applied Materials', 'Broadcom', 'American Express Co', 'Boeing Company', 'CenturyLink Inc', 'Citigroup', 'Copart Inc', 'The Walt Disney Company', 'Equinix', 'Equifax', 'Facebook', 'FedEx Corporation', 'Fidelity National Information Services', 'Fiserv Inc', 'Alphabet Inc Class C', 'Goldman Sachs Group', "Honeywell Int'l", 'Huntington Ingalls Industries', 'Intel Corp.', 'Iron Mountain Incorporated', 'Intuitive Surgical Inc.', 'Johnson & Johnson', 'KLA-Tencor', 'Kraft Heinz Co', 'Laboratory of America Holdings', 'Mastercard Inc.', "McDonald's Corp.", 'Micron Technology', 'Microsoft', 'NetApp', 'Netflix Inc.', 'Nike', 'Nvidia Corporation', 'Oracle Corp.', 'Paychex', 'PayPal', 'S&P Global', 'Starbucks Corp.', 'Synopsys', 'Tesla Inc', 'United Parcel Service', 'Visa Inc.', 'Vertex Pharmaceuticals Inc', 'Wells Fargo', 'Workday', 'Zebra Technologies', 'Zoom'))
name_select.place(x=260,y=510)
name_select.current()
name_select.master.option_add('*TCombobox*Listbox.Background','gray')

nbal=10000
canvas2.create_text(80,20,text='Date:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,50,text='Account Balance:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,160,text='Name:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,200,text='Code:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,240,text='Exchange:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,280,text='Price:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,320,text='Quantity:',fill='Orange',font='Calibri 14')
canvas2.create_text(80,360,text='Total:',fill='Orange',font='Calibri 14')
initial=canvas2.create_text(320,50,text='USD:'+str(nbal),fill='white',font='Calibri 14')
quantityentry=Entry(frame4)
canvas2.create_window(330,560,window=quantityentry)
canvas2.create_text(80,560,text='Quantity:',fill='orange',font='Calibri 14')
ordname=canvas2.create_text(320,160,text='',fill='white',font='Calibri 14')
ordcode=canvas2.create_text(320,200,text='',fill='white',font='Calibri 14')
ordxchng=canvas2.create_text(320,240,text='',fill='white',font='Calibri 14')
cost=canvas2.create_text(320,280,text='0',fill='white',font='Calibri 14')
quantity=canvas2.create_text(320,320,text='0',fill='white',font='Calibri 14')
total=canvas2.create_text(320,360,text='0',fill='white',font='Calibri 14')
def buy():
    global z
    name=name_select.get()
    if name!='':
        x=z+namedict[name][0]
        cost=worksheet.cell(x,3).value
        quantity=quantityentry.get()
        if str.isdigit(quantity):
            quantity=int(quantity)
            total=cost*quantity
            return total
        else:
            total=0
            return total
    else:
        root6=Tk()
        root6.title('Error!')
        root6.geometry('160x50')
        ha=Label(root6,text='Select an Option!').pack(side=TOP)

mystocks=dict()
def bal(bfunc,task):
    global z
    global cost
    global quantity
    global total
    global ordname
    global ordcode
    global ordxchng
    ordtotal=bfunc()
    if ordtotal != 0:
        name=name_select.get()
        x=z+namedict[name][0]
        time=worksheet.cell(int(z),0).value
        canvas2.delete(cost,quantity,total,ordname,ordcode,ordxchng)
        ordcost=worksheet.cell(x,3).value
        ordname=canvas2.create_text(320,160,text=name,fill='white',font='Calibri 14')
        ordcode=canvas2.create_text(320,200,text=namedict[name][2],fill='white',font='Calibri 14')
        ordxchng=canvas2.create_text(320,240,text=namedict[name][1],fill='white',font='Calibri 14')
        cost=canvas2.create_text(320,280,text=ordcost,fill='white',font='Calibri 14')
        ordquantity=quantityentry.get()
        quantity=canvas2.create_text(320,320,text=ordquantity,fill='white',font='Calibri 14')
        total=canvas2.create_text(320,360,text=ordtotal,fill='white',font='Calibri 14')

        buybtn['state']=DISABLED
        sellbtn['state']=DISABLED
        def confirmcmd(taskboy):
            global initial
            global nbal
            confirm.destroy()
            cancel.destroy()
            canvas2.delete(initial)
            if taskboy=='BUY':
                if nbal>=ordtotal:
                    nbal=nbal+float(-ordtotal)
                    trade_cart(time,namedict[name][1],namedict[name][2],name,str(ordcost),str(ordquantity),'BUY')
                    if name in mystocks:
                        mystocks[name]+=int(ordquantity)
                    else:
                        mystocks[name]=int(ordquantity)
                else:
                    root5=Tk()
                    root5.title('Error!')
                    root5.geometry('300x50')
                    ye=Label(root5,text='Invalid Purchase! Check Account Balance.').pack(side=TOP)
            else:
                if name in mystocks:
                    if int(ordquantity)<=mystocks[name]:
                        nbal=nbal+float(ordtotal)
                        trade_cart(time,namedict[name][1],namedict[name][2],name,str(ordcost),str(ordquantity),'SOLD')
                        mystocks[name]-=int(ordquantity)
                    else:
                        root7=Tk()
                        root7.title('Error!')
                        root7.geometry('300x50')
                        ye=Label(root7,text='You are trying to sell more stocks than you own!').pack(side=TOP)
                else:
                    root5=Tk()
                    root5.title('Error!')
                    root5.geometry('300x50')
                    ye=Label(root5,text='You do not own the stock you are trying to sell!').pack(side=TOP)
            initial=canvas2.create_text(320,50,text='USD:'+str(nbal),fill='white',font='Calibri 14')
            buybtn['state']=NORMAL
            sellbtn['state']=NORMAL
        def cancelcmd():
            confirm.destroy()
            cancel.destroy()
            buybtn['state']=NORMAL
            sellbtn['state']=NORMAL
        confirm=tk.Button(frame4,text='Confirm Transaction',background='black',activebackground='gray',fg='white',font='Calibri 14',command=lambda:confirmcmd(task))
        confirm.place(x=40,y=770)
        cancel=tk.Button(frame4,text='Cancel Transaction',background='black',activebackground='gray',fg='white',font='Calibri 14',command=cancelcmd)
        cancel.place(x=240,y=770)
    else:
        root4=Tk()
        root4.title('Error!')
        root4.geometry('160x50')
        ye=Label(root4,text='Invalid Quantity!').pack(side=TOP)

buybtn=tk.Button(frame4,text='BUY',background='green',activebackground='lime',font='Calibri 14',width=6,command=lambda:bal(buy,'BUY'))
buybtn.place(x=80,y=620)
sellbtn=tk.Button(frame4,text='SELL',background='red',activebackground='orange red',font='Calibri 14',width=6,command=lambda:bal(buy,'SELL'))
sellbtn.place(x=280,y=620)

root.after(1196000,root.quit)
root.mainloop()

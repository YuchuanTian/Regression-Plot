#for plotting linear fit物理实验一条龙
import matplotlib.pyplot as plt
import scipy.stats as sp
import numpy as np
import math as m
import xlrd as xr
import xlwt as xw




def read_excel(filename):
    Ans=list()
    book = xr.open_workbook(filename)
    sheet = book.sheet_by_name('Sheet3')
    rows = sheet.nrows #获取行数
    cols = sheet.ncols #获取列数
    for c in range(cols): #读取每一列的数据
        Ans.append(sheet.col_values(c))
    return Ans
    #for r in range(rows): #读取每一行的数据
        #r_values = sheet.row_values(r)
        #print(r_values)
    #print(sheet.cell(1,1)) #读取指定单元格的数据




def linefit(x , y):
    N = float(len(x))
    sx,sy,sxx,syy,sxy=0,0,0,0,0
    for i in range(0,int(N)):
        sx  += x[i]
        sy  += y[i]
        sxx += x[i]*x[i]
        syy += y[i]*y[i]
        sxy += x[i]*y[i]
    a = (sy*sx/N -sxy)/( sx*sx/N -sxx)
    b = (sy - a*sx)/N
    r = abs(sy*sx/N-sxy)/m.sqrt((sxx-sx*sx/N)*(syy-sy*sy/N))
    return a,b,r
#checklist:
#CI half width


#data = pd.read_excel('data.xlsx',sheetname='Sheet1')
Ans=read_excel('data.xlsx')
t_95={3:12.706,4:4.303,5:3.182,6:2.776,
        7:2.571,8:2.447,9:2.365,10:2.306,
        11:2.262,12:2.228,15:2.160,20:2.101}

A=np.asarray(Ans[0])
B=np.asarray(Ans[1])
#errorbar length
errorx=np.ones(np.size(A))*0
errory=np.ones(np.size(B))*0.1


a,b,r,p,stderr=sp.linregress(A,B)
#a=1
#b=2
#r=3
print(a,b,r)

x=np.arange(0,A.max()/9*10,0.01)
y=a*x+b

data = {'a': np.arange(50),
        'c': np.random.randint(0, 50, 50),
        'd': np.random.randn(50)}
data['b'] = data['a'] + 10 * np.random.randn(50)
data['d'] = np.abs(data['d']) * 100

plt.errorbar(A,B,errory, errorx, 'b.')
plt.plot(x,y,'r')
plt.legend(['Linear Fits','Points'],loc='upper left')
#plt.scatter('a', 'b', c='c', s='d', data=data)
plt.xlabel('x')
plt.ylabel('y')
n=np.size(A)
print(n)
col_labels = ['value']
row_labels = ['Equation','Slope','Intercept','Pearson r', 'r-square','Std Error','CI H-width']
table_vals = [['y=ax+b'],[np.round(a,5)],[np.round(b,5)],[np.round(r,5)],[np.round(m.pow(r,2),5)],[np.round(stderr,5)],[np.round(stderr*t_95[n],5)]]
row_colors = ['white','white','white','white','white','white','white']
my_table = plt.table(cellText=table_vals, colWidths=[0.1]*2,
                     rowLabels=row_labels, colLabels=col_labels,
                     rowColours=row_colors, colColours=row_colors,
                     loc='center right')
plt.xlabel('l')#need input
plt.ylabel('F')#need input



plt.show()

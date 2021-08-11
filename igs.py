import os,re
from openpyxl import Workbook,load_workbook

def find_plane(txt,sheet,row):

    ans = re.finditer('110,(.*,.*,.*D-?\d).*\n(.*,.*,.*D-?\d).*\n110,(.*,.*,.*D-?\d).*\n.*\n110,.*\n.*\n110,.*\n.*\n128,.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n102.*\n110,.*\n.*\n110,.*\n.*\n110,.*\n.*\n110,.*\n.*\n102.*\n142.*\n144.*\n',txt)
    ans=list(ans)

    #输出
    for i in ans:
        col = 1
        for o in i.group(1,2,3):
            p = chaifen(o)
            for u in p:
                sheet[row][col].value = u
                col += 1
        row += 1

def find_line(txt,sheet,row):
    ans = re.finditer('126,1,1,1,0,0,0,0D0,0D0,1D0,1D0,1D0,1D0,(-?\d\.\d+D-?\d).*\n(-?\d\.\d+D-?\d,-?\d\.\d+D-?\d),(-?\d\.\d+D-?\d).*\n(-?\d\.\d+D-?\d,-?\d\.\d+D-?\d),.*\n.*\n126,.*\n.*\n.*\n.*\n.*\n',txt)
    ans=list(ans)
    #输出
##    print(ans)
    for i in ans:
##        print(i.group(1))
        col = 1 + 9
        p = chaifen(i.group(1)+i.group(2))
        for u in p:
            sheet[row][col].value = u
            col += 1
        p = chaifen(i.group(3)+i.group(4))
        for u in p:
            sheet[row][col].value = u
            col += 1
        row += 1
    
def find_point(txt,sheet,row):
    ans = re.finditer('116,.*\n',txt)
    ans=list(ans)
    #输出
    for i in ans:
        col = 1+9+6
        for o in chaifen(i.group()):
            sheet[row][col].value = o
            col += 1
        row += 1

##拆分 'x,y,z' -> x,y,z
def chaifen(ans):
    ps = re.findall('-?\d\.\d+D-?\d',ans)
    ped = []
    for i in ps:
        ped.append(huanyuan(i))
    return ped

##还原数字 '1.3333D1' -> 13.333
def huanyuan(p):
    t = re.search('D',p)
    loc = t.span()[0]
    num = float(p[0:loc-1]) * 10**(int(p[loc+1:]))
    return num

##遍历目录
def walk(path = '.',keyword = None):
    for a, b, c in os.walk(path):
        break
    d = []

    if keyword == None:
        return c
    
    for i in c:
        if re.search(keyword,i) != None:
            d.append(i)
    return d

files = walk('.','.igs')

rd = load_workbook('temp.xlsx',data_only = 1)

for file in files:
    filename = file
    print(filename)
    f = open(filename)
    txt = f.read()
    row = 3

    #按文件名写入同名sheet, 不存在则复制模板
    try:
        sheet = rd[filename]
    except:
        sheet = rd.copy_worksheet(rd['Sheet1'])
        sheet.title = filename
        
    plane = find_plane(txt,sheet,row)
    line = find_line(txt,sheet,row)
    point = find_point(txt,sheet,row)
    
    rd.save('temp2.xlsx')
rd.close

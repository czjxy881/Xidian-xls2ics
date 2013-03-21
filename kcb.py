# coding:utf-8
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import *
from xls2ics import *
import xlrd,re,Pmw
x=0
def Fileopen(): #打开文件
    r=askopenfilename(title='选择课程表',filetypes=[('Excel文件','*.xls *.xlsx')])
    eny_path.delete(0,END)
    eny_path.insert(0,r)
def load():
    global file_path,x,f,Sheet,Dic,now
    s=eny_path.get()
    if s[-3:]!="xls" and s[-4:]!="xlsx":
        showerror('错误','文件格式错误')
    else:
        try:
           x=xlstoics(s)
        except:
           showerror('错误','文件打开失败')
        Sheet=[]
        Dic=[]
        now.set(0)
        f.destroy()
        f=Frame(root,width=385,height=220,relief=GROOVE,borderwidth=2)
        f.place(relx=0.02,y=65)
        Label(root,text='具体课程:').place(relx=0.06,y=55,anchor=NW)
        ff=Pmw.ScrolledFrame(f,labelpos=N,label_text=" ",usehullsize=1,hull_height=75,hull_width=380)
        ff.pack()
        a=ff.interior()
        width=14
        try:
         for i in range(len(x)):
            sheet=x.opensh(i)
            Sheet.append(sheet)
            Dic.append(x.xls(sheet))
            Radiobutton(a,variable=now,value=i,indicatoron=0,width=width,text=sheet.name,command=change).grid(row=0,column=i)
        except:
            showerror('错误','读取错误，请检查格式或联系作者')
        change()
        but_der =Button(root,text='导出',command=der)
        but_der.place(anchor=NE,x=370,y=275,width=50,height=25)
        Label(root,text='所有ics均导出到本程序所在目录下').place(relx=0.43,y=55,anchor=NW)
def change():#项目卡变化
   global now,ssf
   t=now.get()
   ssf.destroy()
   ssf=Pmw.ScrolledFrame(f,usehullsize=1,hull_height=132,hull_width=380)
   ssf.pack()
   sf=ssf.interior()
   num=0
   for i in Dic[t]:
      # print i
       l=Dic[t][i]
      # print l
      # print l[0]
       cc=Checkbutton(sf,text=l[0],command=(lambda i=i:fuch(i)))
       if l[9]==1:cc.select()
       cc.grid(row=num%4,column=num/4,sticky=NW)
       num+=1

def fuch(i):#复选框变化
    global now
    t=now.get()
   # print i
    Dic[t][i][-1]=1-Dic[t][i][-1]
   # print Dic[t][i][-1]

def der():
    global x
    t=0;a=0
    try:
     for i in range(len(x)):
        num=0
        for j in Dic[i]:
            if Dic[i][j][-1]==1:num+=1
        if num==0:continue
        a+=x.ics(Dic[i],Sheet[i])
        t+=num
     showinfo('成功','成功导出%d门课,共%d次活动。\nPS:每个分页均有一个导出文件!'%(t,a))
    except:
      showerror('错误','导出失败，请检查写入权限或联系作者')

file_path=""
root=Tk() #主界面
w,h = root.maxsize()
w-=400
h-=300
root.geometry("400x300+%d+%d"%(w/2,h/2))
root.resizable(False,False)
root.title("西电课表导入 作者Email:czjxy8898@gmail.com")
lab_excel=Label(root,text="Excel文件路径",anchor=W)
lab_excel.pack(anchor=NW,fill=X)
eny_path=Entry(root,width='48') #路径输入框
eny_path.pack(anchor=NW,side="left")
but_path=Button(root,text="...",command=Fileopen)
but_path.pack(anchor=NE,side='left',fill=X)
but_load=Button(root,text="载入",command=load)
but_load.pack(anchor=NW,side='left')
#读入框完成


f=Frame(root,width=385,height=220,relief=GROOVE,borderwidth=2)
f.place(relx=0.02,y=65)
Label(root,text='选择课程:').place(relx=0.06,y=55,anchor=NW)
Label(root,text='请选择EXCEL文件路径！').place(relx=0.34,rely=0.5,anchor=NW)
ssf=Frame(f,width=380,height=140)
Sheet=[]#存放所有sheet
Dic=[]#存放所有dic
now=IntVar()#现在选中的sheet
vv=StringVar()#复选框

root.mainloop()




















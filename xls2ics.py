# -*- coding: utf-8 -*-
import xlrd,re
from tkMessageBox import *                          
class xlstoics():
  file_path=""
  def __init__(self,path):
    self.file_path=path
    #print path
    self.file=xlrd.open_workbook(self.file_path,formatting_info=True)
  
  def event(self,start,end,name,detail,location): #一次事件
       return '''
BEGIN:VEVENT
DTSTART:%s
DTEND:%s
CLASS:PUBLIC
DESCRIPTION:%s
LOCATION:%s
SUMMARY:%s
END:VEVENT
'''%(start,end,detail,location,name)
      
  def __len__(self):
    return len(self.file.sheets())
  def name(self):
    return self.file.sheets()
    
  #xls导入 打开num页
  def opensh(self,num): 
    sheet=self.file.sheets()[num]
    return sheet
  #
  #传入参数：表页面
  #传出课程代号为索引的课程名称Map
  def xls(self,sheet):  
    rows=sheet.nrows;
    cols=sheet.ncols;
    dic={} #课程代号为索引的课程详情
    #获得课程详情
    #获得课程详情的起始位置
    for i in range(0,rows): 
      if sheet.row_values(i)[0].replace(' ','')==('课程'.decode('utf-8')):
        rowt=i+2
        break
    ind=[] #确定详情的位置索引
    for i in range(0,cols):
      t=sheet.cell(rowt-1,i).value+sheet.cell(rowt-2,i).value
      if t!="" and t!="考核\n方式".decode('utf-8'): #选修没有考核方式
        ind.append(i);
    
    for i in range(rowt,rows):
       s=sheet.row_values(i)
       j=0;l=[];key="";flag=1
       if s[0].replace(' ','')=="备注".decode('utf-8'):break #课程结束
       for j in range(9):
         now=s[ind[j]]
         if isinstance(now,float)==True:now=str(now);
         now="".join(now.split(" "))
         if j==1:
            if now=='':flag=0;break; #代表没有课程编号
            key=now
            #print now
            if '0'<now[-1].encode('utf-8')<'9':l[0]+=now[-1].encode('utf-8')
            j+=1;
         l.append(now.replace(' ','').replace('\n','').encode('utf-8'))
       l.append(0)#选中标志
       if flag==1:dic[key]=l
       #print;
    return dic
  dd=[0,31,28,31,30,31,30,31,31,30,31,30,31]
  nd=0;ny=0;nm=0
  def aday(self): 
      self.nd+=1;flag=0
      if self.nm==2 and self.ny%4==0 and (self.ny%100!=0 or self.ny%400==0):flag=1;
      if(self.nd>self.dd[self.nm]+flag):
        self.nm+=1;self.nd=1;
        if self.nm>12:self.nm=1;self.ny+=1;
  def ics(self,dic,sheet):
    rows=sheet.nrows;
    cols=sheet.ncols;
    for i in range(0,rows):
      if sheet.row_values(i)[0].replace(' ','')==('课程'.decode('utf-8')):
        rowt=i+2
        break
    rs=3 #开始行数
    #确定学期起始时间
    #month对应中文
    mm={'一月':1,"二月":2,'三月':3,'四月':4,'五月':5,'六月':6,'七月':7,'八月':8,
       '九月':9,'十月':10,'十一月':11,'十二月':12}
    while isinstance(sheet.cell(rs,2).value,float)!=True:
      rs+=1
    rs-=1
    s=sheet.cell(rs,2).value.encode('utf-8')
    while s[-1]==' ':s=s[:-1] #月份末尾可能出现空格…………
    month=mm[s]
    s=sheet.cell(rs+2,2).value
    day=int(re.split('\D',s,1)[0])
    c=sheet.name
    s=sheet.cell(rs-2,17).value;
    #print month,day
    while s[0]==' ':s=s[1:]
    icsn=(c+'-'+s).encode('utf-8'); #课程表名称
    year=int(re.split('\D',s,1)[0])
    
    #ics文件开始
    ics=open('%s.ics'%icsn.decode('utf-8').encode('cp936'),'w');
    ics.write('''BEGIN:VCALENDAR
METHOD:PUBLISH
X-WR-CALNAME:%s
X-WR-TIMEZONE:Asia/Shanghai
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE'''%icsn)
    
    
    #确定学期周数
    weeks=cols-1;
    while sheet.cell(rs+1,weeks)=="": weeks-=1
    weeks-=2
    rowt-=2 #确定课表边界
    
    #扩展课程里的合并单元格
    for block in sheet.merged_cells:
      u,d,l,r=block
      if l>1 and  r<weeks and u>rs+2 and d<rowt:
        cell=sheet.cell(u,l)
        a=cell.ctype;
        b=cell.value;
        if b=='节日'.decode('utf-8'):continue;
        #print l,r,u,d,b
        for i in range(l,r):
          for j in range(u,d):
             sheet.put_cell(j,i,a,b,0)
    
    #确定课程时间
    #冬季
    BW=['003000','022500','060000','075500','110000']
    EW=['020500','040000','073500','093000','123000']
    #夏季
    BS=['003000','022500','063000','082500','113000']
    ES=['020500','040000','080500','100000','130000']
    B=[BS,BW];E=[ES,EW]
    
    #确定课程
    Name=['','','学时:','学分:','性质:','主讲老师:','职称:','教室:','']
    nn=0
    self.ny=year;self.nm=month;self.nd=day;
    que=sheet.col_values(1)
    ddd=sheet.col_values(0)
    for i in range(2,weeks+2):
      s=sheet.col_values(i);
      k=0;f=1;
      

      if self.nm>4 and self.nm<10:f=0
      #print rs,rowt
      for j in range(rs+3,rowt):
        #确定节次
        p=que[j][0]
      #  print p
        if p=='1':nt=0;
        elif p=='3':nt=1;
        elif p=='5':nt=2;
        elif p=='7':nt=3
        else:nt=4
        if k==0:today='%4d%02d%02dT'%(self.ny,self.nm,self.nd)
       # print today
       # print k,s[j],today
        k+=1
        #print ddd[j+1]
        if ddd[j+1]!='':self.aday();k=0;
        s[j]="".join(s[j].split()) #去除空格
        if s[j]=='':continue
        name=""
        ff=0
        if s[j][-1]=='0':
          s[j]=s[j][:-1];name="考试:";ff=1
        try:
          l=dic[s[j]]
         # if s[j]=='电实'.decode('utf-8'):
          # print l
        except:
          #print s[j]
          continue
        if l[-1]==0:continue #去除未选中项
        name+=l[0]
        detail=''
        for t in range(2,9):
          if l[t]=='':continue;
          if ff==1 and t==8:continue;
          detail+=Name[t]+l[t]+'\\n'
        #print detail
        if ff==1:
          detail='考试！'
          l[8]=''
        ics.write(self.event(today+B[f][nt]+'Z',today+E[f][nt]+'Z',name,detail,l[7]))
        nn+=1
     # self.aday()
      self.aday() #周日
    ics.writelines('END:VCALENDAR')
    ics.close()
    return nn

if __name__=='__main__':
  xlrd.open_workbook('全校2013年下学期人文素质限选课课表.xls'.decode('utf-8'),formatting_info=True)
  x=xlstoics('全校2013年下学期人文素质限选课课表.xls'.decode('utf-8'))
  s=x.opensh(1)
  dic=x.xls(s)
  print x.ics(dic,s)

      

#设备状态检测V1.0---by yjs @Y161129
#import工具库
import tkinter
from tkinter.font import Font
import socket
from html.parser import HTMLParser
import requests
import threading
import queue
import os,re,time



#使用类g，设置全局变量
class g(object):
    is_ping = False
    is_connect = False
    is_login = False
    name={'HB-1':'HB-1',
    'HB-2':'HB-2',
    'HB-3':'HB-3',
    'HB-4':'HB-4',
    'C2263':'C2263',
    'M261N':'M261N'}

    dev={'HB-1':'10.100.1.51',
    'HB-2':'10.100.1.52',
    'HB-3':'10.100.1.53',
    'HB-4':'10.100.1.54',
    'C2263':'10.100.1.190',
    'M261N':'10.100.1.44'}

    payload = {'ACTION_POST':'LOGIN','LOGIN_USER':'admin','LOGIN_PASSWD':'jf2012','login':'登录'}
    #payload={'ACTION_POST':'LOGIN','LOGIN_USER':'admin','LOGIN_PASSWD':'jf20112','login':'登录'} #错误密码测试
    color = {'HB-1':'green',
    'HB-2':'green',
    'HB-3':'green',
    'HB-4':'green',
    'C2263':'green',
    'M261N':'green'} #设置颜色： 绿色-正常；橙色-故障；红色-离线；
    count={'HB-1':0,
    'HB-2':0,
    'HB-3':0,
    'HB-4':0} #统计客户端连接数



#可连接性测试
def connect(ipdict,port=80):
    for key,ip in ipdict.items():
        #连接ip
##        print(ip,end='')
        with socket.socket() as s:
            ip_port=(ip,port)
            r_code=s.connect_ex(ip_port)
            if r_code==0: g.is_connect=True
            if r_code==10061: g.is_ping=True
##            print(':{: <5}{}'.format(port,r_code))
            
            #需要登录
            if g.count.get(key,-1)!=-1: 
                if g.is_connect:
                    login(key,ip)
                    g.is_connect = False
                elif g.is_ping:
                    setOrange(key)
                else:
                    setRed(key) 
                    

            #不需登录
            else:
                if g.is_connect or g.is_ping:
                    setGreen(key)
                    g.is_connect=False
                    g.is_ping=False
                else:
                    setRed(key)




#登录测试
def login(key,ip):
    #登录
    url='http://{addr}'.format(addr=ip)
    with requests.Session() as s:
        r=s.post(url+'/login.php',data=g.payload)
        if r.status_code==200 and r.text.find('login_fail.php')==-1:g.is_login=True
        if g.is_login: #登录成功，恢复标志is_login
            num=get_num(s,url) #获取客户端连接数
            g.count.update({key:num})
##            print(ip,'  num:',num) #成功输出
            setGreen(key)
            g.is_login = False
        else:#登录失败
##            print(ip,'login fail!') #失败输出
            setOrange(key)

def setGreen(key):
    g.color.update({key:'green'})

def setOrange(key):
    g.color.update({key:'orange'})

def setRed(key):
    g.color.update({key:'red'})
#获取客户端连接数           
def get_num(session,url): #session 为requests.Session对象
    #获取num
    s=session
    s.get(url+'/index.php')
    gu=s.get(url+'/st_info.php')
    if gu.encoding.lower()!='utf-8':
        gu.encoding='utf-8'
    num=num_inhtml(gu.text)
    return num


class flag(object):
    num = None
    is_got = False
        
class myParser(HTMLParser):
    def handle_starttag(self,tag,attrs):
        tag_text = 'td'
        attr_text = 'td_right'
        if tag_text==tag:
            for attr in attrs:
                if attr_text in attr:
                    flag.is_got=True
    def handle_data(self, data):
        if flag.is_got:
            for v in data.splitlines():
                try:
                    flag.num=int(v)
                except Exception:
                    continue
            flag.is_got=False

    
def num_inhtml(html): #根据html分析提取客户端连接数
    num = None
    parser = myParser()
    parser.feed(html)
    num,flag.num = flag.num,num
    return num   


#monitor主面板    

class App(tkinter.Frame):
    def __init__(self,master=None):
        tkinter.Frame.__init__(self,master)
        self.name=['HB-1','HB-2','HB-3','HB-4','C2263','M261N']
        self.dev=list()
        self.status=list()
        self.qu=queue.Queue()
        self.bind('<<TimeChanged>>',self.timeChanged)
        self.grid()
        self.Create_widgets() 
##        self.bt=tkinter.Button(self,text='run',width=10,command=self.getdata)
##        self.bt.grid(column=0,columnspan=6)
        self.update()
        
    def Create_widgets(self):
        self.ap_font=Font(size=24,weight='bold')
        for i,v in enumerate(self.name):
            self.dev.append(tkinter.LabelFrame(self,text=v,padx=30,width=50))
            self.status.append([tkinter.Label(self.dev[i],text='●',font=self.ap_font,fg='green'),tkinter.Label(self.dev[i])])
            self.status[i][0].grid()
            self.status[i][1].grid()
            self.dev[i].grid(row=0,column=i)
    ## Use a binding on the custom event to get the new time value
    ## and change the variable to update the display



    def getdata(self):
        th=threading.Thread(target=self.timeThread())
        th.start()
        th.join()

            
    def timeThread(self):
        while True:
            ## Each time the time increases, put the new value in the queue...
            connect(g.dev)
            self.qu.put(g.count)

            ## ... and generate a custom event on the main window
            try:
                self.event_generate('<<TimeChanged>>')
##                print('ok')
            ## If it failed, the window has been destoyed: over
            except Exception:
                break
            ## Next
            time.sleep(1)
            
    def timeChanged(self,event):
        count=self.qu.get()
##        print('count=',count)
        for i,k in enumerate(self.name):
            self.status[i][0]['fg']=g.color[k]
            if 'HB' in k:
                self.status[i][1]['text']=count.get(k,0)
            self.dev[i].update()
##        self.master.update()



#运行tk图形界面
def runtk():

    mon=App()
##    mon.master.geometry('600x400')
    mon.master.title('HB dev_status monitor App')
    mon.getdata()
    mon.master.mainloop()

def test():
    #connect({'hb-1':'10.100.1.51'}) 
    connect(g.dev)
##    print(g.color)


#运行代码
if __name__=="__main__":
    
    runtk()




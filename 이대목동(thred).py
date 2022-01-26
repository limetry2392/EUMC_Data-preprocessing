import threading

import tkinter
import tkinter.font
import tkinter.ttk as ttk

import pandas as pd
from urllib import request
import os




class BackgroundTask():

    def __init__( self, taskFuncPointer ):
        self.__taskFuncPointer_ = taskFuncPointer
        self.__workerThread_ = None
        self.__isRunning_ = False

    def taskFuncPointer( self ) : return self.__taskFuncPointer_

    def isRunning( self ) : 
        return self.__isRunning_ and self.__workerThread_.isAlive()

    def start( self ): 
        if not self.__isRunning_ :
            self.__isRunning_ = True
            self.__workerThread_ = self.WorkerThread( self )
            self.__workerThread_.start()

    def stop( self ) : self.__isRunning_ = False

    class WorkerThread( threading.Thread ):
        def __init__( self, bgTask ):      
            threading.Thread.__init__( self )
            self.__bgTask_ = bgTask

        def run( self ):
            try :
                self.__bgTask_.taskFuncPointer()( self.__bgTask_.isRunning )
            except Exception as e: print (repr(e))
            self.__bgTask_.stop()

def center_window(root, width, height):
  screenwidth = root.winfo_screenwidth()
  screenheight = root.winfo_screenheight()
  size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
  root.geometry(size)

def tkThreadingTest():
    from tkinter import Tk, Label, Button, StringVar, Frame, messagebox, filedialog
    from time import sleep
     
    class UnitTestGUI:
           
        def __init__( self, master ):
            global font, font2, font3, font4, menubar
            font = tkinter.font.Font(family="현대하모니L", size=22, weight="bold")
            font2 = tkinter.font.Font(family="현대하모니L", size=8, slant="italic", weight="bold", overstrike = 'TRUE')
            font3 = tkinter.font.Font(family="현대하모니L", size=8, slant="italic", weight="bold")
            font4 = tkinter.font.Font(family="Times", size=9, slant="italic", weight="bold")

            
 
            self.master = master
            master.title( "Weekly analysis tool" )

            menubar = tkinter.Menu(master)

            editmenu = tkinter.Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=editmenu)
            editmenu.add_command(label="Open", command = self.onFileClicked)
            editmenu.add_separator()
            editmenu.add_command(label="Exit", command = self.close)

            filemenu = tkinter.Menu(menubar, tearoff=0, bg = 'white')
            menubar.add_cascade(label="Help", menu=filemenu)
            filemenu.add_command(label="How to use", command = self.how)
            filemenu.add_separator()
            filemenu.add_command(label="LICENSE", command = self.lic)

            master.config(menu=menubar)


            self.testFame = Frame( background = "white",
                                         relief='ridge',
                                          height=100,
                                          width=80) 
            self.testFame.pack(side='bottom', fill='both')

            
            self.titleLabelVar = StringVar()
            self.titleLabel = Label( master, text = "이대목동 임상실험자 주간 평균 자료 생성",
                                  font=font, 
                                  foreground='white',
                                  background='gray20',
                                  relief="flat",
                                  bd = 3,
                                  padx = 40,
                                  pady = 3)
            self.titleLabel.pack(fill='both')

            self.fileLabelVar = StringVar()
            self.fileLabel = Label( master, 
                                       bg = 'white',
                                       foreground='black',
                                       pady=4,
                                       font=("현대하모니L", 8, "bold"),
                                       borderwidth=2)
            self.fileLabel.pack(fill='both')

            self.testLabelVar = StringVar()
            self.testLabel = Label( self.testFame, 
                                       bg = 'white',
                                       foreground='black',
                                       pady=4,
                                       font=("현대하모니L", 8, "bold"),
                                       borderwidth=2)
            self.testLabel.pack()

            self.test2LabelVar = StringVar()
            self.test2Label = Label( master, 
                                       font = font4,
                                       bg = 'white',
                                       foreground='black',
                                       pady=4,
                                       borderwidth=2)
            self.test2Label.configure(text="Made By. Yong-Lim")
            self.test2Label.place(x=495, y=300)


            self.fileButton = Button( 
                self.master, 
                text="Load MetaData", 
                background = 'white',
                fg = 'black',
                activebackground = 'gray',
                relief = 'raised',
                bd = 2.5,
                activeforeground = 'white',
                overrelief = 'groove',
                width = 15, 
                height = 3,
                font=font3,
                command=self.onFileClicked )
            self.fileButton.pack(pady = 5)

            self.threadedButton = Button( 
                self.master, 
                text="Data Download", 
                background = 'white',
                fg = 'black',
                activebackground = 'gray',
                relief = 'sunken',
                bd = 2.5,
                activeforeground = 'white',
                overrelief = 'groove',
                width = 15, 
                height = 3,
                state = 'disabled',
                font=font2,
                command=self.onThreadedClicked )
            self.threadedButton.pack(pady = 5)


            self.weekButton = Button( 
                self.master, 
                text="Weekly Average", 
                background = 'white',
                fg = 'black',
                activebackground = 'gray',
                relief = 'sunken',
                bd = 2.5,
                activeforeground = 'white',
                overrelief = 'groove',
                width = 15, 
                height = 3,
                state = 'disabled',
                font=font2,
                command=self.onThreaded2Clicked )
            self.weekButton.pack(pady = 5)


            self.bgTask = BackgroundTask( self.APIProcess )
            self.bg2Task = BackgroundTask( self.weeklyProcess )

        def close( self ) :
            MsgBox = messagebox.askquestion ('Exit App','Really Quit?',icon = 'info')
            if MsgBox == 'yes':
               try: self.bgTask.stop()
               except: pass
               try: self.bg2Task.stop()
               except: pass
               self.master.quit()
            else:
               messagebox.showinfo('Welcome Back','Welcome back')

        def how( self ):
            toplevel = tkinter.Toplevel(root)
            center_window(toplevel, 380, 280)
            toplevel.configure(bg='white')
            toplevel.title("How to use...")
            toplevel.resizable(width=False, height=False)

            lbl = Label(toplevel, text="★사용방법★",bg='white', font=('현대하모니L', 22, 'bold'))
            lbl.pack()    

            txt = Label(toplevel, bg='white', foreground='black',
                        justify='left',
                        relief='ridge',
                        padx=10,
                        pady=10,
                        text="\n1-1. 메타데이터를 불러온다.\n\n2-1. 데이터 다운로드를 진행한다.\n\n2-2. 메타데이터 양식이 맞지 않으면, \n      데이터가 다운로드 되지 않습니다.\n      ※다운 받아진 양식에 맞춰 다시 진행하시면 됩니다.\n\n2-3. 메타데이터가 비어있어도 다운로드 되지 않습니다.\n\n3-1. 데이터 다운로드가 완료되면 주간 평균 진행하시면 됩니다.\n")
            
            txt.pack()  

        def lic( self ):
            toplevel2 = tkinter.Toplevel(root)
            center_window(toplevel2, 600, 400)
            toplevel2.configure(bg='white')
            toplevel2.title("LICENSE")
            toplevel2.resizable(width=False, height=False)
            txt = Label(toplevel2, text="The license belongs to Yonglim",bg='white', foreground='black', font=font)
            txt2 = Label(toplevel2, 
                         justify='left', 
                         relief='ridge', 
                         padx= 10,
                         text='\nMIT\nCopyright (c) 2021 CHO-YONG-LIM (Yong)\nPermission is hereby granted, free of charge, to any person obtaining a copy\nof this software and associated documentation files (the "Software"), to deal\nin the Software without restriction, including without limitation the rights\nto use, copy, modify, merge, publish, distribute, sublicense, and/or sell\ncopies of the Software, and to permit persons to whom the Software is\nfurnished to do so, subject to the following conditions:\n\nThe above copyright notice and this permission notice shall be included in all\ncopies or substantial portions of the Software.\n\nTHE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\nIMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\nFITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\nAUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\nLIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\nOUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE\nSOFTWARE.',bg='white', foreground='black')
            txt.pack(pady=5)
            txt2.pack()

        def onFileClicked( self ):
            global metadata
            root.file = filedialog.askopenfile(
            initialdir='path', 
            title='select file', 
            filetypes=(('csv files', '*.csv'), 
            ('all files', '*.*')))
    
            A = "Load Complete!"   
            B = '0126'

            vvv = str(root.file)

            if(vvv == 'None'):
                messagebox.showerror("Load...",'Cansle')
            else:
                self.fileLabel.configure(text="Metadata path: " + root.file.name)
                self.fileLabel['relief'] = tkinter.SUNKEN
                self.fileLabel['background'] = 'light gray'
                metadata = root.file.name
                metadata = pd.read_csv(metadata, dtype=str, encoding='CP949')
                messagebox.showinfo("Metadata Load...",A)
                AA = 1

                if(B == '0126'):
                     self.threadedButton['state'] = tkinter.NORMAL 
                     self.threadedButton['relief'] = tkinter.RAISED
                     self.threadedButton['font'] = font3
                     
                else:
                     self.threadedButton['state'] = tkinter.DISABLED  
        

        def onThreadedClicked( self ):
            try: self.bgTask.start()
            except: pass

        def onThreaded2Clicked( self ):
            try: self.bg2Task.start()
            except: pass

        def APIProcess( self, isRunningFunc=None) :

            try:
               if not isRunningFunc() :
                   return
            except : pass

            self.threadedButton['state'] = tkinter.DISABLED
            self.fileButton['state'] = tkinter.DISABLED
            self.weekButton['state'] = tkinter.DISABLED

            aaa = metadata.columns
            aaa = str(aaa)[6:-17]
            bbb = ['측정기시리얼넘버', '대상자관리번호', '대상자이름', '측정기설치장소', '시작일', '종료일', '주소']
            bbb = str(bbb)
            if (aaa == bbb):
                A = 0
                h = metadata.shape[0]
        
                if (h <= 1):
                     try:
                        if not isRunningFunc() :
                            return
                     except : pass
                     messagebox.showerror("ERROR",'DATA EMPTY')
                     self.threadedButton['state'] = tkinter.NORMAL
                     self.fileButton['state'] = tkinter.NORMAL
            
                else:
                     self.testLabel.configure(text="Data Downloading... ")  
                     progressbar = ttk.Progressbar(self.testFame, style='yong.Horizontal.TProgressbar', length=150, maximum=150, mode="indeterminate")
                     progressbar.start(10)
                     progressbar.pack(pady = 8)         
                     while A < h:
                         try:
                            if not isRunningFunc() :
                                return
                         except : pass
 
       
                         check = metadata.종료일[A]
                         
                         if(pd.isna(check)):
                              metadata.대상자관리번호[A] = str(metadata.대상자관리번호[A])
                              Z = A + 1
                              Z = str(Z)
                              aaaaa = Z + '. ' +metadata.대상자관리번호[A] + '->' + 'There is no end date data.\n'
                              #print(aaaaa)
                         else:
                              Z = A + 1
                              Z = str(Z)
                              ccc = Z + '. ' + metadata.대상자관리번호[A] + ' -> ' + 'Importing Data...\n'


                              main_url = 'https://datacenter.kweather.co.kr/api/collection/data/excel/download'  #고정
    
                              serial = metadata.측정기시리얼넘버 [A]
                              serial1 = 'serial'+ '=' + serial                                           #시리얼번호
        
                              from datetime import datetime
    
                              startTime = metadata.시작일[A]
                              startTime = datetime.strptime(startTime, '%Y-%m-%d').strftime('%Y/%m/%d')
                              startTime = 'startTime=' + startTime + '-00:00:00'                                      #시작시간
    
                              endTime = metadata.종료일[A]
                              endTime = datetime.strptime(endTime, '%Y-%m-%d').strftime('%Y/%m/%d')
                              endTime = 'endTime=' + endTime + '-23:59:59'                                           #종료시간
    
                              standard = 'standard=sum'                                                         #시간평균 설정
    
                              deviceType = 'deviceType=iaq'                                                     #장비종료
        
    
                              variable = startTime + '&' +  endTime + '&' + standard + '&' + serial1 + '&' + deviceType
    
                              api_url = main_url + '?' + variable
    
                              file = serial + '-' + metadata.대상자관리번호[A]
                              file = './' + file + '.' + 'xlsx'
        
                              request.urlretrieve(api_url, file)
                         try:
                            if not isRunningFunc() :
                                return
                         except : pass
 
                         A = A + 1 
                progressbar.destroy()
                progressbar = ttk.Progressbar(self.testFame,  length=150, maximum=150, value=150)
                progressbar.pack(pady = 8)
                A = "Data Download Complete!"   
                self.testLabel.configure(text="Done!")
                messagebox.showinfo("File Load",A)
                progressbar.destroy()
                self.testLabel.configure(text="")
                self.threadedButton['state'] = tkinter.NORMAL
                self.fileButton['state'] = tkinter.NORMAL
                self.weekButton['state'] = tkinter.NORMAL
                self.weekButton['relief'] = tkinter.RAISED 
                self.weekButton['font'] = font3
                
            else:
                 messagebox.showerror("Data Download...","Data Download Fail\nIt doesn't fit the style.\nPlease fill it out according to the form.") 
                 metadataaa = pd.DataFrame([['필수 값','필수 값','','','YYY-MM-DD','YYY-MM-DD','']], columns=['측정기시리얼넘버',
                                                                                           '대상자관리번호',
                                                                                           '대상자이름',
                                                                                           '측정기설치장소',
                                                                                           '시작일',
                                                                                           '종료일',
                                                                                           '주소'])
                 metadataaa.to_csv('./[SAMPLE] MetaData.csv' ,index=False, encoding='cp949')
                 self.threadedButton['state'] = tkinter.NORMAL
                 self.fileButton['state'] = tkinter.NORMAL


        def weeklyProcess( self, isRunningFunc=None) :

            try:
               if not isRunningFunc() :
                   return
            except : pass

            self.testLabel.configure(text="Weekly Processing... ")
            progressbar_indeter = ttk.Progressbar(self.testFame, length=150, maximum=150, mode="indeterminate")
            progressbar_indeter.start(10)
            progressbar_indeter.pack(pady = 8)

            self.threadedButton['state'] = tkinter.DISABLED
            self.fileButton['state'] = tkinter.DISABLED
            self.weekButton['state'] = tkinter.DISABLED

            path = './'
            file_list = os.listdir(path)
            file_list_py = [file for file in file_list if file.endswith('.xlsx')]
            VB = 0
            while(VB < len(file_list_py)):
                try:
                   if not isRunningFunc() :
                       return
                except : pass
        
                file = (file_list_py[VB])
                n = pd.read_excel(file, engine='openpyxl')
                n['데이터 시간'] = pd.to_datetime(n['데이터 시간'])
        
                import datetime
        
                te = n.shape[0]
                te = te - 1
                te = n['데이터 시간'][te]

                test = te + datetime.timedelta(days=1)
                c = datetime.timedelta(days=0)
                h_mean = pd.DataFrame()
                wedate = n['데이터 시간'][1]
                wedate = str(wedate)
                wedate = wedate[0:11]
                test2 = str(test)
                test2 = test2[0:11]
        
                wsdate = n['데이터 시간'][1]
        
                while(wedate <= test2):
                    try:
                       if not isRunningFunc() :
                           return
                    except : pass
            
                    wsdate = wsdate + c
                    wsdate2 = str(wsdate)
                    wsdate2 = wsdate2[0:11]

                    wedate = wsdate + datetime.timedelta(days=7)
                    wedate2 = str(wedate)
                    wedate2 = wedate2[0:11]
            
                    wedate3 = wedate -  datetime.timedelta(days=1)
                    wedate3  = str(wedate3)
                    wedate3 = wedate3[0:11]
            
                    week = n[(n['데이터 시간'] >= wsdate2) & (n['데이터 시간'] < wedate2)]
            
                    aa = wsdate2 + '~' + wedate3
                    bb= week['미세먼지(ug/m³)'].mean()
                    cc= week['초미세먼지(ug/m³)'].mean()
                    dd= week['이산화탄소(ppm)'].mean()
                    ee= week['휘발성유기화합물(ppb)'].mean()
                    ff= week['온도(℃)'].mean()
                    gg= week['습도(%)'].mean()
                    hh= week['소음(dB)'].mean()
                    ii= week['통합지수'].mean()
                    jj= week['미세먼지 (원본)'].mean()
                    kk= week['초미세먼지 (원본)'].mean()
            
                    week_f = pd.DataFrame([[aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk]], columns=['DATE',
                                                                                                   '미세먼지(ug/m³)',
                                                                                                   '초미세먼지(ug/m³)',
                                                                                                   '이산화탄소(ppm)',
                                                                                                   '휘발성유기화합물(ppb)',
                                                                                                   '온도(℃)',
                                                                                                   '습도(%)',
                                                                                                   '소음(dB)',
                                                                                                   '통합지수',
                                                                                                   '미세먼지 (원본)',
                                                                                                   '초미세먼지 (원본)'])
            
                    h_mean = pd.concat([h_mean, week_f])
                    c = datetime.timedelta(days=7)
            
                    wedate = str(wedate)
                    wedate = wedate[0:11]
        
   
                file2 = file[:-5]    
                file2 = file2 + '-' + 'week' + '.csv'
                try:
                   if not isRunningFunc() :
                       return
                except : pass   
                h_mean.to_csv(file2 ,index=False, encoding='cp949')
                VB = VB + 1    



            progressbar_indeter.destroy()
            progressbar_indeter = ttk.Progressbar(self.testFame, length=150, maximum=150, value=150)
            progressbar_indeter.pack(pady = 8)
            self.testLabel.configure(text="Done!")
            messagebox.showinfo("Data processing...","Data processing Done") 
            progressbar_indeter.destroy()
            self.testLabel.configure(text="")
            self.threadedButton['state'] = tkinter.NORMAL
            self.fileButton['state'] = tkinter.NORMAL
            self.weekButton['state'] = tkinter.NORMAL



    root = Tk()    
    root.configure(bg='white')
    s = ttk.Style()
    s.theme_use('alt')
    s.configure("yong.Horizontal.TProgressbar", foreground='#458577', background='#458577')
    center_window(root, 600, 320)
    root.resizable(width=False, height=False)
    gui = UnitTestGUI( root )
    root.protocol( "WM_DELETE_WINDOW", gui.close )
    root.mainloop()

if __name__ == "__main__": 
    tkThreadingTest()
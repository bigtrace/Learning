#coding = utf - 8
import time
from datetime import *
import pycurl
import io
import re
import lxml.html
import calendar
import win32com.client as win32
import ctypes


#class definition
class jieqi_class():

    url_1 = 'www.travelchinaguide.com/intro/focus/solar-term.htm'
    url_2 = 'www.chinahighlights.com/festivals/the-24-solar-terms.htm'

    def __init__(self):
        self.Get_jieqi_list()
        self.Check_Today_is_Jieqi()
        self.send_email()



    def Get_jieqi_list(self):
        print("start to gather infomation")
        c = pycurl.Curl()
        #c.setopt(pycurl.PROXY, 'http://192.168.87.15:8080')
        #c.setopt(pycurl.PROXYUSERPWD, 'LL66269:')
        #c.setopt(pycurl.PROXYAUTH, pycurl.HTTPAUTH_NTLM)


        USER_AGENT = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
        c.setopt(c.FOLLOWLOCATION, 1)
        c.setopt(pycurl.VERBOSE, 0)
        c.setopt(pycurl.FAILONERROR, True)
        c.setopt(pycurl.SSL_VERIFYPEER, False)
        c.setopt(pycurl.USERAGENT, USER_AGENT)
        c.setopt(pycurl.URL, self.url_1)
        buffer = io.BytesIO()
        c.setopt(c.WRITEDATA, buffer)
        c.perform()
        
        body = buffer.getvalue().decode('utf-8', 'ignore')
        doc = lxml.html.fromstring(body)
        jieqi_list = doc.xpath("//table[@class='c_table']/tbody/tr")
        self.jieqi_map={}
        for i,each_row in enumerate(jieqi_list):
            if i>0:
                detail=[]
                #print "----"
                each_row_ = each_row.xpath(".//td")
                for each_column in each_row_:
                    detail.append(each_column.text_content())

                self.jieqi_map[i]=detail



        buffer = io.BytesIO()
        c.setopt(pycurl.URL, self.url_2)
        c.setopt(c.WRITEDATA, buffer)
        c.perform()
        c.close()
        body = buffer.getvalue().decode('utf-8', 'ignore')
        doc = lxml.html.fromstring(body)
        jieqi_list_1 = doc.xpath("//table[@class='table']/tbody/tr")
        self.jieqi_explanation_map={}
        for i,each_row in enumerate(jieqi_list_1):
            if i>0:
                more_detail=[]
                #print "----"
                more_detail = each_row.xpath(".//td/p")[3].text_content()


                self.jieqi_explanation_map[i]=more_detail   

        #print self.jieqi_explanation_map



    def Check_Today_is_Jieqi(self):
        self.abbr_to_num = {name: num for num, name in enumerate(calendar.month_abbr) if num}
        
        self.hit = False
        Today= date.today()  #date(2018, 5, 5)
        
        for key,detail in self.jieqi_map.items():
            Month = re.search(r"^(\w{3})", detail[1]).group(1)
            Month_num=self.abbr_to_num[Month]
            day=re.search(r"(\d+)", detail[1]).group(1)

            date_jieqi=date(2018, Month_num, int(day))

            if Today==date_jieqi:  
                print ("Today, a new jieqi begin! --" + detail[0])
                self.hit = True

                if key>2:
                    self.index=key-2
                else:
                    self.index=22+key
                
                self.detail=detail



    def send_email(self):
        if self.hit == True:
            print("sending email!")

            jieqiname=self.detail[0]
            File_name = re.search(r"\((.*)\)", jieqiname).group(1).replace(" ", "").lower()+".jpg"
            Summary = self.detail[2]
            detail_summary= self.jieqi_explanation_map[self.index]

        
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'fiona.ding@tdsecurities.com'
            mail.Subject = jieqiname
            mail.CC = "xingwanlibigtrace@gmail.com"
            mail.BCC = ""
            html_body = "<html><body><p><strong>" + Summary + "</strong></p><p>"+detail_summary+"</p><img src='cid:"+File_name+"' height=1200 width=1920 /></body></html>"

            attachment_path = "H:/Learning/jieqi/photo/"+File_name
            sttachment = mail.Attachments.Add(attachment_path)

            sttachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", File_name )
            mail.HTMLBody = html_body
            mail.send
            #self.set_wallpaper(attachment_path)


    def set_wallpaper(self,attachment_path):
        print("---")
        print(attachment_path)
        output=ctypes.windll.user32.SystemParametersInfoA(20, 0, attachment_path , 3)
        if output ==1 :
            print ("set wallpaper successfully!")

app = jieqi_class()


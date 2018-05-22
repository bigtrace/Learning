# coding=utf-8
import pycurl
import io
import json
from pprint import pprint
import re
import codecs
import win32com.client as win32


def Send_Email(source):


    if source == 'time':
        sort_by = 'latest'
    else:
        sort_by = 'top'

    buffer = io.BytesIO()
    c = pycurl.Curl()
    c.setopt(pycurl.URL, 'https://newsapi.org/v1/articles?source='+ source + '&sortBy=' + sort_by + '&apiKey=d38a68daaacf4e209072898ee50386bb')
    c.setopt(pycurl.SSL_VERIFYPEER,False)
    c.setopt(pycurl.VERBOSE, 0)
    c.setopt(c.WRITEDATA, buffer)

    # c.setopt(c.WRITEDATA, f)

    c.perform()

    body = buffer.getvalue()

    # print body

    data = json.loads(body)

    # pprint(data)

    html_buffer=[]

    header='<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><table  style="border-collapse: separate">   <tr>'

    html_buffer.append(header)

    html_buffer.append('<th> Top  News from ' + source + '  </th>')

    html_buffer.append('</tr>')


    for articles in data['articles']:



        html_buffer.append('<tr> <td> ><a href="' + articles['url'] + '"> ' + '<font  size="6" >' + articles['title'] + '</font ></a></td> </tr>')

        if articles.get('description'):
            html_buffer.append('<tr> <td >' + articles['description'] + '</td> </tr>')
        if articles.get('publishedAt'):
            html_buffer.append('<tr> <td ><font  color="grey"><i>'
                    + articles['publishedAt']
                    + '</i></font></td> </tr>')
        if articles.get('urlToImage'):
            html_buffer.append("<tr> <td>  <img src='" + articles['urlToImage'] + "' height=200 > </td></tr>")
    html_buffer.append('</table>')

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'fiona.ding@tdsecurities.com'
    mail.CC = 'xingwanlibigtrace@gmail.com'
    mail.Subject = 'Bloomberg Daily News report '

    # mail.body = 'Message body'


    mail.HTMLBody = " ".join(html_buffer)
    mail.send


if __name__ == "__main__":

    source_list = ['bloomberg']

    # All available source:  "bloomberg", "reuters","google-news","bloomberg","time"

    for each in source_list:
        Send_Email(each)
#encodig=utf-8
# _*_ coding:utf-8 _*_
import os  
import sys
import smtplib  
import mimetypes  
import pandas as pd
import codecs
import chardet
import string
import argparse
import time
from email.MIMEMultipart import MIMEMultipart  
from email.MIMEBase import MIMEBase  
from email.MIMEText import MIMEText  
from email.MIMEAudio import MIMEAudio  
from email.MIMEImage import MIMEImage  
from email import encoders
from email.utils import parseaddr, formataddr
from email.header import Header
from email.Encoders import encode_base64
from email.utils import COMMASPACE

########################################################################################
##　この脚本を使用する前に以下の内容をどうかご記入すること：
count=206             #本期简报的期数
account_number=0   #这次发送简报最开始使用第几个邮箱(我是16)(从1开始)
donated=0       #这次发送简报是有捐赠版(1)还是无捐赠版(0)
tired=0       #いつか次のメールアドレスに交換するをコントロールする量
##　もし本日このメールアドレスを使用したことがございませんとしたら，tiredを０のままにして頂戴
##　でなければ、回数を５０で割る数字を記入してください
########################################################################################

parser = argparse.ArgumentParser()
parser.add_argument("-cm", "--checkmyself", help="send a check mail to yourself", action="store_true")
parser.add_argument("-cl", "--checkleader", help="send a test mail to your leader and yourself", action="store_true")
parser.add_argument("-sa", "--sendall", help="send a mail to your responisible part", action="store_true")
parser.add_argument("-t", "--testmail", help="send a mail to your test table", action="store_true")
parser.add_argument("-num", "--birefnumber", help="the number of working brief you gonna send this time, for example, 206",type=int)
args = parser.parse_args()
if args.birefnumber:
    count = args.birefnumber
print count

def get_msg():
    # メールアドレス情報取得
    global gmailUser,gmailPassword,emails
    global subject,content,pdf_path
    global md
    print 'Loading data ...'
    if args.testmail:
        print 'testing'
        md = pd.read_excel(u'./index/test_mails.xlsx',header=None)#test_mails
        emails = [x.encode('utf-8') for x in md[2] if pd.isnull(x)==False]
    else:
        md = pd.read_excel(u'./index/mydutypart.xlsx',header=None)#test_mails
        emails = [x.encode('utf-8') for x in md[2] if pd.isnull(x)==False and type(x)!=float]        
    we = pd.read_excel(u'./index/work_emails.xlsx')
    gmailUser = list(we[u'账号'])[account_number]
    gmailPassword = list(we[u'密码'])[account_number]

    # ファイルメッセージ情報取得
    if count%2==0:
        filepathprefix = u"./source/"+str(count)+u'/无捐款/'
    else:
        filepathprefix = u"./source/"+str(count)+'/'
    for fname in os.listdir(filepathprefix):
        if os.path.splitext(fname)[1] == '.pdf':
            pdf_name = fname
        elif os.path.splitext(fname)[1] == '.txt':
            txt_name = fname
    pdf_path,txt_path = filepathprefix+pdf_name,filepathprefix+txt_name
    subject = u"Dian团队工作简报第"+str(count)+u"期"

    with open(txt_path, 'rb') as fp:
        file_data = fp.read()
        result = chardet.detect(file_data)
        content = file_data.decode(encoding=result['encoding'])

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'gb2312').encode(), \
        addr.encode('gb2312') if isinstance(addr, unicode) else addr))
 
def prepmsg(subject, text, *attachmentFilePaths):
    global msg
    recipient = []  
    # ルートコンテナをセットします
    msg = MIMEMultipart()  
    msg['From'] = _format_addr(u'Newsletter Dian <%s>' % gmailUser) 
    msg['To'] = COMMASPACE.join(recipient)
    msg['Subject'] = Header(subject, 'gb2312').encode()  
    msg.attach(MIMEText(content, 'plain', 'gb2312'))  
    for attachmentFilePath in attachmentFilePaths:  
        msg.attach(getAttachment(attachmentFilePath)) 
    
def sendMail(other):   
    global msg
    global mailServer
    recipient = []
    msg['Bcc'] = COMMASPACE.join(other)
    print '='*60,"\n現在では発信することを実行しているアカウントは：",gmailUser
    mailServer.sendmail(gmailUser, recipient+other, msg.as_string())  
    print "Sent email to ", other
   
def getAttachment(attachmentFilePath):  
    contentType, encoding = mimetypes.guess_type(attachmentFilePath)  
    if contentType is None or encoding is not None:  
        contentType = 'application/octet-stream' 
    mainType, subType = contentType.split('/', 1)  
    file = open(attachmentFilePath, 'rb')  
   
    if mainType == 'text':  
        attachment = MIMEText(file.read())  
    elif mainType == 'message':  
        attachment = email.message_from_file(file)  
    elif mainType == 'image':  
        attachment = MIMEImage(file.read(),_subType=subType)  
    elif mainType == 'audio':  
        attachment = MIMEAudio(file.read(),_subType=subType)  
    else:  
        attachment = MIMEBase(mainType, subType)  
    attachment.set_payload(file.read())  
    encode_base64(attachment)    
    file.close()  
    ## 添付ファイルのヘッダをセットします
    attachment.add_header('Content-Disposition', 'attachment',   filename=os.path.basename(attachmentFilePath))  
    return attachment  

def OneUsrSendMail():
    global tired
    global mailServer
    start = time.time()
    allsent=0
    rested=len(emails)
    mailServer = smtplib.SMTP('mail.hust.edu.cn',25)
    mailServer.ehlo()  
    mailServer.starttls()  
    mailServer.ehlo()  
    mailServer.login(gmailUser, gmailPassword)  
    prepmsg(subject, content, pdf_path)
    while(rested!=0):
        if(rested>=50):
            sendMail(emails[allsent:(allsent+50)])
            allsent=allsent+50
            print ("この 50 通のメールはこのメールアドレスより出します：%s ，今既に %d 通のメールを送っています，残りの部分は直ぐに出しますので，少々お待ちください...\n" % (gmailUser.encode('UTF-8'),allsent)).decode('UTF-8')
            rested=rested-50
        else:
            sendMail(emails[allsent:])
            allsent=len(emails)
            print ("この %d 通のメールはこのメールアドレスより出します：%s ，もはや全部の %d 通のメールを送りましたので，ご利用ありがとうございます:)" % ((allsent%50),gmailUser.encode('UTF-8'),allsent)).decode('UTF-8')
            mailServer.close()  
            rested=0
        print '='*60
        if(tired>=50):
            resetMail()
        else:
            tired=tired+1
    end = time.time()
    print "\nこの度、本プログラムの実行時間は合わせて： ",end-start," 秒です。\n",'='*60

def resetMail():
    global account_number
    global gmailUser
    global gmailPassword
    global tired
    global mailServer
    account_number=account_number+1
    gmailUser = list(we[u'账号'])[account_number]
    gmailPassword = list(we[u'密码'])[account_number]
    mailServer = smtplib.SMTP('mail.hust.edu.cn',25)#'smtp.gmail.com'　'smtp.qq.com' 587
    mailServer.ehlo()  
    mailServer.starttls()  
    mailServer.ehlo()  
    mailServer.login(gmailUser, gmailPassword)  
    tired=0


get_msg()
if args.checkmyself:
    emails = ['786671043@qq.com']#,'dengpanstudio@outlook.com']
    OneUsrSendMail()
    tired=0

if args.checkleader:
    emails = ['786671043@qq.com','dengpanstudio@outlook.com']
    OneUsrSendMail()
    tired=0

if args.sendall:
    emails = [x.encode('utf-8') for x in md[2] if pd.isnull(x)==False and type(x)!=float]
    OneUsrSendMail()
    tired=0

if args.testmail:
    emails = [x.encode('utf-8') for x in md[2] if pd.isnull(x)==False and type(x)!=float]
    OneUsrSendMail()
    tired=0


if  len(sys.argv) == 1:
    print 'please use -h to get the help!'
    
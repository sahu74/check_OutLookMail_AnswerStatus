import os
import re
import time
import win32com.client
from  datetime import datetime,timedelta
from reportMail import mailData, HtmlMailReport

import configparser





class CheckMailer:

    def __init__(self,daysOfReport:int=0,folderindex=6):
        self.outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(folderindex)
        self.sendBox = self.outlook.GetDefaultFolder(5)
        self.messages = self.inbox.Items
        self.sendMessages = self.sendBox.Items
        self.messages.Sort("ReceivedTime", True)
        self.sendMessages.Sort("ReceivedTime", True)
        self.daysOfReport=daysOfReport
        self.totalHours = time.localtime().tm_hour
        self.totalHours += 24



        if daysOfReport !=0:
            received_dt = datetime.now() - timedelta(days=self.daysOfReport)
            received_dt = received_dt.strftime('%Y/%m/%d %H:%M %p')
            self.messages = self.messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
            self.sendMessages = self.sendMessages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        else:
            received_dt = datetime.now() - timedelta(hours=self.totalHours)
            received_dt = received_dt.strftime('%Y/%m/%d %H:%M %p')
            self.messages = self.messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
            self.sendMessages = self.sendMessages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    def getInBoxMesages(self):
        return self.messages

    def msgFoundInSentFolde(self,inboxMsgSub):
        for i in self.sendMessages:
            sub=i.Subject
            if sub.find(inboxMsgSub) != -1:
                sentTime=i.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                if sub.find("FW:") != -1:
                    return "YESFW::" + str(sentTime)
                elif sub.find("RE:") != -1:
                    return "YESRE::" + str(sentTime)
                else:
                    return "YES::" + str(sentTime)

        return "NO"

    def testSendMsg(self,mailId='rsukanya@netapp.com'):
        #Firstly, your code will fail if you have an item other than MailItem in the folder, such as ReportItem, MeetingItem, etc.
        # You need to check the Class property. Secondly, you need to check the sender email address type and use the
        # SenderEmailAddress only for the "SMTP" address type
        foundMailId=''
        for i in self.messages:
            if i.Class == 43:
                if i.SenderEmailType == "EX":
                    if i.Sender.GetExchangeUser() is not None:
                        print(i.Sender.GetExchangeUser().PrimarySmtpAddress)
                        if i.Sender.GetExchangeUser().PrimarySmtpAddress == mailId:
                            return True
                    else:
                        print(i.Sender.GetExchangeDistributionList().PrimarySmtpAddress)
                        if i.Sender.GetExchangeDistributionList().PrimarySmtpAddress == mailId:
                            return True
                else:
                    print(i.SenderEmailAddress)
                    if str(i.SenderEmailAddress) == mailId:
                        return True



    def getInBox(self):
        return self.inbox

    def getInboxItem(self):
        return self.inbox.Items

    def showInboxFolders(self):
        inBoxFolder=self.getInBox()
        [print(ib) for ib in inBoxFolder.Folders]

    def showTotalMessages(self):
        print("Total Messages in inbox Folders ",len(self.messages))
        self.showMesages()

    def showMesages(self):
        msg= self.messages.GetFirst()

        msg = self.messages.GetFirst()
        while msg:
            try:
                print(msg.ReceivedTime)
                msg = self.messages.GetNext()
            except:
                continue

    def isSendersEmailIdMatching(self,msg,mailId):
        #Firstly, your code will fail if you have an item other than MailItem in the folder, such as ReportItem, MeetingItem, etc.
        # You need to check the Class property. Secondly, you need to check the sender email address type and use the
        # SenderEmailAddress only for the "SMTP" address type
        foundMailId=''
        if msg.Class == 43:
            if msg.SenderEmailType == "EX":
                if msg.Sender.GetExchangeUser() is not None:
                    if msg.Sender.GetExchangeUser().PrimarySmtpAddress == mailId:
                        print("1", msg.Sender.GetExchangeUser().PrimarySmtpAddress)
                        return True
                else:
                    if msg.Sender.GetExchangeDistributionList().PrimarySmtpAddress == mailId:
                        print("2", msg.Sender.GetExchangeDistributionList().PrimarySmtpAddress)
                        return True
            else:
                if str(msg.SenderEmailAddress) == mailId:
                    print("3",msg.SenderEmailAddress)
                    return True

        return False

    #chk, chk.totalHours, chk.daysOfReport
    def constructReportData(self):
        ll =[]
        i:int=1
        #os.getcwd() + '\\mail\\'
        config = configparser.ConfigParser()
        config.read('project.cfg')
        # mailId = config.get('HIGHLIGHT_MAIL', 'mail_id')
        mailId=config['HIGHLIGHT_MAIL']['mail_id']
        highLightColor=config['HIGHLIGHT_MAIL']['color']
        os.makedirs("email",mode=0o777,exist_ok=True)
        mesages=self.getInBoxMesages()
        for msg in mesages:
            md = mailData()
            i+=1
            md.Id=str(i)
            name = str(msg.Subject)
            name = re.sub('[^A-Za-z0-9]+', '', name) + '.msg'
            if self.isSendersEmailIdMatching(msg,mailId):
                s = "style=\"" + "color: red\""
                subStr = "<p><a href=" + "email\\" + name + " " + s + ">"
                subStr += str(msg.Subject)
                subStr += "</a></p>"
                print(" hara mohan")
            elif str(msg.Subject).find("FORM 16") != -1 or str(msg.Subject).find("Action Required") != -1 or \
                    str(msg.Subject).find("Case#") != -1 :
                # s="style=\"" + "color: #73E600\""
                s="style=\"" + "color: #CCCC00\""
                subStr = "<p><a href=" + "email\\" + name +" "+ s+">"
                subStr += str(msg.Subject)
                subStr += "</a></p>"
            else:
                subStr = "<p><a href=" + "email\\" + name +">"
                subStr += str(msg.Subject)
                subStr += "</a></p>"

            md.MailSub=subStr
            md.MailRead="NO" if msg.UnRead else "YES"
            md.MailRecvdTm=msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")

            md.MailReply=self.msgFoundInSentFolde(msg.Subject)

            # md.MailRemark= "Please reply if required" if md.MailReply=="NO" else "You have replied to this mail"
            getReplyTime= md.MailReply.split("::")

            try:
                timeStr=str(getReplyTime[1])
                print(" INDIA : ",timeStr)
            except IndexError:
                timeStr=''
                pass
            if md.MailReply.find("YESFW::")!= -1:
                md.MailRemark ="You have forwarded this mail on "+timeStr if len(timeStr)> 0 else ''
            elif md.MailReply.find("YESRE") != -1:
                # md.MailRemark="You have replied to this mail"
                ss="<p style=\"" + "color:#87F717\"" +">"
                mystr =  ss
                mystr += "You have replied to this mail on "+timeStr if len(timeStr)> 0 else ''+"</p>"
                md.MailRemark=mystr
                print(md.MailRemark)
            elif md.MailReply.find("YES::") !=-1:
                md.MailRemark = "you have taken action on this mail on " +timeStr if len(timeStr)> 0 else ''
            else:
                md.MailRemark ="Please reply if required"

            md.MailReply = "YES" if md.MailReply.find("YES") != -1  else "NO"
            try:
                msg.SaveAs(os.getcwd() + '\\email\\' + name)
            except Exception as e:
                print("error when saving the attachment:" + str(e))


            ll.append(md)

        hTblObj=HtmlMailReport(ll,self.totalHours,self.daysOfReport)
        hTblObj.writeReport()
        hTblObj.openInBrowser()


def main():
    chk=CheckMailer()
    # chk.testSendMsg()
    chk.constructReportData()
    # chk.showInboxFolders()
    # chk.showTotalMessages()
    # chk.showMesages()
    # chk.testSendMsg()






if __name__=="__main__":
    main()
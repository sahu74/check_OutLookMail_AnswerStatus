import webbrowser
import configparser

class mailData:
    def __init__(self,id:int=None,sub:str=None,recvdTm:str=None,sentM:str=None,read:str=None,reply:str=None,remarks:str=None):
        self._Id=id
        self._MailSub=sub
        self._MailRecvdTm=recvdTm
        self._MailSentTM=sentM
        self._MailRead=read
        self._MailReply=reply
        self._MailRemark=remarks


    @property
    def Id(self):
        return self._Id
    @Id.setter
    def Id(self,id):
        self._Id=id
    @property
    def MailSub(self):
        return self._MailSub
    @MailSub.setter
    def MailSub(self,ms):
        self._MailSub=ms
    @property
    def MailRecvdTm(self):
        return self._MailRecvdTm

    @MailRecvdTm.setter
    def MailRecvdTm(self,mrt):
        self._MailRecvdTm=mrt

    @property
    def MailSentTM(self):
        return self._MailRecvdTm

    @MailSentTM.setter
    def MailSentTM(self,ms):
        self._MailSent=ms
    @property
    def MailRead(self):
        return self._MailRead

    @MailRead.setter
    def MailRead(self,mr):
        self._MailRead=mr

    @property
    def MailReply(self):
        return self._MailReply

    @MailReply.setter
    def MailReply(self, mr):
        self._MailReply = mr

    @property
    def MailRemark(self):
        return self._MailRemark

    @MailRemark.setter
    def MailRemark(self, rmk):
        self._MailRemark = rmk

class HtmlMailReport:
    def __init__(self,lData:list,hrs=0,daysOfReport=0):
        self.last_hrs = str(hrs) + " Hours"  if daysOfReport == 0 else str(daysOfReport) + " Days"
        self.daysOfReport=daysOfReport
        self.htmlFileName='MailStatusReport.html'
        self.htmlMailData=''
        self.listMailData=lData
        self.htmlFile=open(self.htmlFileName,'w')
        self.htmlMessage= '<!DOCTYPE html>'+"\n"
        self.fillReportTableData()
        self.createMailStatusReportTable()
        self.column_num:int=0




    def fillReportTableData(self):
        self.config = configparser.ConfigParser()
        self.config.read('project.cfg')
        self.column_num=int(self.config['COLUMN_HIDE']['column_num'])
        for i in self.listMailData:
            if self.column_num != 1:
                self.htmlMailData +="<td align=center>"+str(i.Id)+"</td>"+"\n"
            if self.column_num != 2:
                self.htmlMailData +="<td align=center>"+i.MailSub+"</td>"+"\n"
            if self.column_num != 3:
                self.htmlMailData +="<td align=center>"+str(i.MailRecvdTm).strip()+"</td>"+"\n"
            if self.column_num != 4:
                self.htmlMailData +="<td align=center>"+str(i.MailSentTM).strip()+"</td>"+"\n"
            if self.column_num != 5:
                self.htmlMailData +="<td align=center>"+i.MailRead+"</td>"+"\n"
            if self.column_num != 6:
                self.htmlMailData +="<td align=center>"+i.MailReply+"</td>"+"\n"
            if self.column_num != 7:
                self.htmlMailData +="<td align=center>"+i.MailRemark+"</td>"+"\n"
            self.htmlMailData +="</tr>"+"\n"

    def createMailStatusReportTable(self):
        self.config = configparser.ConfigParser()
        self.config.read('project.cfg')
        self.column_num=int(self.config['COLUMN_HIDE']['column_num'])

        self.htmlMessage += '<html>' + "\n"
        self.htmlMessage += '<head>' + "\n"
        self.htmlMessage += '<style>' + "\n"
        self.htmlMessage += 'table,th,td' + "\n"
        self.htmlMessage += '{' + "\n"
        self.htmlMessage += 'border:1px solid black ;' + "\n"
        self.htmlMessage += 'border-collapse:collapse;' + "\n"
        self.htmlMessage += '}' + "\n"
        self.htmlMessage += '</style>' + "\n"
        self.htmlMessage += '</head>' + "\n"
        self.htmlMessage +="<p style="+"\n" +"color:blue;text-align:center" +">"

        self.htmlMessage +="<b> "+"\n"
        self.htmlMessage +="Your mails for last " +str(self.last_hrs) + ". Please have a look and respond accordingly!!!"

        self.htmlMessage += '<body>' + "\n"
        self.htmlMessage +="\n \n"
        # // DISPLAY  TABLE
        self.htmlMessage += '<table>' + "\n"
        self.htmlMessage += '<tr bgcolor="#99C4E7">' + "\n"
        if self.column_num !=1:
            self.htmlMessage += '<td align=center>&nbsp;&nbsp;#ID&nbsp;&nbsp;</td>'+"\n"
        if self.column_num != 2:
            self.htmlMessage += '<td align=center>&nbsp;MAIL_SUBJECT&nbsp;</td>'+"\n"
        if self.column_num != 3:
            self.htmlMessage += '<td align=center>&nbsp;RECEIVED_TIME&nbsp;</td>'+"\n"
        if self.column_num != 4:
            self.htmlMessage += '<td align=center>&nbsp;REPLIED_TIME&nbsp;</td>'+"\n"
        if self.column_num != 5:
            self.htmlMessage += '<td align=center>&nbsp;MAIL_READ&nbsp;</td>'+"\n"
        if self.column_num != 6:
            self.htmlMessage += '<td align=center>&nbsp;REPLIED&nbsp;</td>'+"\n"
        if self.column_num != 7:
            self.htmlMessage += '<td align=center>&nbsp;REMARKS&nbsp;</td>'+"\n"

        self.htmlMessage += '</tr>' + "\n"
        self.htmlMessage += self.htmlMailData
        self.htmlMessage += '</table>' + "\n"
        self.htmlMessage += '</body>' + "\n"
        self.htmlMessage += '</html>' + "\n"

        # print(self.htmlMessage)

    def writeReport(self):
        self.htmlFile.write(self.htmlMessage)
        self.htmlFile.close()
        #
    def openInBrowser(self):
        webbrowser.open_new_tab(self.htmlFileName)





# #self,id:int,sub:str,recvd:str,sentM:str,reply:str,remarks:str):
# mData= mailData(1,"Action Required: EINC12029880 in Awaiting User response","NO","Yes","NO","yes", "Please Reply")
# mData1= mailData(2,"[confluence] CPE > mNode api endpoints","NO","Yes","NO","NO", "You have responded")
# mData2= mailData(3,"The importance of personalized experiences in banks","Yes","NO","Yes","no", "sent new mail")

# l=[mData,mData1,mData2]
#
# hTblObj=HtmlMailReport(l)
#
# # hTblObj.createMailStatusReportTable()
# hTblObj.writeReport()
# hTblObj.openInBrowser()


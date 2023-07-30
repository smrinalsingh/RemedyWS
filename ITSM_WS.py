import requests as _requests
import re as _re
from getpass import getpass as _gp
#import win32com.client as _win32
import codecs
import configparser


"""This module would help fetch tickets using the ITSM webservices"""

filePath = 'exeConfig.ini'
config = configparser.ConfigParser(delimiters=('|'))
config.read(filePath)
#if 'user' not in config['DEFAULT'] or 'password' not in config['DEFAULT'] or 'url' not in config['DEFAULT'] or 'webservname' not in config['DEFAULT']:

_user = config['DEFAULT']['user']
_password = config['DEFAULT']['password']
_url = config['DEFAULT']['url']
_webservname = config['DEFAULT']['webservname']

    

#_outlook = _win32.Dispatch('outlook.application')
_headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': ''}

#_user=str(input("Username: "))
#_password= _gp("Password: ")


class getList:
    
    _headers['SOAPAction'] = '%s/HelpDesk_SearchList_Service'%(_webservname)
    def __init__(self, status = "Assigned", assignedGroup = "GSD Automation", optcat1="Personal Computing", optcat2="Mobile Pass", optcat3="Request for MobilePASS", startRec = "?", maxLimit = "?"):
        self._status = status
        self._assignedGroup = assignedGroup
        self._optcat1 = optcat1
        self._optcat2 = optcat2
        self._optcat3 = optcat3
        self.startRec = startRec
        self.maxLimit = maxLimit


    def get(self):
        global _user
        global _password
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_GSD">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_SearchList_Service>
         <urn:Qualification>?</urn:Qualification>
         <urn:startRecord>"""
        bg4 = """</urn:startRecord>
         <urn:maxLimit>"""
        bg5 = """</urn:maxLimit>
         <urn:Status>"""
        bg6 = """</urn:Status>
         <urn:AssignedGroup>"""
        bg7 = """</urn:AssignedGroup>
         <urn:OptCat1>"""
        bg8 = """</urn:OptCat1>
         <urn:OptCat2>"""
        bg9 = """</urn:OptCat2>
         <urn:OptCat3>"""
        bg10 = """</urn:OptCat3>

      </urn:HelpDesk_SearchList_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        bodyGet = (bg1+_user + bg2 + _password + bg3 + self.startRec + bg4 + self.maxLimit + bg5 + self._status + bg6 + self._assignedGroup\
                   + bg7 + self._optcat1 +bg8 + self._optcat2 + bg9 + self._optcat3 + bg10)
        getResp = _requests.post(_url, data=bodyGet, headers=_headers)
        self.getResp = getResp
        return getResp


    def list(self):
        resp = (self.getResp).text
        _ticNums = _re.findall('<ns0:IncidentNumber>([A-Za-z0-9 ]*)</ns0:IncidentNumber>', resp)
        return _ticNums


    def __repr__(self):
        return ("<<HelpDesk_SearchList_Service>>")





class getListNoCat:
    #_headers['SOAPAction'] = 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_SearchList_NoCategory_Service'
    _headers['SOAPAction'] = '%s/HelpDesk_SearchList_NoCategory_Service'%(_webservname)
    
    def __init__(self, assignedGroup="GSD Automation", status="Assigned", startRec="?", maxLimit="?"):
        self._group = assignedGroup
        self._status = status
        self._startRec = startRec
        self._maxLimit = maxLimit


    def get(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_SearchList_NoCategory_Service>
         <urn:Qualification>?</urn:Qualification>
         <urn:startRecord>"""
        bg4 = """</urn:startRecord>
         <urn:maxLimit>"""
        bg5 = """</urn:maxLimit>
         <urn:Status>"""
        bg6 = """</urn:Status>
         <urn:AssignedGroup>"""
        bg7 = """</urn:AssignedGroup>
      </urn:HelpDesk_SearchList_NoCategory_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        bodyGet = (bg1 + _user + bg2 + _password + bg3 + self._startRec + bg4 + self._maxLimit + bg5 + self._status + bg6 + self._group + bg7)
        self.getResp = _requests.post(_url, data=bodyGet, headers=_headers)
        return self.getResp


    def list(self):
        resp = (self.getResp).text
        _ticNums = _re.findall('<ns0:IncidentNumber>([A-Za-z0-9 ]*)</ns0:IncidentNumber>', resp)
        return _ticNums


    def __repr__(self):
        return ("<<HelpDesk_SearchList_NoCat_Service>>")

        
class modTicket:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_StatusMod_Service'}
    def __init__(self, ticNum, status="Assigned", statReason = "", resComment = ""):
        self._ticNum = ticNum
        self._status = status
        self._statReason = statReason
        self._resComment = resComment
        
        
    def mod(self):
        
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_GSD_1">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_StatusMod_Service>
         <!--Optional:-->
         <urn:IncidentNumber>"""
        bg4 = """</urn:IncidentNumber>
         <!--Optional:-->
         <urn:Status>"""
        bg5 = """</urn:Status>
         <!--Optional:-->
         <urn:StatusReason>"""
        bg6 = """</urn:StatusReason>
         <!--Optional:-->
         <urn:ResComment>"""
        bg7 = """</urn:ResComment>
      </urn:HelpDesk_StatusMod_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        _bodyMod = (bg1+_user+bg2+_password+bg3+self._ticNum+bg4+self._status+bg5+self._statReason+bg6+self._resComment+bg7)
        getResp = _requests.post(_url, data=_bodyMod, headers=_headers)

        return getResp


    def __repr__(self):
        return ("<<HelpDesk_ModTicket_Service>>")
    


class getTicket:
    """This class helps extract the ticket data returned and passes it back to us. The data
        is in list format and index is as follows:
        1. Summary
        2. Notes
        3. Status
        4. Assigned Group
        5. Assigned Group ID
        6. OptCat1
        7. OptCat2
        8. OptCat3
        9. Assignee
        10. User's ID
        11. Status Reason
        12. Assignee Login ID
        13. Resolution Comment

        Note: List index starts from 0. The above mentioned is generic readable format."""

    _headers['SOAPAction'] = '%s/HelpDesk_SearchInc_Service'%(_webservname)
    #_headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_SearchInc_Service'}

    def __init__(self, ticNum):
        self._ticNum = ticNum

    def get(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_GSD_1">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_SearchInc_Service>
         <urn:IncidentNumber>"""
        bg4 = """</urn:IncidentNumber>
      </urn:HelpDesk_SearchInc_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        _bodyGet = (bg1+_user+bg2+_password+bg3+self._ticNum+bg4)
        self.getResp = _requests.post(_url, data=_bodyGet, headers=_headers)
        
        return self.getResp


    def searchREO(self, terms):
        foundData = {}
        for each in terms:
            resp = self.getResp.text
            blank = _re.search('nothing( )much', "nothing much")
            foundeachObj = _re.search('<ns0:' + _re.escape(each) + '>(.*)</ns0:' + _re.escape(each) + '>', resp)
            foundeachObj = foundeachObj if foundeachObj != None else blank
            foundVal = foundeachObj.groups(0)[0]
            foundData[each] = foundVal
        return foundData


    def list(self):
        resp = self.getResp.text
        blank = _re.search('nothing( )much', "nothing much")
        searchTerms = ["Summary", "Notes", "Status", "IncidentNum", "AssignedGroup", "AssignedGroupID", "OptCat_1", "OptCat_2", "OptCat_3", "Assignee", "CustomerID", "StatusReason", "AssigneeLoginID", "ResComment", "Site", "Region", "ResportedSource", "GroupTransfers", "SubmitDate", "Organization", "VIP", "TargetDate", "ClosedDate", "AssignedOrganization", "Priority", "LastModDate", "LastModBy", "LastResolveDate", "IncidentType", "Submitter", "ReportedDate"]
        res = self.searchREO(searchTerms)
        return res

    def __repr__(self):
        return ("<<HelpDesk_GetTicket_Service>>")
    


class reassignTic:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_Reassignment_Service'}
    def __init__(self, ticNum, status="Assigned", reassignTo="GSD Automation", reassignReason=""):
        self.ticNum = ticNum
        self.status = status
        self.reassignTo = reassignTo
        self.reassignReason = reassignReason

    def reassign(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_Reassignment_Service>
         <!--Optional:-->
         <urn:Status>"""
        bg4 = """</urn:Status>
         <urn:Incident_Number>"""
        bg5 = """</urn:Incident_Number>
         <!--Optional:-->
         <urn:Assigned_Support_Group>"""
        bg6 = """</urn:Assigned_Support_Group>
         <!--Optional:-->
         <urn:Reassignment_Reason>"""
        bg7 = """</urn:Reassignment_Reason>
         <!--Optional:-->
         <urn:Assignee_Login_ID></urn:Assignee_Login_ID>
         <!--Optional:-->
         <urn:Assigned_Support_Company></urn:Assigned_Support_Company>
         <!--Optional:-->
         <urn:Assigned_Support_Organization></urn:Assigned_Support_Organization>
         <!--Optional:-->
         <urn:Assignee></urn:Assignee>
      </urn:HelpDesk_Reassignment_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        _bodyMod = (bg1 + _user + bg2 + _password + bg3 + self.status + bg4 + self.ticNum +
                   bg5 + self.reassignTo + bg6 + self.reassignReason + bg7)
        getResp = _requests.post(_url, data=_bodyMod, headers=_headers)
        return getResp


    def __repr__(self):
        return ("<<HelpDesk_ReassignTicket_Service>>")

    


class getWorkNotes:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_GetWorkInfoList_Service'}
    def __init__(self, ticNum):
        self.ticNum = ticNum

    def get(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_GetWorkInfoList_Service>
         <urn:Incident_Number>"""
        bg4 = """</urn:Incident_Number>
      </urn:HelpDesk_GetWorkInfoList_Service>
   </soapenv:Body>
</soapenv:Envelope>"""

        _bodyMod = (bg1 + _user + bg2 + _password + bg3 + self.ticNum + bg4)
        getResp = _requests.post(_url, data=_bodyMod, headers=_headers)
        return getResp

    def __repr__(self):
        return ("<<HelpDesk_GetWorkNotes_Service>>")



class addWorkNotes:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2/HelpDesk_AddWorkInfo_Service'}

    def __init__(self, ticNum, fileName="?", filePath="?", workinfoType="General Information", workinfo=""):
        self._ticNum = ticNum
        self._fileName = fileName
        self._filePath = filePath
        self._wiType = workinfoType
        self._workinfo = workinfo
        self._workinfoSummary = "WI Summary"


    def base64encode(self):
        if self._filePath == "?" or self._filePath == "":
            return ""
        else:
            with open(self._filePath, "rb") as fh:
                return (codecs.encode(fh.read(), encoding="base64")).decode()

            
    def addNote(self):
        self._base64encoded = self.base64encode()
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_ServiceDesk2">
   <soapenv:Header>
      <urn:AuthenticationInfo>
         <urn:userName>"""
        bg2 = """</urn:userName>
         <urn:password>"""
        bg3 = """</urn:password>
         <!--Optional:-->
         <urn:authentication>?</urn:authentication>
         <!--Optional:-->
         <urn:locale>?</urn:locale>
         <!--Optional:-->
         <urn:timeZone>?</urn:timeZone>
      </urn:AuthenticationInfo>
   </soapenv:Header>
   <soapenv:Body>
      <urn:HelpDesk_AddWorkInfo_Service>
         <!--Optional:-->
         <urn:WorkInfo>"""
        bg4 = """</urn:WorkInfo>
         <!--Optional:-->
         <urn:IncidentNumber>"""
        bg5 = """</urn:IncidentNumber>
         <!--Optional:-->
         <urn:WorkInfoSummary>"""
        bg6 = """</urn:WorkInfoSummary>
         <!--Optional:-->
         <urn:WorkInfoType>"""
        bg7 = """</urn:WorkInfoType>
         <!--Optional:-->
         <urn:ViewAccess>Public</urn:ViewAccess>
         <!--Optional:-->
         <urn:Locked>Yes</urn:Locked>
         <!--Optional:-->
         <urn:AttachmentName>"""
        bg8 = """</urn:AttachmentName>
         <urn:AttachmentData>"""
        bg9 = """</urn:AttachmentData>

         <!--Optional:--></urn:HelpDesk_AddWorkInfo_Service>
   </soapenv:Body>
</soapenv:Envelope>"""


        _bodyMod = (bg1 + _user + bg2 + _password + bg3 + self._workinfo + bg4 + self._ticNum + bg5 + self._workinfoSummary + bg6 + self._wiType + bg7 + self._fileName
                    + bg8 + self._base64encoded + bg9)
        
        getResp = _requests.post(_url, data=_bodyMod, headers=_headers)

        return getResp


    def __repr__(self):
        return ("<<HelpDesk_AddWorkNotes_Service>>")
    

d = """
class mail:
    def __init__(self, to="", cc="",  subject="", body="", attachment=False, html=None):
        self.To = to
        self.Subject = subject
        self.body = body
        self.cc =  cc
        self.attach = attachment
        self.html = html

    def send(self):
        mail = _outlook.CreateItem(0)
        mail.To = self.To
        mail.Subject = self.Subject
        mail.Body = self.body
        mail.cc = self.cc
        if self.attach == True:
            mail.Attachments.Add(self.attach)
        else:
            pass
        
        mail.Send()
"""

#getTics = getList()
#print ((getTics.get()).text)

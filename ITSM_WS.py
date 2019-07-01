import requests as _requests
import re as _re
from getpass import getpass as _gp
import win32com.client as _win32
import codecs
import sys


"""This module would help fetch tickets using the ITSM webservices"""

_outlook = _win32.Dispatch('outlook.application')
_url = "http://gditmutwswv51p.corp.capgemini.com:8080/arsys/services/ARService?server=gditmutapwv51p&webService=CAP:HPD_IncidentInterface_EUS_Automation_1"

_user=str(input("Username: "))
_password= _gp("Password: ")



class getList:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_SearchInc_Service'}
    def __init__(self, status = "Assigned", assignedGroup = "GSD Automation", optcat1="Personal Computing", optcat2="Mobile Pass", optcat3="Request for MobilePASS", startRec = "?", maxLimit = "?"):
        self._status = status
        self._assignedGroup = assignedGroup
        self._optcat1 = optcat1
        self._optcat2 = optcat2
        self._optcat3 = optcat3
        self.startRec = startRec
        self.maxLimit = maxLimit
        
    def get(self):
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
        bg4 = """?</urn:startRecord>
         <urn:maxLimit>"""
        bg5 = """?</urn:maxLimit>
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
        return (_ticNums, len(_ticNums))

    def __repr__(self):
        return ("<<HelpDesk_SearchList_Service>>")


class getListNoCat:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_SearchList_NoCategory_Service'}
    def __init__(self, assignedGroup="GSD Automation", status="Assigned", startRec="?", maxLimit="?"):
        self._group = assignedGroup
        self._status = status
        self._startRec = startRec
        self._maxLimit = maxLimit

    def get(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_Automation_1">
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
        return (_ticNums, len(_ticNums))

    def __repr__(self):
        return ("<<HelpDesk_SearchList_NoCategory_Service>>") 



class getTicket:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_SearchInc_Service'}
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

    def list(self):
        resp = self.getResp.text
        blank = _re.search('nothing( )much', "nothing much")
        

        _summary = _re.search('<ns0:Summary>(.*)</ns0:Summary>', resp)
        _summary = _summary if _summary != None else blank
        _notes = _re.search('<ns0:Notes>(.*)</ns0:Notes>', resp)
        _notes = _notes if _notes != None else blank
        _status = _re.search('<ns0:Status>(.*)</ns0:Status>', resp)
        _status = _status if _status != None else blank
        _assignedGroup = _re.search('<ns0:AssignedGroup>(.*)</ns0:AssignedGroup>', resp)
        _assignedGroup = _assignedGroup if _assignedGroup != None else blank
        _assignedGroupID = _re.search('<ns0:AssignedGroupID>(.*)</ns0:AssignedGroupID>', resp)
        _assignedGroupID = _assignedGroupID if _assignedGroupID != None else blank
        _optcat1 = _re.search('<ns0:OptCat_1>(.*)</ns0:OptCat_1>', resp)
        _optcat1 = _optcat1 if _optcat1 != None else blank
        _optcat2 = _re.search('<ns0:OptCat_2>(.*)</ns0:OptCat_2>', resp)
        _optcat2 = _optcat2 if _optcat2 != None else blank
        _optcat3 = _re.search('<ns0:OptCat_3>(.*)</ns0:OptCat_3>', resp)
        _optcat3 = _optcat3 if _optcat3 != None else blank
        _assignee = _re.search('<ns0:Assignee>(.*)</ns0:Assignee>', resp)
        _assignee = _assignee if _assignee != None else blank
        _custID = _re.search('<ns0:CustomerID>(.*)</ns0:CustomerID>', resp)
        _custID = _custID if _custID != None else blank
        _statReason = _re.search('<ns0:StatusReason>(.*)</ns0:StatusReason>', resp)
        _statReason = _statReason if _statReason != None else blank
        _assignLoginID = _re.search('<ns0:AssigneeLoginID>(.*)</ns0:AssigneeLoginID>', resp)
        _assignLoginID = _assignLoginID if _assignLoginID != None else blank
        _resComment = _re.search('<ns0:ResComment>(.*)</ns0:ResComment>', resp)
        _resComment = _resComment if _resComment != None else blank
        
        return [_summary.groups(0)[0], _notes.groups(0)[0], _status.groups(0)[0], _assignedGroup.groups(0)[0], _assignedGroupID.groups(0)[0],
                _optcat1.groups(0)[0], _optcat2.groups(0)[0], _optcat3.groups(0)[0], _assignee.groups(0)[0], _custID.groups(0)[0],
                _statReason.groups(0)[0], _assignLoginID.groups(0)[0], _resComment.groups(0)[0]]



class modTicket:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_StatusMod_Service'}
    def __init__(self, ticNum, status="Assigned", statReason = "", resComment = "Test"):
        self._ticNum = ticNum
        self._status = status
        self._statReason = statReason
        self._resComment = resComment
        
        
    def mod(self):
        _
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


class reassignTic:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_Reassignment_Service'}
    def __init__(self, ticNum, status="Assigned", reassignTo="GSD Automation", reassignReason=""):
        self.ticNum = ticNum
        self.status = status
        self.reassignTo = reassignTo
        self.reassignReason = reassignReason

    def reassign(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_Automation_1">
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


class getWorkNotes:
    _headers = {'content-type': 'text/xml;charset=UTF-8', 'SOAPAction': 'urn:CAP:HPD_IncidentInterface_EUS_Automation_1/HelpDesk_GetWorkInfoList_Service'}
    def __init__(self, ticNum):
        self.ticNum = ticNum

    def get(self):
        bg1 = """<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CAP:HPD_IncidentInterface_EUS_Automation_1">
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


class mail:
    def __init__(self, to="noexistant.domain", cc="",  subject="", body="", attachment=False, html=None):
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

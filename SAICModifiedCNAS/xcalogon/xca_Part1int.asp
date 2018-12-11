<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head><script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

DIM TypeOfRequest, RequestStatus


'Dim OtherCarrierType as String

P1EditTixCancel=clng(session("P1EditTixCancel"))
EntityID=cint(session("P1UserEntityID"))
UserID=session("UserUserID")
NPA=session("P1CONPA")
BlankP1=session("BlankP1")

AuthorizedRep=Replace(Request.Form("AuthorizedRep"),"'","''")
AuthorizedRepTitle=Replace(Request.Form("AuthorizedRepTitle"),"'","''")
<!-- Added Ucase() by KTWalsh on 2016-01-14 to force the OCN to Upper case for consistency  -->
OCN=Ucase(Replace(Request.Form("OCN"),"'","''"))
<!-- Change by G.Brown on Sept 26 2001 to force the LATA to 888 for Canada LATA=Replace(Request.Form("LATA"),"'","''") -->
LATA=888
WireCenter=Replace(Request.Form("WireCenter"),"'","''")
SwitchID=Ucase(Replace(Request.Form("SwitchID"),"'","''"))
RateCenter=Replace(Request.Form("RateCenterAssignLookup"),"'","''")
RouteNPA=Replace(Request.Form("RouteNPA"),"'","''")
RouteNXX=Replace(Request.Form("RouteNXX"),"'","''")
CenterNPA=Replace(Request.Form("CenterNPA"),"'","''")
CenterNXX=Replace(Request.Form("CenterNXX"),"'","''")

ApplicationDate=Request.Form("ApplicationDate")
CorrespondenceDate=Request.Form("CorrespondenceDate")
if BlankP1="Applicant" then 
ApplicationDate=date()
EntityID=cint(session("P1UserEntityID"))
Response.Write "app<br>"
end if

'debugResponse.Write EntityID 

'Response.Redirect www

RequestedEffDate=Request.Form("RequestedEffDate")
OtherCarrierType=Replace(Request.Form("OtherCarrierType"),"'","''")
TypeOfService=Replace(Request.Form("TypeOfService"),"'","''")
CertificationNoExplained=Replace(Request.Form("CertificationNoExplained"),"'","''")
RequiredYesExplanation=Replace(Request.Form("RequiredYesExplanation"),"'","''")
RequiredNoExplanation=Replace(Request.Form("RequiredNoExplanation"),"'","''")
NXX2A=Replace(Request.Form("NXX2A"),"'","''")
NXX3A=Replace(Request.Form("NXX3A"),"'","''")
'NXX4A=Replace(Request.Form("NXX4A"),"'","''")
'NXX5A=Replace(Request.Form("NXX5A"),"'","''")
NoNXX1A=Replace(Request.Form("NoNXX1A"),"'","''")
NoNXX2A=Replace(Request.Form("NoNXX2A"),"'","''")
NoNXX3A=Replace(Request.Form("NoNXX3A"),"'","''")
NoNXX4A=Replace(Request.Form("NoNXX4A"),"'","''")
NoNXX5A=Replace(Request.Form("NoNXX5A"),"'","''")
NXX2R=Replace(Request.Form("NXX2R"),"'","''")
NXX3R=Replace(Request.Form("NXX3R"),"'","''")
'NXX4R=Replace(Request.Form("NXX4R"),"'","''")'
'NXX5R=Replace(Request.Form("NXX5R"),"'","''")'
NoNXX1R=Replace(Request.Form("NoNXX1R"),"'","''")
NoNXX2R=Replace(Request.Form("NoNXX2R"),"'","''")
NoNXX3R=Replace(Request.Form("NoNXX3R"),"'","''")
NoNXX4R=Replace(Request.Form("NoNXX4R"),"'","''")
NoNXX5R=Replace(Request.Form("NoNXX5R"),"'","''")
NXXUpdate=Replace(Request.Form("NXXUpdate"),"'","''")
RequestNewNecessary=Replace(Request.Form("RequestNewNecessary"),"'","''")
RequestNewOther=Replace(Request.Form("RequestNewOther"),"'","''")
ReasonForRequest=Replace(Request.Form("ReasonForRequest"),"'","''")
CarrierType=Replace(Request.Form("CarrierType"),"'","''")
CertificationRequired=Replace(Request.Form("CertificationRequired"),"'","''")
RequiredCertificationReady=Replace(Request.Form("RequiredCertificationReady"),"'","''")
NXXAssign=Replace(Request.Form("NXXAssign"),"'","''")
NXXUpdate=Replace(Request.Form("NXXUpdate"),"'","''")
NXXReserve=Replace(Request.Form("NXXReserve"),"'","''")
CodeRequestNew =Replace(Request.Form("CodeRequestNew"),"'","''")
AuthorizationPart2=Replace(Request.Form("AuthorizationPart2"),"'","''")
NPAinJeopardy=Replace(Request.Form("NPAinJeopardy"),"'","''")
DateresponseDue=date() + 10
DueDate=date()+10
DateofReceipt=date()
TypeOfRequest=Replace(Request.Form("TypeOfRequest"),"'","''")
SyncField=now()
P1Process="Part1"

NXXGrowthCal=Replace(Request.Form("NXXGrowthCal"),"'","''")
if Trim(Request.Form("TNs")<>"") then
TNs=clng(Request.Form("TNs"))
else
TNs=0
end if
if Trim(Request.Form("Prev6Month1"))<>"" then
Prev6Month1=clng(Request.Form("Prev6Month1"))
else
Prev6Month1=0
end if
if Trim(Request.Form("Prev6Month2"))<>"" then
Prev6Month2=clng(Request.Form("Prev6Month2"))
else
Prev6Month2=0
end if
if Trim(Request.Form("Prev6Month3"))<>"" then
Prev6Month3=clng(Request.Form("Prev6Month3"))
else
Prev6Month3=0
end if
if Trim(Request.Form("Prev6Month4"))<>"" then
Prev6Month4=clng(Request.Form("Prev6Month4"))
else
Prev6Month4=0
end if
if Trim(Request.Form("Prev6Month5"))<>"" then
Prev6Month5=clng(Request.Form("Prev6Month5"))
else
Prev6Month5=0
end if
if Trim(Request.Form("Prev6Month6"))<>"" then
Prev6Month6=clng(Request.Form("Prev6Month6"))
else
Prev6Month6=0
end if
'	Prev6Month2=clng(Request.Form("Prev6Month2"))
'	Prev6Month3=clng(Request.Form("Prev6Month3"))
'	Prev6Month4=clng(Request.Form("Prev6Month4"))
'	Prev6Month5=clng(Request.Form("Prev6Month5"))
'	Prev6Month6=clng(Request.Form("Prev6Month6"))

if Trim(Request.Form("ProjGrowth16Month1")<>"") then
	ProjGrowth16Month1=clng(Request.Form("ProjGrowth16Month1"))
else
	ProjGrowth16Month1=0
end if
if Trim(Request.Form("ProjGrowth16Month2")<>"") then
	ProjGrowth16Month2=clng(Request.Form("ProjGrowth16Month2"))
else
	ProjGrowth16Month2=0
end if
if Trim(Request.Form("ProjGrowth16Month3"))<>"" then
	ProjGrowth16Month3=clng(Request.Form("ProjGrowth16Month3"))
else
	ProjGrowth16Month3=0
end if
if Trim(Request.Form("ProjGrowth16Month4"))<>"" then
	ProjGrowth16Month4=clng(Request.Form("ProjGrowth16Month4"))
else
	ProjGrowth16Month4=0
end if
if Trim(Request.Form("ProjGrowth16Month5"))<>""then
	ProjGrowth16Month5=clng(Request.Form("ProjGrowth16Month5"))
else
	ProjGrowth16Month5=0
end if
if Trim(Request.Form("ProjGrowth16Month6"))<>"" then
	ProjGrowth16Month6=clng(Request.Form("ProjGrowth16Month6"))
else
	ProjGrowth16Month6=0
end if

'ProjGrowth16Month2=clng(Request.Form("ProjGrowth16Month2"))
'ProjGrowth16Month3=clng(Request.Form("ProjGrowth16Month3"))
'ProjGrowth16Month4=clng(Request.Form("ProjGrowth16Month4"))
'ProjGrowth16Month5=clng(Request.Form("ProjGrowth16Month5"))
'ProjGrowth16Month6=clng(Request.Form("ProjGrowth16Month6"))
if Trim(Request.Form("ProjGrowth712Month1"))<>"" then
ProjGrowth712Month1=clng(Request.Form("ProjGrowth712Month1"))
else
ProjGrowth712Month1=0
end if
if Trim(Request.Form("ProjGrowth712Month2"))<>"" then
ProjGrowth712Month2=clng(Request.Form("ProjGrowth712Month2"))
else
ProjGrowth712Month2=0
end if
if Trim(Request.Form("ProjGrowth712Month3"))<>"" then
ProjGrowth712Month3=clng(Request.Form("ProjGrowth712Month3"))
else
ProjGrowth712Month3=0
end if
if Trim(Request.Form("ProjGrowth712Month4"))<>"" then
ProjGrowth712Month4=clng(Request.Form("ProjGrowth712Month4"))
else
ProjGrowth712Month4=0
end if
if Trim(Request.Form("ProjGrowth712Month5"))<>"" then
ProjGrowth712Month5=clng(Request.Form("ProjGrowth712Month5"))
else
ProjGrowth712Month5=0
end if
if Trim(Request.Form("ProjGrowth712Month6"))<>"" then
ProjGrowth712Month6=clng(Request.Form("ProjGrowth712Month6"))
else
ProjGrowth712Month6=0
end if
'ProjGrowth712Month2=clng(Request.Form("ProjGrowth712Month2"))
'ProjGrowth712Month3=clng(Request.Form("ProjGrowth712Month3"))
'ProjGrowth712Month4=clng(Request.Form("ProjGrowth712Month4"))
'ProjGrowth712Month5=clng(Request.Form("ProjGrowth712Month5"))
'ProjGrowth712Month6=clng(Request.Form("ProjGrowth712Month6"))

AvgMonGrowthRate=csng(Request.Form("AvgMonGrowthRate"))
MonthsToExhaust=csng(Request.Form("MonthsToExhaust"))
AppendixBExplanation=Replace(Request.Form("AppendixBExplanation"),"'","''")
''
',NXXGrowthCal,TNs,Prev6Month1,Prev6Month2,Prev6Month3,Prev6Month4,Prev6Month5,Prev6Month6,ProjGrowth16Month1,ProjGrowth16Month2,ProjGrowth16Month3,ProjGrowth16Month4,ProjGrowth16Month5,ProjGrowth16Month6,ProjGrowth712Month1,ProjGrowth712Month2,ProjGrowth712Month3,ProjGrowth712Month4,ProjGrowth712Month5,ProjGrowth712Month6,AvgMonGrowthRate,MonthsToExhaust,AppendixBExplanation,
''
''

SQLgetParm= "Select Value from xca_Parms where Name='P1DAYS'"
    P1DataCon.setSQLText(SQLgetParm)
	P1DataCon.open
         
       Getdiffnum = P1DataCon.fields.getValue("Value")
       session("P1diffnum")=Getdiffnum
       P1DataCon.close
       
session("P1DiffErr")= null
       



Select Case TypeOfRequest
Case "A"
	 RequestStatus="NW"
	 Status="Q"
	 LogNPA=NPA
	 NXX1Preferred=NXXAssign
	 LogNXX=NXX1Preferred
	 Action="Input"
	 ActionText="Assignment"
	 NXX2=NXX2A
	 NXX3=NXX3A
	 'NXX4=NXX4A
	 'NXX5=NXX5A
	 NoNXX1=NoNXX1A
	 NoNXX2=NoNXX2A
	 NoNXX3=NoNXX3A
	 NoNXX4=NoNXX4A
	 NoNXX5=NoNXX5A
	 
	 
	 d= datediff("d", ApplicationDate, RequestedEffDate)
	
	 if d < clng(Getdiffnum) then 
		session("P1DiffErr")= "Missing"
	 else 
		session("P1DiffErr")= "true"
 	end if
	 
Case "U"
	RequestStatus="UP"
	Status="I"
	LogNPA=NPA
	NXX1Preferred=NXXUpdate
	LogNXX=NXX1Preferred
	Action="Applied"
	ActionText="Update"

Case "R"		
	RequestStatus="NW"
	Status="Q"
	LogNPA=NPA
	NXX1Preferred=NXXReserve
	LogNXX=NXX1Preferred
	Action="Input"
	ActionText="Reservation"
	NXX2=NXX2R
	NXX3=NXX3R
	'NXX4=NXX4R
	'NXX5=NXX5R
	NoNXX1=NoNXX1R
	NoNXX2=NoNXX2R
	NoNXX3=NoNXX3R
	NoNXX4=NoNXX4R
	NoNXX5=NoNXX5R

End Select   
''''''''''''''''''''''''''''''''''''''''''''''''''
'check for dup tix
If P1EditTixCancel =0 or P1EditTixCancel="" then

SQLcheckDup= "Select * from xca_Part1 where RequestStatus='NW' and NPA='"&LogNPA&"' and NXX1preferred='"&LogNXX&"'"
    P1DataCon.setSQLText(SQLcheckDup)
	P1DataCon.open
         
       P1DupTix = P1DataCon.fields.getValue("RequestStatus")
    
     
     

	If P1DupTix ="NW" then 
		 P1DataCon.close
		session("P1DupTix")=P1DupTix
		Response.redirect "xca_Part3Deny.asp"
	end if
	P1DataCon.close
end if
''''''''''''''''''''''''''''''''''''''''''''''''''	
'check route and center npa is in service
session("CenterNPAErr")="" 	
session("RouteNPAErr")=""	


if RouteNPA <>"" then 
	chkRoute="Select Status from xca_COCode where NPA='"&RouteNPA&"' and NXX='"&RouteNXX&"'"
    P1DataCon.setSQLText(chkRoute)
	P1DataCon.open
	CORouteStatus=P1DataCon.fields.getValue("Status")
	
	If CORouteStatus = "I" then 
	else
		session("RouteNPAErr")= "Missing"	
	end if
	P1DataCon.close
end if

if CenterNPA <>"" then
	chkCenter="Select Status from xca_COCode where NPA='"&CenterNPA&"' and NXX='"&CenterNXX&"'"
	P1DataCon.setSQLText(chkCenter)
	P1DataCon.open
	COCenterStatus=P1DataCon.fields.getValue("Status")
	
	If COCenterStatus = "I" then
	else 
		session("CenterNPAErr")="Missing"
	end if
	P1DataCon.close
end if 
''''''''''''''''''''''''''''''''''''''''''''''''''
'check for errors and display bad fields.  

If session("CenterNPAErr") ="Missing" then 
	Response.redirect "xca_Part1Missing.asp"
end if

if session("P1DiffErr") ="Missing" then 
	Response.redirect "xca_Part1Missing.asp"
end if

if session("RouteNPAErr") ="Missing" then 
	Response.redirect "xca_Part1Missing.asp"
end if

''''''''''''''''''''''''''''''''''''''''''''''''''
'check COCode table if code is available


CheckCOCode="Select Status from xca_COCode where NPA= '"&LogNPA&"' and NXX= '"&LogNXX&"'"
	
	P1DataCon.setSQLText(CheckCOCode)
	P1DataCon.open
	
		NXXStatus = P1DataCon.fields.getValue("Status")
		P1DataCon.close
		
	if (NXXStatus = "S") or (NXXStatus = "Q") or (NXXStatus = "R")or (NXXStatus = "I")then
	
	else
	Response.Redirect "xca_Part3Deny.asp"
	end if
''''''''''''''''''''''''''''''''''''''''''''''''''
	 
 'Update Part 1 and get tix #
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

Response.Write "INSERT INTO xca_Part1 (RequiredYesExplanation, DateofReceipt, DueDate, SyncField, UserID, DateresponseDue, EntityID, RequestStatus, CodeRequestNew, AuthorizationPart2, NPAinJeopardy, NXX1Preferred, RequiredCertificationReady, CarrierType, CertificationRequired, NPA, TypeOfRequest, ReasonForRequest, NXXUpdate, RequestNewNecessary, RequestNewOther, NoNXX1, NoNXX2, NoNXX3, NoNXX4, NoNXX5, NXX2, NXX3, CertificationNoExplained, RequiredNoExplanation, OtherCarrierType, TypeOfService, ApplicationDate, RequestedEffDate, CenterNPA, CenterNXX, RateCenter, RouteNPA, RouteNXX, WireCenter, SwitchID, LATA , OCN, AuthorizedRep,  AuthorizedRepTitle,NXXGrowthCal,TNs,Prev6Month1,Prev6Month2,Prev6Month3,Prev6Month4,Prev6Month5,Prev6Month6,ProjGrowth16Month1,ProjGrowth16Month2,ProjGrowth16Month3,ProjGrowth16Month4,ProjGrowth16Month5,ProjGrowth16Month6,ProjGrowth712Month1,ProjGrowth712Month2,ProjGrowth712Month3,ProjGrowth712Month4,ProjGrowth712Month5,ProjGrowth712Month6,AvgMonGrowthRate,MonthsToExhaust,AppendixBExplanation,CorrespondenceDate)  VALUES ('"& RequiredYesExplanation &"','"& DateofReceipt &"','"& DueDate &"','"& SyncField &"','"& UserID &"','"& DateresponseDue &"','"& EntityID &"','"& RequestStatus &"','"& CodeRequestNew &"','"& AuthorizationPart2 &"','"& NPAinJeopardy &"','"& NXX1Preferred &"','"& RequiredCertificationReady &"','"& CarrierType &"','"& CertificationRequired &"','"& NPA &"','"& TypeOfRequest &"','"& ReasonForRequest &"','"& NXXUpdate &"','"& RequestNewNecessary &"','"& RequestNewOther &"','"& NoNXX1 &"','"& NoNXX2 &"','"& NoNXX3 &"','"& NoNXX4 &"','"& NoNXX5 &"','"& NXX2 &"','"& NXX3 &"','"& CertificationNoExplained &"','"& RequiredNoExplanation &"','"& OtherCarrierType &"','"& TypeOfService &"','"& ApplicationDate &"','"& RequestedEffDate &"','"& CenterNPA &"','"& CenterNXX &"','"& RateCenter &"','"& RouteNPA &"','"& RouteNXX &"','"& WireCenter &"','"& SwitchID &"','"& LATA &"','"& OCN &"','"& AuthorizedRep &"','"& AuthorizedRepTitle &"','"&NXXGrowthCal&"','"&TNs&"','"&Prev6Month1&"','"&Prev6Month2&"','"&Prev6Month3&"','"&Prev6Month4&"','"&Prev6Month5&"','"&Prev6Month6&"','"&ProjGrowth16Month1&"','"&ProjGrowth16Month2&"','"&ProjGrowth16Month3&"','"&ProjGrowth16Month4&"','"&ProjGrowth16Month5&"','"&ProjGrowth16Month6&"','"&ProjGrowth712Month1&"','"&ProjGrowth712Month2&"','"&ProjGrowth712Month3&"','"&ProjGrowth712Month4&"','"&ProjGrowth712Month5&"','"&ProjGrowth712Month6&"','"&AvgMonGrowthRate&"','"&MonthsToExhaust&"','"&AppendixBExplanation&"','"&CorrespondenceDate&"')" 

Response.Write "Ready to write the canned row - KT aaabbb" + vbCRLF

'SQLstmt = "INSERT INTO xca_Part1 (RequiredYesExplanation, DateofReceipt, DueDate, SyncField, UserID, DateresponseDue, EntityID, RequestStatus, CodeRequestNew, AuthorizationPart2, NPAinJeopardy, NXX1Preferred, RequiredCertificationReady, CarrierType, CertificationRequired, NPA, TypeOfRequest, ReasonForRequest, NXXUpdate, RequestNewNecessary, RequestNewOther, NoNXX1, NoNXX2, NoNXX3, NoNXX4, NoNXX5, NXX2, NXX3, CertificationNoExplained, RequiredNoExplanation, OtherCarrierType, TypeOfService, ApplicationDate, RequestedEffDate, CenterNPA, CenterNXX, RateCenter, RouteNPA, RouteNXX, WireCenter, SwitchID, LATA , OCN, AuthorizedRep, AuthorizedRepTitle,NXXGrowthCal,TNs,Prev6Month1,Prev6Month2,Prev6Month3,Prev6Month4,Prev6Month5,Prev6Month6,ProjGrowth16Month1,ProjGrowth16Month2,ProjGrowth16Month3,ProjGrowth16Month4,ProjGrowth16Month5,ProjGrowth16Month6,ProjGrowth712Month1,ProjGrowth712Month2,ProjGrowth712Month3,ProjGrowth712Month4,ProjGrowth712Month5,ProjGrowth712Month6,AvgMonGrowthRate,MonthsToExhaust,AppendixBExplanation,CorrespondenceDate) VALUES ('asdf','08/05/2013','08/05/2013','08/05/2013','195','08/05/2013','159','NW','','n','','200','Y','l','Y','204','A','aic','','','','','','','','','200','200','','','','asdf','08/05/2013','08/05/2013','','','Belmont','','','asd','ASDFASDFASD','888','asdf','asdf','asdf','','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','','31/05/2013')"
SQLstmt = "INSERT INTO xca_Part1 (RequiredYesExplanation, DateofReceipt, DueDate, SyncField, UserID, DateresponseDue, EntityID, RequestStatus, CodeRequestNew, AuthorizationPart2, NPAinJeopardy, NXX1Preferred, RequiredCertificationReady, CarrierType, CertificationRequired, NPA, TypeOfRequest, ReasonForRequest, NXXUpdate, RequestNewNecessary, RequestNewOther, NoNXX1, NoNXX2, NoNXX3, NoNXX4, NoNXX5, NXX2, NXX3, CertificationNoExplained, RequiredNoExplanation, OtherCarrierType, TypeOfService, ApplicationDate, RequestedEffDate, CenterNPA, CenterNXX, RateCenter, RouteNPA, RouteNXX, WireCenter, SwitchID, LATA , OCN, AuthorizedRep,  AuthorizedRepTitle,NXXGrowthCal,TNs,Prev6Month1,Prev6Month2,Prev6Month3,Prev6Month4,Prev6Month5,Prev6Month6,ProjGrowth16Month1,ProjGrowth16Month2,ProjGrowth16Month3,ProjGrowth16Month4,ProjGrowth16Month5,ProjGrowth16Month6,ProjGrowth712Month1,ProjGrowth712Month2,ProjGrowth712Month3,ProjGrowth712Month4,ProjGrowth712Month5,ProjGrowth712Month6,AvgMonGrowthRate,MonthsToExhaust,AppendixBExplanation,CorrespondenceDate)  VALUES ('"& RequiredYesExplanation &"','"& DateofReceipt &"','"& DueDate &"','"& SyncField &"','"& UserID &"','"& DateresponseDue &"','"& EntityID &"','"& RequestStatus &"','"& CodeRequestNew &"','"& AuthorizationPart2 &"','"& NPAinJeopardy &"','"& NXX1Preferred &"','"& RequiredCertificationReady &"','"& CarrierType &"','"& CertificationRequired &"','"& NPA &"','"& TypeOfRequest &"','"& ReasonForRequest &"','"& NXXUpdate &"','"& RequestNewNecessary &"','"& RequestNewOther &"','"& NoNXX1 &"','"& NoNXX2 &"','"& NoNXX3 &"','"& NoNXX4 &"','"& NoNXX5 &"','"& NXX2 &"','"& NXX3 &"','"& CertificationNoExplained &"','"& RequiredNoExplanation &"','"& OtherCarrierType &"','"& TypeOfService &"','"& ApplicationDate &"','"& RequestedEffDate &"','"& CenterNPA &"','"& CenterNXX &"','"& RateCenter &"','"& RouteNPA &"','"& RouteNXX &"','"& WireCenter &"','"& SwitchID &"','"& LATA &"','"& OCN &"','"& AuthorizedRep &"','"& AuthorizedRepTitle &"','"&NXXGrowthCal&"','"&TNs&"','"&Prev6Month1&"','"&Prev6Month2&"','"&Prev6Month3&"','"&Prev6Month4&"','"&Prev6Month5&"','"&Prev6Month6&"','"&ProjGrowth16Month1&"','"&ProjGrowth16Month2&"','"&ProjGrowth16Month3&"','"&ProjGrowth16Month4&"','"&ProjGrowth16Month5&"','"&ProjGrowth16Month6&"','"&ProjGrowth712Month1&"','"&ProjGrowth712Month2&"','"&ProjGrowth712Month3&"','"&ProjGrowth712Month4&"','"&ProjGrowth712Month5&"','"&ProjGrowth712Month6&"','"&AvgMonGrowthRate&"','"&MonthsToExhaust&"','"&AppendixBExplanation&"','"&CorrespondenceDate&"')"

	
	objCmd.CommandText=SQLStmt
	objCmd.Execute
	
	
	'get Tix and P1EntityEmail
	SQLget= "Select Tix from xca_Part1 where SyncField= '"&SyncField&"' and UserID= '"&UserID&"'"
    P1DataCon.setSQLText(SQLget)
	P1DataCon.open
         
       RetrieveTix = P1DataCon.fields.getValue("Tix")
       P1DataCon.close
       
				session("P1COTix")=RetrieveTix
				
	SQLgetEEmail= "Select EntityEmail from xca_Entity where EntityID='"&EntityID&"'"
    P1DataCon.setSQLText(SQLgetEEmail)
	P1DataCon.open
         
      P1EntityEmail = P1DataCon.fields.getValue("EntityEmail")
      P1DataCon.close	
      
	SQLgetUserEEmail= "Select UserEmail from xca_User,xca_Part1  where xca_Part1.Tix = '"&P1EditTixCancel&"' and xca_User.UserID=xca_Part1.UserID  "
    P1DataCon.setSQLText(SQLgetUserEEmail)
	P1DataCon.open
         
      P1UserEmail = P1DataCon.fields.getValue("UserEmail")
      P1DataCon.close				
''''''''''''''''''''''''''''''''''''''''''''''''''				
				
'Close old Edit Tix if there
	if P1EditTixCancel <>0 then
	SQLCloseOldP1="Update xca_Part1 Set RequestStatus= 'CC' where Tix='"&P1EditTixCancel&"'"
	
	objCmd.CommandText=SQLCloseOldP1
	objCmd.Execute
	end if
''''''''''''''''''''''''''''''''''''''''''''''''''	
'Remove existing Tix info
	SQLstmt1="Update xca_COCode Set Status= 'S', EntityID= 0,Tix=0 where Tix='"&P1EditTixCancel&"'  "
	objCmd.CommandText=SQLStmt1
	objCmd.Execute
'Update COCodes with tix#
Response.Write "Update xca_COCode Set Status= '"&Status&"', EntityID= '"&EntityID&"', Tix='"&RetrieveTix&"' where NPA= '"&NPA&"' and NXX='"&NXX1Preferred&"'" + vbCRLF

SQLstmt1="Update xca_COCode Set Status= '"&Status&"', EntityID= '"&EntityID&"', Tix='"&RetrieveTix&"' where NPA= '"&NPA&"' and NXX='"&NXX1Preferred&"'"
	objCmd.CommandText=SQLStmt1
	objCmd.Execute
	
	
If  err.number>0 then
      response.write "VBScript Errors Occured:" & "<P>"
      response.write "Error Number=" & err.number & "<P>"
      response.write "Error Descr.=" & err.description & "<P>"
      response.write "Help Context=" & err.helpcontext & "<P>" 
      response.write "Help Path=" & err.helppath & "<P>"
      response.write "Native Error=" & err.nativeerror & "<P>"
      response.write "Source=" & err.source & "<P>"
      response.write "SQLState=" & err.sqlstate & "<P>"
end if
IF  objConn.errors.count> 0 then
      response.write "Database Errors Occured" & "<P>"
      response.write SQLstmt & "<P>"
for counter= 0 to conn.errors.count
      response.write "Error #" & objConn.errors(counter).number & "<P>"
      response.write "Error desc. -> " & conn.errors(counter).description & "<P>"
next
else
		objConn.Close
  session("Part1Complete")="complete"
   
    
end if
''''''''''''''''''''''''''''''''''''''''''
'send email and log

mail1 = session("UserUserEmail")
if P1EntityEmail <> "" then
	mail2 = ", " & P1EntityEmail
else
	mail2 = ""
end if

if P1UserEmail <>"" then
	mail3 = ", "& P1UserEmail
else
	mail3 = ""
end if
twoEmail= mail1 &mail2  & mail3
'
' This section was modified by G. Brown Sep 21, 1999
' 
'emailText="Ticket # " & RetrieveTix & ", CO Code " & LogNPA & " " & LogNXX & ", " & ActionText & " has been submitted.  "
'
emailText="Ticket # " & RetrieveTix & ", NPA - "  & LogNPA & ", NXX - " & LogNXX & ", Rate Centre - " & RateCenter & ", " & ActionText & " has been submitted.  Part 3 will follow."
log "R",LogNPA,LogNXX,UserID,Now,RetrieveTix,Action,ActionText,P1Process
'
'This section was added by G. Brown Sep 7,1999.  S. Khare only wants an email sent -->
'when the applicant fills out the form on-line. -->
'
UserEntityType=session("UserEntityType")
UserUserType=session("UserUserType")
If UserEntityType <> "a" and UserUserType <> "a" then
'   
' This section was modified by G. Brown Sep 21,1999 to reflect a change in message format.
'
'email session("AdminEntityEmail"),twoEmail,"","CNAS Part 1 Status",emailText
email "database@cnac.ca",twoEmail,"","CNAS Part 1 Status",emailText
end if
session("P1TixCook")=RetrieveTix
session("P1NPACook")=LogNPA
session("P1NXXCook")=LogNXX
session("P1TwoEmailsCook")=twoEmail
session("P1RateCenter")=RateCenter

Response.Redirect "xca_part1Confirm.asp"





</script>

<title></title>
</head>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<form name="thisForm" METHOD="post">
<!--#Include file="xca_CNASlib.inc"-->
</form>

<p>&nbsp;</p>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</body>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P1DataCon style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qTables\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qP1DataCon\q,TCPPConn=\qcnasadmin\q,RCDBObject=\qRCDBObject\q,TCPPDBObject_Unmatched=\qTables\q,TCPPDBObjectName_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initP1DataCon()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
//Recordset DTC error: Failed to get command text
	cmdTmp.CommandText = '';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	P1DataCon.setRecordSource(rsTmp);
	if (thisPage.getState('pb_P1DataCon') != null)
		P1DataCon.setBookmark(thisPage.getState('pb_P1DataCon'));
}
function _P1DataCon_ctor()
{
	CreateRecordset('P1DataCon', _initP1DataCon, null);
}
function _P1DataCon_dtor()
{
	P1DataCon._preserveState();
	thisPage.setState('pb_P1DataCon', P1DataCon.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
</html>

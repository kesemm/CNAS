<%@ Language=VBScript %>

<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#Include file="xca_CNASlib.inc"-->

<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<%
Dim COTix
Function validDate (date1,date2)
	if isdatereal(date1) and isdatereal(date2) then
		if DateValue(date1) >= DateValue(date2)then			
			validDate = true
		else 
			validDate = false
		end if
	else 
		validDate = false
	end if
	
end function
P1TypeofRequest=session("P1TypeofRequest")
'P1email=session("P1Email")
P1P3CONPA=session("P1P3CONPA")
CNAUserName=session("UserUserName")  
UserID=int(session("UserUserID"))
EntityID=session("P1EntityID")
Tix=int(Session("P1P3COTix"))
EmailTix=Tix
Part3Result=Replace(Request.Form("Part3Result"),"'","''")
AssignedNPA=Replace(request.form("AssignedNPA"),"'","''")
AssignedNXX=Replace(request.form("AssignedNXX"),"'","''")
AssignedNPANXXDate=Date()
RRComplete=Replace(request.form("RRComplete"),"'","''")
RRDescription=Replace(request.form("RRDescription"),"'","''")
CNAResponsible=Replace(request.form("CNAResponsible"),"'","''")
LERGDate=Replace(request.form("LERGDate"),"'","''")
RRReturnDate=Replace(request.form("RRReturnDate"),"'","''")
ReservedNPA=Replace(request.form("ReservedNPA"),"'","''")
ReservedNXX=Replace(int(request.form("ReservedNXX")),"'","''")
ReservedNPANXXDate=Date()
ReservedNPANXXHonorDate=Replace(request.form("ReservedNPANXXHonorDate"),"'","''")
Part3IncompleteDescription=Replace(request.form("Part3IncompleteDescription"),"'","''")
Part3DenialDescription=Replace(request.form("Part3DenialDescription"),"'","''")
Part3SuspendedDescription=Replace(request.form("Part3SuspendedDescription"),"'","''")
Part3SuspendedFurtherAction=Replace(request.form("Part3SuspendedFurtherAction"),"'","''")
Remarks=Replace(request.form("Remarks"),"'","''")
UpdatedNPA=Replace(Request.Form("UpdatedNPA"),"'","''")
UpdatedNXX=Replace(Request.Form("UpdatedNXX"),"'","''")
UpdatedNPANXXDate=date()
ExtentionDate=Replace(Request.Form("ExtentionDate"),"'","''")
If Request.Form("ExtentionDate") <> "" then
	ExtentionDateVal=cdate(Request.Form("ExtentionDate"))
end if
P3Process="Part3"
EffDate=Replace(Request.Form("EffDate"),"'","''")
P3DateofReceipt=Date()
NPAinJeopardy=Replace(Request.Form("NPAinJeopardy"),"'","''")
ReserveDate450=Date()+450
ApplicationDate=session("ApplicationDate")
appDate180 = cdate(ApplicationDate)+180
appDate360 = cdate(ApplicationDate)+360
appDate540 = cdate(ApplicationDate)+540
'Response.Write ApplicationDate&"<-ApplicationDate<BR>"


'Get P1 data
SQLget= "Select * from xca_Part1 where Tix= '"&Tix&"'"
	P3Access.setSQLText(SQLget)
	P3Access.open
               
			LATA = P3Access.fields.getValue("LATA")
			OCN= P3Access.fields.getValue("OCN")
			SwitchID= P3Access.fields.getValue("SwitchID")
WireCenter= Replace(trim(P3Access.fields.getValue("WireCenter")),"'","''")
RateCenter= Replace(trim(P3Access.fields.getValue("RateCenter")),"'","''")
			NXX1preferred=P3Access.fields.getValue("NXX1preferred")
			' KT CHANGED 2013-06-12: Added formatting around date before it gets tested
			RequestedEffDate=FormatDateTime(P3Access.fields.getValue("RequestedEffDate"),vbShortDate)
			P1EntityID=P3Access.fields.getValue("EntityID")
			P1UserID=P3Access.fields.getValue("UserID")
			NXX2=P3Access.fields.getValue("NXX2")
			NXX3=P3Access.fields.getValue("NXX3")
			NoNXX1=P3Access.fields.getValue("NoNXX1")
			NoNXX2=P3Access.fields.getValue("NoNXX2")
			NoNXX3=P3Access.fields.getValue("NoNXX3")
			NoNXX4=P3Access.fields.getValue("NoNXX4")
			NoNXX5=P3Access.fields.getValue("NoNXX5")
			RequestStatus=P3Access.fields.getValue("RequestStatus")
			NXXUpdate = P3Access.fields.getValue("NXXUpdate")
		'	Response.Write NXXUpdate&"<--NXXUpdate<br>"
			P3Access.close
			
if NXXUpdate <> "" then
	RejNXX=NXXUpdate
end if	
if ReservedNXX <> "" then
	RejNXX=ReservedNXX
end if	
if AssignedNXX <> "" then
	RejNXX=AssignedNXX
end if

LogNPA=P1P3CONPA
session("AssignedNXXErr")= ""

session("EFFDateErr")= ""

session("RRDescriptionErr")= ""

session("LERGDateErr")= ""

session("RRReturnErr")= ""

session("ReservedNXXErr")= ""

session("CNAResponsibleErr")= ""

session("ReservedNXXErr")= ""

Select Case Part3Result


Case  "i" 
	ReTRequestStatus="CI"
	Action="Rejected"
	ActionText="Incomplete"
	COTix= "Deny"
	skip = true
	LogNXX=RejNXX
	Part3ResultTxt=("Part 1 Form Incomplete. "&chr(10)&chr(13)&"Additional information required in the following setion(s): "&chr(10)&chr(13)&Part3IncompleteDescription&chr(10)&chr(13))
Case  "d"
	ReTRequestStatus="CD"
	Action="Rejected"
	ActionText="Denied"
	COTix= "Deny"
	skip = true
	LogNXX=RejNXX
	Part3ResultTxt=("Part 1 Form completed, Code Request denied."&chr(10)&chr(13)&"Explanation is: "&chr(10)&chr(13)&Part3DenialDescription)
	
Case  "s" 
	ReTRequestStatus="CP"
	Action="Rejected"
	ActionText="Suspended"
	COTix= "Deny"
	skip = true
	LogNXX=RejNXX
	Part3ResultTxt=("Part 1 Assignment Activity Suspended by the Administrator."&chr(10)&chr(13)&"Explanation is: "&chr(10)&chr(13)&Part3SuspendedDescription&chr(10)&chr(13)&"Further Action: "&chr(10)&chr(13)&Part3SuspendedFurtherAction)

Case else
		


Select case P1TypeofRequest
case "U"
	UpdatedNPA=P1P3CONPA
	EffDate=date()
	LogNPA=P1P3CONPA
	LogNXX=NXXUpdate
	Part3Result="u"
	ReTRequestStatus="CU"
	Action="Applied"
	ActionText="Updated"
	Status="I"
	Part3ResultTxt = ("Requested NPA: "&LogNPA&chr(10)&chr(13)& "Updated NXX: " &LogNXX)
case "A"
	AssignedNPA=P1P3CONPA
	EffDate=Replace(Request.Form("EffDate"),"'","''")
	LogNPA=P1P3CONPA
	LogNXX=AssignedNXX
	Action="Applied"
	ActionText="Assigned"
	Part3Result="a"
	Status="A"
	ReTRequestStatus="AS"
	
	'if the previous was a reservation also set due date of the 
	'previous reservation to the greater of the (application date + 360) and
	'Extention date else the due date is the effective date
	
	if RequestStatus="RS" then
		if appDate360 > ExtentionDateVal then
			RetDueDate=Replace(appDate360,"'","''")
		else 
			RetDueDate = Replace(ExtentionDateVal,"'","''")
		end if
	else
		RetDueDate=EffDate
	end if
	
	If (not IsNumeric(AssignedNXX)) then
		session("AssignedNXXErr")= "Missing"
	end if
	
	If AssignedNXX = "" then
		session("AssignedNXXErr")= "Missing"
	end if
	
	If AssignedNXX <200 then
		session("AssignedNXXErr")= "Missing"
	end if
	
	If EffDate = "" then
		session("EFFDateErr")= "Missing"
	end if
	
	 
	if not validDate(EffDate,RequestedEffDate)  then 
		session("EFFDateErr")= "Missing"
	end if
	
	if RRComplete = "N" and RRDescription= "" then
		session("RRDescriptionErr")= "Missing"
	End if
	
	If  LERGDate ="" then
		session("LERGDateErr")= "Missing"
	end if
	
	if RRReturnDate ="" then
		session("RRReturnErr")= "Missing"
	end if
	
	If CNAResponsible = "" then
		session("CNAResponsibleErr")= "Missing"
	end if
	Part3ResultTxt1 = "Requested NPA: "&AssignedNPA&chr(10)&chr(13)&"Reserved NXX: "& LogNXX&chr(10)&chr(13)&"Secondary NXXs: "&NXX2&" "&NXX3&chr(10)&chr(13)&"Undersirable NXXs: "&NoNXX1&" "&NoNXX2&" "&NoNXX3&" "&NoNXX4&" "&NoNXX5&chr(10)&chr(13)&"NXX Effective Date: "&EffDate&chr(10)&chr(13) &"Switch Identification(Switching Entity/POI): "&SwitchID&chr(10)&chr(13) &"Rate Center: "&RateCenter &chr(10)&chr(13)	&"Is Routing and Rating information complete?-(Y)es/(N)o  "& RRComplete&chr(10)&chr(13)	&"If no, Additional RDBS and BRIDS information is required as follows:  "&RRDescription &chr(10)&chr(13)
	Part3ResultTxt = Part3ResultTxt1&"The Code Administrator is responsible for inputting Part 2 Information into RDBS and BRIDS? - (Y)es/(N)o"& CNAResponsible&chr(10)&chr(13)&"To be published in the LERG and TPM by: "& LERGDate&chr(10)&chr(13)&"information needs to be received by the Code Administrator no later than: "&RRReturnDate &chr(10)&chr(13)
case "R"

	ReservedNPA=P1P3CONPA
	EffDate=date() + 450
	LogNPA=P1P3CONPA
	LogNXX=ReservedNXX
	Action="Applied"
	ActionText="Reserved"
	Part3Result="r"
	Status="R"
	ReTRequestStatus="RS"
	RetDueDate=cdate(ApplicationDate) + 180
'	RetDueDate=date() + 450
	Part3ResultTxt = "Requested NPA: "&LogNPA&chr(10)&chr(13)&"Reserved NXX: "& LogNXX&chr(10)&chr(13)&"Secondary NXXs: "&NXX2&" "&NXX3&chr(10)&chr(13)&"Undesirable NXXs: "&NoNXX1&" "&NoNXX2&" "&NoNXX3&" "&NoNXX4&" "&NoNXX5&chr(10)&chr(13)&"Date of Reservation: "&ReservedNPANXXDate&chr(10)&chr(13)&"Your code reservation will be honored until: "&ReserveDate450&chr(10)&chr(13)&"Switch Identification(Switching Entity/POI): "&chr(10)&chr(13)
					 
	
	If (not IsNumeric(ReservedNXX)) then
		session("ReservedNXXErr")= "Missing"
	end if
	
	If ReservededNXX = "" then
		session("ReservedNXXErr")= "Missing1"
	end if
	
	If ReservededNXX < 200 then
		session("ReservedNXXErr")= "Missing2"
	end if
	
	


if session("RRDescriptionErr")<> "" then
		Response.Redirect "xca_Part3Missing.asp"
	end if
if session("AssignedNXXErr")<> "" then
		Response.Redirect "xca_Part3Missing.asp"
	end if
if session("EFFDateErr")<> "" then
		Response.Redirect "xca_Part3Missing.asp"
	end if
if session("RRDescriptionErr")<> "" then
		Response.Redirect "xca_Part3Missing.asp"
	end if
if session("LERGDateErr")<> "" then
		Response.Redirect "xca_Part3Missing.asp"
	end if
if session("ReservedNXXErr")<> "" then
		'Response.Redirect "xca_Part3Missing.asp"
		
	end if


end select
End Select





%>

<BODY>


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P3Access style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sP1Access\swhere\sTix=?\q,TCControlID_Unmatched=\qP3Access\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sP1Access\swhere\sTix=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initP3Access()
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
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from P1Access where Tix=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	P3Access.setRecordSource(rsTmp);
	if (thisPage.getState('pb_P3Access') != null)
		P3Access.setBookmark(thisPage.getState('pb_P3Access'));
}
function _P3Access_ctor()
{
	CreateRecordset('P3Access', _initP3Access, null);
}
function _P3Access_dtor()
{
	P3Access._preserveState();
	thisPage.setState('pb_P3Access', P3Access.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->





<%

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
    objCmd.ActiveConnection = objConn	
	
'check to see if P1 CO Code='S' miss if deny
if skip = false then 

CheckCOCode="Select Status from xca_COCode where NPA= '"&LogNPA&"' and NXX= '"&LogNXX&"'"
	
	P3Access.setSQLText(CheckCOCode)
	P3Access.open
	
		NXXStatus = P3Access.fields.getValue("Status")
		P3Access.close
'Response.Write LogNPA & "<-LogNpa<br>" & LogNXX & "<--LogNXX<br>" & NXXStatus&"<--status<br>"
 
	if (NXXStatus = "S") or (NXXStatus = "Q") or (NXXStatus = "R") or (NXXStatus = "I") then
	'no-op
	else
	Response.Redirect "xca_Part3Deny.asp"
	end if
	

end if

if (ExtentionDateVal > appDAte180) and (ExtentionDateVal < appDate540) then
	RetDueDate = ExtentionDate
end if	

'Update P1 Status
'Response.Write ReTRequestStatus&"<- Due Dart (ReTRequestStatus)<BR>"
SQLP1Update="Update xca_Part1 Set RequestStatus= '"&ReTRequestStatus&"', DueDate='"&RetDueDate&"' where Tix='"&Tix&"'"

objCmd.CommandText=SQLP1Update
	objCmd.Execute

'Get P1EntiyID and P1UserID Email address


Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
'
' Changed by G Brown on July 30 2003 to solve Named Pipe Problem
'
'	objCmd.CommandText="Select UserEmail from xca_User where UserID='"&P1UserID&"'"
'				
'	set objRec=objCmd.Execute
      GlenVar1="Select UserEmail from xca_User where UserID='"&P1UserID&"'"
      set objRec=objConn.Execute(GlenVar1)
'
' End Change
	if not objRec.EOF then
		
		
		P1UserEmail=objRec("UserEmail")
		
	else
		P1UserEmail=""						
	end if	
' Changed by G Brown on July 30 2003 to solve Named Pipe Problem
'	objCmd.CommandText="Select EntityEmail from xca_Entity where EntityID='"&P1EntityID&"'"
'				
'set objRec=objCmd.Execute
      GlenVar2="Select EntityEmail from xca_Entity where EntityID='"&P1EntityID&"'"
      set objRec=objConn.Execute(GlenVar2)
'
' End Change
	if not objRec.EOF then
		
		
		P1EntityEmail=objRec("EntityEmail")
		
	else
		P1EntityEmail=""						
	end if	
	
'Delete old Record of the tix 
P3DeleteTix = "DELETE FROM xca_Part3 where Tix = '"&tix&"'"
'Insert Part 3

SQLstmt="INSERT INTO xca_Part3 (P3DateofReceipt, EffDate, ExtentionDate, CNAUserName, UpdatedNPA, UpdatedNXX, UpdatedNPANXXDate, UserID, EntityID, Tix, Part3Result, AssignedNPA, AssignedNXX, AssignedNPANXXDate, RRComplete, RRDescription, CNAResponsible, LERGDate, RRReturnDate, ReservedNPA, ReservedNXX, ReservedNPANXXDate, ReservedNPANXXHonorDate, Part3IncompleteDescription, Part3DenialDescription, Part3SuspendedDescription, Part3SuspendedFurtherAction, Remarks,NPAinJeopardy) VALUES ('"& P3DateofReceipt &"','"& EffDate &"','"& ExtentionDate &"','"& CNAUserName &"','"& UpdatedNPA &"','"& UpdatedNXX &"','"& UpdatedNPANXXDate &"','"& UserID &"','"& EntityID &"','"& Tix &"','"& Part3Result &"','"& AssignedNPA &"','"& AssignedNXX &"','"& AssignedNPANXXDate &"','"& RRComplete &"','"& RRDescription &"','"& CNAResponsible &"','"& LERGDate &"','"& RRReturnDate &"','"& ReservedNPA &"','"& ReservedNXX &"','"& ReservedNPANXXDate &"','"& ReservedNPANXXHonorDate &"','"& Part3IncompleteDescription &"','"& Part3DenialDescription &"','"& Part3SuspendedDescription &"','"& Part3SuspendedFurtherAction &"','"& Remarks &"','"& NPAinJeopardy &"')"
	objCmd.CommandText=P3DeleteTix
	objCmd.Execute   
	objCmd.CommandText=SQLStmt
	objCmd.Execute
	

' if NXX1preferred <> NXXLog,  go to co code table for nxx1preferred and clear tix,status,entity id			
if NXX1preferred <> LogNXX then
RetTix=0 '=zero
RetSTATUS="S"  'available/spare
RetEntityID=0  '=null

SQLstmt1ResetCO="Update xca_COCode Set Status= '"&RetStatus&"', Tix= '"&RetTix&"', EntityID='"&RetEntityID&"' where NPA= '"&LogNPA&"' and NXX='"&NXX1preferred&"'"
	objCmd.CommandText=SQLstmt1ResetCO
	objCmd.Execute
	
		
end if 	

'''check if co code was originally a reserve and an assignment was denied


if skip= true then

	P1oldLook= "select * from xca_Part1 where NPA= '"&LogNPA&"' and NXX1preferred= '"&LogNXX&"' and RequestStatus='RS'"
	P3Access.setSQLText(P1oldLook)
	P3Access.open
	OldTix = P3Access.fields.getValue("Tix")
			
			
	if oldTix <> "" then
		Tix=oldTix
		COTix="accept"
			LATA = P3Access.fields.getValue("LATA")
			OCN= P3Access.fields.getValue("OCN")
			SwitchID= P3Access.fields.getValue("SwitchID")
WireCenter= Replace(trim(P3Access.fields.getValue("WireCenter")),"'","' '")
RateCenter= Replace(trim(P3Access.fields.getValue("RateCenter")),"'","' '")
			NXX1preferred=P3Access.fields.getValue("NXX1preferred")
			RequestedEffDate=P3Access.fields.getValue("RequestedEffDate")
			P1EntityID=P3Access.fields.getValue("EntityID")
			P1UserID=P3Access.fields.getValue("UserID")
			Status="R"
	end if
			P3Access.close



end if
						
'Update COCodes	only if assign, reserve, or update and if a previouse reserve existed	
	UpdatedNXX=NXXUpdate		
if COTix <> "Deny" then
	
	SQLstmtCOCode="Update xca_COCode Set Status= '"&Status&"', EntityID= '"&EntityID&"', Tix= '"&Tix&"', SwitchID= '"&SwitchID&"', WireCenter= '"&WireCenter&"', LATA= '"&LATA&"', OCN='"&OCN&"', RateCenter='"&RateCenter&"' where NPA= '"&LogNPA&"' and NXX='"&LogNXX&"'"
	objCmd.CommandText=SQLstmtCOCode
	objCmd.Execute

else
'cleanup incomplete, denied, suspend records
Status="S"
EntityID= 0
RejTix=0
SwitchID=""
WireCenter=""
LATA=""
OCN=""
RateCenter=""
SQLstmtCOCode="Update xca_COCode Set Status= '"&Status&"', EntityID= '"&EntityID&"', Tix= '"&RejTix&"', SwitchID= '"&SwitchID&"', WireCenter= '"&WireCenter&"', LATA= '"&LATA&"', OCN='"&OCN&"', RateCenter='"&RateCenter&"' where NPA= '"&LogNPA&"' and NXX='"&RejNXX&"'"
	objCmd.CommandText=SQLstmtCOCode
	objCmd.Execute
	
end if

		
	
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
NPAJeopTxt = "NPA Jeopardy: "&NPAinJeopardy&chr(10)&chr(13) 
ARDateTxt =  "Date of Requested Application: "&ApplicationDate&chr(10)&chr(13) 			&"Date of Receipt:"&P3DateofReceipt&chr(10)&chr(13) 			&"Requested Effective Date of CO Code: "&EffDate&chr(10)&chr(13) 			&"Extention Date: "&ExtentionDate&chr(10)&chr(13)
		
session("Part3Complete")="complete"
twoEmail=session("UserUserEmail") & ", " & P1EntityEmail & "," & P1UserEmail
emailText= "Ticket # " & EmailTix & ", CO Code: " & LogNPA & " " & LogNXX & ", Status: " & ActionText & ". View part 3 for details "&chr(10)&chr(13)  &ARDateTxt&Part3ResultTxt&NPAJeopTxt

log  "R",LogNPA,LogNXX,UserID,Now,EmailTix,Action,ActionText,P3Process   

'email  session("AdminEntityEmail"),twoEmail,"","CNAS Part 3 Status",emailText

session("P3TixCook")=EmailTix
session("P3NPACook")=LogNPA
session("P3NXXCook")=LogNXX
session("P3RateCenter")=RateCenter
session("P3TwoEmailsCook")=twoEmail
	
	if skip = false then
	Response.Redirect "xca_Part3Confirm.asp"
	
	end if
	
	
	Response.Redirect "xca_Part3DenyP1.asp"


%>


</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

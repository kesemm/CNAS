<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers

function NoTix(){
        
        
	alert("That Ticket does not exist.  Please try again.....");

}
// end hiding -->
</script>


</head>

<%
Sub btnReturnToMenu_onclick()
	Response.Redirect "xca_MenuRptPost.asp"
End Sub

TixManual=Request.Form("Tix")
'Tix=int(session("Tix"))
Tix=session("Tix")

if  session("Tix") =""  then
'if  int(session("Tix")) =""  then
	Tix=TixManual
end if




'check to see where data coming from
ViewP134=session("ViewP134")

Tix=int(Request.Form("Tix"))
UserEntityID=int(session("UserEntityID"))
session("P134Tix")=Tix
session("P1EntityID")=AppEntityID
'AdminUserName=session("UserUserName")
UserLogon=session("UserLogon")

	If session("UserEntityType")= "a" then
		sqlnoTix="Select * from xca_Part1 where Tix= '"&Tix&"'"
			GetPart1Data.setSQLText(sqlnoTix)
			GetPart1Data.Open
			checkTIX=GetPart1Data.fields.getValue("Tix")
			UserEntityID=GetPart1Data.fields.getValue("EntityID")
	End If

	If session("UserEntityType") = "u" then
		sqlno12Tix="Select * from xca_Part1 where Tix= '"&Tix&"' and EntityID='"&UserEntityID&"'"
			GetPart1Data.setSQLText(sqlno12Tix)
			GetPart1Data.Open
			checkTIX = GetPart1Data.fields.getValue("Tix")
			UserEntityID = GetPart1Data.fields.getValue("EntityID")
	End If
			'Check for invalid tix
if checkTix="" then	
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

end if


session("NoTixSent")=""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
AdminEntityID=int(session("AdminEntityID"))
''sql 1 gets the entity recordset of the ADMIN
sql1="Select EntityID from xca_Entity where EntityName= '"&AdminEntityID&"'"


AdminData=session("ADMIN")

'get Admin info for top of form
sqlADMIN="Select * from xca_Entity where EntityName ='"&AdminData&"'"
	GetAdminEntityName.setSQLText(sqlADMIN)
	GetAdminEntityName.Open
	
''sql2 gets the p1 recordset of the P1 request using the Tix from the input form.
sql2="Select * from xca_Part1 where Tix= '"&Tix&"'"
	GetPart1Data.setSQLText(sql2)
	GetPart1Data.Open
	P1UserID = GetPart1Data.fields.getValue("UserID")
'''''This P1UserID is giving the correct response.
	'Response.Write "P1UserID is " & P1UserID
	
	
RequestStatusValue=GetPart1Data.fields.getValue("RequestStatus")
Select Case RequestStatusValue
Case "NW"
	'Response.Redirect "xca_RptPrtsFrmsDeny.asp"
Case "UP"
	'Response.Redirect "xca_RptPrtsFrmsDeny.asp"
Case "AS"
	'Response.Redirect "xca_RptPrtsFrmsDeny.asp"
Case "RS"
	'Response.Redirect "xca_RptPrtsFrmsDeny.asp"
Case "CU"
	RequestStatuschar="Closed - Updated"
Case "CD"
	RequestStatuschar="Closed - Denied"
Case "CI"
	RequestStatuschar="Closed - Incomplete"
Case "CP"
	RequestStatuschar="Closed - Suspended"
Case "CS"
	RequestStatuschar="Closed - InService"
Case "CA"
	RequestStatuschar="Closed - Assigned"
Case "CC"
	RequestStatuschar="Closed - Cancelled by Code Applicant"
End Select


if RequestStatusValue="NW" or RequestStatusValue="AS" or RequestStatusValue="UP" or RequestStatusValue="RS" then
	Response.Redirect "xca_RptPrtsFrmsDeny.asp"
elseif RequestStatuschar="Closed - InService" then

''sql gets user entity recordset of user
'sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserName = '"&UserName&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserID = '"&P1UserID&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
'sql = "Select * from xca_Entity where EntityID = '"&UserEntityID&"'"
	GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open
''''''This sql statement is getting the correct information from Tix, P1UserI, and UserEntityID.
	'Response.Write "P1UserID is " & P1UserID
	'Response.Write "UserEntityID is " & UserEntityID
	Tix=GetUserEntityName.fields.getValue("Tix")
	'Response.Write "Tix is " & Tix
	AppEntityName=GetUserEntityName.fields.getValue("EntityName")
	'Response.Write "AppEntityName is " & AppEntityName
	UserName=GetUserEntityName.fields.getValue("UserName")
	'Response.Write "UserName is " &UserName
	UserEmail=GetUserEntityName.fields.getValue("UserEmail")
	'Response.Write "UserEmail is " & UserEmail
	UserFax=GetUserEntityName.fields.getValue("UserFax")
	'Response.Write "UserFax is " & UserFax
	UserTelephone=GetUserEntityName.fields.getValue("UserTelephone")
	'Response.Write "UserTelephone is " & UserTelephone
	UserExtension=GetUserEntityName.fields.getValue("UserExtension")
	'Response.Write "UserExtension is " & UserExtension
	EntityPostalCode=GetUserEntityName.fields.getValue("EntityPostalCode")
	'Response.Write "EntityPostalCode is " & EntityPostalCode
	EntityAddress=GetUserEntityName.fields.getValue("EntityAddress")
	EntityCity=GetUserEntityName.fields.getValue("EntityCity")
	EntityProvince=GetUserEntityName.fields.getValue("EntityProvince")



sqlrsv="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'R'"
	ToRRsvRec.setSQLText(sqlrsv)
	ToRRsvRec.open

sqlass="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'A'"
	ToRAssRec.setSQLText(sqlass)
	ToRAssRec.open


''sql3 gets the NPASplitID of the NPA-preferred NXX that was requested from the COCode table
sql3="SELECT  xca_COCode.NPASplitID, xca_Part1.NXX1preferred FROM xca_COCode FULL OUTER JOIN xca_Part1 ON xca_COCode.Tix = xca_Part1.Tix WHERE xca_Part1.Tix = '"&Tix&"'"
	GetCOCodeData.setSQLText(sql3)
	GetCOCodeData.Open
	NPASplitIDValue=GetCOCodeData.fields.getValue("NPASplitID")
	Select Case NPASplitIDValue
	Case "Included"
		NPASplitIDValue1="Included"
	Case "Excluded"
		NPASplitIDValue1="Excluded"
	End Select



CodeReqNew=GetPart1Data.fields.getValue("CodeRequestNew")
Select Case CodeReqNew
Case "c"
	CodeReqNewchar1="**"
Case "o"
	CodeReqNewchar2="**"
End Select


CertReqValue=GetPart1Data.fields.getValue("CertificationRequired")
Select Case CertReqValue
Case "Y"
	CertReqchar="YES"
Case "N"
	CertReqchar="NO"
End Select


ReqCertReadyValue=GetPart1Data.fields.getValue("RequiredCertificationReady")
Select Case ReqCertReadyValue
Case "Y"
	ReqCertReadychar="YES"
Case "N"
	ReqCertReadychar="NO"
End Select


TypeEntityValue=GetPart1Data.fields.getValue("CarrierType")
Select Case TypeEntityValue
Case "l"
	TypeEntitychar="Local Exchange Carrier"
Case "w"
	TypeEntitychar="Wireless Service Provider"
Case "o"
	TypeEntitychar="Other(Specify)"
End Select


AuthPart2Value=GetPart1Data.fields.getValue("AuthorizationPart2")
Select Case AuthPart2Value
Case "y"
	AuthPart2char1="**"
Case "n"
	AuthPart2char2="**"
End Select


TyReqvalue=GetPart1Data.fields.getValue("TypeOfRequest")
Select Case TyReqvalue
Case "A"
	TyReqchar1="**"
Case "U"
	TyReqchar2="**"
Case "R"
	TyReqchar3="**"
End Select


Reas4ReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case Reas4ReqValue
Case "aic"
	Reas4Reqchar="a) Initial Code for new Switching Entity or new Point of Interconnection (Complete Part 2)"
Case "aau"
	Reas4Reqchar="b) Code request for New Application for existing switching entity or point of interconnection (Code Aplicant must complete Section 1.7)"
Case "aag"
	Reas4Reqchar="c) Additional Code for Growth (Code Applicant must complete Section 1.6)"
End Select


ReasForReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case ReasForReqValue
Case "ric"
	ReasForReqchar="a) Initial Code"
Case "rau"
	ReasForReqchar="b) New Application (Complete Section 1.7)"
Case "rag"
	ReasForReqchar="c) Growth (Complete Section 1.6)"
End Select


JeopardyValue = GetPart1Data.fields.getValue("NPAinJeopardy")
select case JeopardyValue
case "y"
	JeopardyName1="YES"
	P3Jeopardychar="YES"
case "n"
	JeopardyName2="NO"
	P3Jeopardychar="NO"
case else
	P3Jeopardychar="NO"
end select


NXXGrowthCalValue = GetPart1Data.fields.getValue("NXXGrowthCal")
TNsValue = GetPart1Data.fields.getValue("TNs")
Prev6Month1Value = GetPart1Data.fields.getValue("Prev6Month1")
Prev6Month2Value = GetPart1Data.fields.getValue("Prev6Month2")
Prev6Month3Value = GetPart1Data.fields.getValue("Prev6Month3")
Prev6Month4Value = GetPart1Data.fields.getValue("Prev6Month4")
Prev6Month5Value = GetPart1Data.fields.getValue("Prev6Month5")
Prev6Month6Value = GetPart1Data.fields.getValue("Prev6Month6")
ProjGrowth16Month1Value = GetPart1Data.fields.getValue("ProjGrowth16Month1")
ProjGrowth16Month2Value = GetPart1Data.fields.getValue("ProjGrowth16Month2")
ProjGrowth16Month3Value = GetPart1Data.fields.getValue("ProjGrowth16Month3")
ProjGrowth16Month4Value = GetPart1Data.fields.getValue("ProjGrowth16Month4")
ProjGrowth16Month5Value = GetPart1Data.fields.getValue("ProjGrowth16Month5")
ProjGrowth16Month6Value = GetPart1Data.fields.getValue("ProjGrowth16Month6")
ProjGrowth712Month1Value = GetPart1Data.fields.getValue("ProjGrowth712Month1")
ProjGrowth712Month2Value = GetPart1Data.fields.getValue("ProjGrowth712Month2")
ProjGrowth712Month3Value = GetPart1Data.fields.getValue("ProjGrowth712Month3")
ProjGrowth712Month4Value = GetPart1Data.fields.getValue("ProjGrowth712Month4")
ProjGrowth712Month5Value = GetPart1Data.fields.getValue("ProjGrowth712Month5")
ProjGrowth712Month6Value = GetPart1Data.fields.getValue("ProjGrowth712Month6")
AvgMonGrowthRateValue = GetPart1Data.fields.getValue("AvgMonGrowthRate")
MonthsToExhaustValue = GetPart1Data.fields.getValue("MonthsToExhaust")
AppendixBExplanationValue = GetPart1Data.fields.getValue("AppendixBExplanation")

''sql3 gets the p3 recordset of the P3 request using the Tix from the input form.
sql3="Select * from xca_Part3 where Tix= '"&Tix&"'"
	GetPart3Data.setSQLText(sql3)
	GetPart3Data.Open

Part3ResultsValue=GetPart3Data.fields.getValue("Part3Result")
Select Case Part3ResultsValue
Case "a"
	Part3ResultsChar1="**"
Case "r"
	Part3ResultsChar2="**"
Case "u"
	Part3ResultsChar3="**"
Case "i"
	Part3ResultsChar4="**"
Case "d"
	Part3ResultsChar5="**"
Case "s"
	Part3ResultsChar6="**"
End Select


sql6="SELECT xca_Part3.*, xca_Part1.SwitchID, xca_Part1.RateCenter FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE xca_Part3.Tix = '"&Tix&"' AND xca_Part3.Part3Result = 'a'"
	CodeAssignRec.setSQLText(sql6)
	CodeAssignRec.Open

RRCompleteValue=CodeAssignRec.fields.getValue("RRComplete")
Select Case RRCompleteValue
Case "Y"
	RRCompletechar="YES"
Case "N"
	RRCompletechar="NO"
End Select


CNAResponsibleValue=CodeAssignRec.fields.getValue("CNAResponsible")
Select Case CNAResponsibleValue
Case "Y"
	CNAResponsiblechar1="IS"
Case "N"
	CNAResponsiblechar1="IS NOT"
end Select



sql7="SELECT xca_Part3.*, xca_Part1.SwitchID FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE xca_Part3.Tix = '"&Tix&"' AND xca_Part3.Part3Result = 'r'"
	CodeResvRec.setSQLText(sql7)
	CodeResvRec.Open

ExtDateValue=CodeResvRec.fields.getValue("ExtentionDate")
If ExtDateValue="1/1/1900" Then
	ExtDateValue=""
Else
	ExtDateValue="'"&ExtentionDate&"'"
End If


function GetP4DAYS()

	dim GetP4DAYSTemp
	dim objConn
	dim objCmd
	dim objRec

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	objCmd.CommandText="Select Value From xca_Parms where Name='P4DAYS'"
	set objRec=objCmd.Execute
	if not objRec.EOF then
		GetP4DAYSTemp=objRec("Value")
	else
		GetP4DAYSTemp=""					
	end if	
				
	objRec.close
	objConn.close
	
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	
	GetP4DAYS=GetP4DAYSTemp
	
end function



'Response.Write Tix

'sql4 gets the p4 recordset of the P4 request using the Tix from the input form.
'sql4="SELECT * FROM xca_Entity, xca_User, xca_Part4 WHERE xca_Part4.Tix = '"&Tix&"' AND xca_User.UserLogon = '"&&"'AND xca_Entity.EntityID = ?
sql4="Select * from xca_Part4 where Tix= '"&Tix&"'"
	GetPart4Data.setSQLText(sql4)
	GetPart4Data.Open
	P4Tix=GetPart4Data.fields.getValue("Tix")
	'Response.Write "|--|P4Tix is =" & P4Tix
	Signature=GetPart4Data.fields.getValue("Signature")
	'Response.Write "|--|Signature is =" & Signature
	P4UserEntityID=GetPart4Data.fields.getValue("EntityID")
	'Response.Write "|--|P4UserEntityID is =" & P4UserEntityID

'sql33="Select * from xca_User, xca_Entity where xca_User.UserLogon = '"&Signature&"'"
'	GetAdminUserName.setSQLText(sql33)
'	GetAdminUserName.open
'	P4Signature=GetAdminUserName.fields.getValue("UserLogon")
'	Response.Write P4Signature
'	P4UserID=GetAdminUserName.fields.getValue("UserID")
'	Response.write "P4UserID is " & P4UserID
'	Response.write "P4UserEntityID is " & P4UserEntityID

''''''''''''get app info for bottom of form
sql = "SELECT * FROM xca_Entity, xca_User, xca_Part4 WHERE xca_Part4.Tix = '"&P4Tix&"' and xca_User.UserLogon = '"&Signature&"' and xca_Entity.EntityID = '"&P4UserEntityID&"'"'"
'sql = "SELECT * FROM xca_Entity, xca_User, xca_Part4 WHERE xca_Entity.EntityID = '"&UserEntityID&"' AND xca_User.UserID = '"&session("UserUserID")&"' AND xca_Part4.Signature = '"&session("pSignature")&"' AND xca_Part4.Tix = '"&session("Tix")&"'"
'sql = "Select * from xca_Entity, xca_User, xca_Part3 Where xca_User.UserID = '"&session("UserUserID")&"' and xca_Entity.EntityID = '"&UserEntityID&"' and xca_Part3.AssignedNPA = '"&session("pPart4NPA")&"' and xca_Part3.AssignedNXX = '"&session("pPart4NXX")&"'"
'sql = "SELECT * FROM xca_Part4 INNER JOIN xca_User ON xca_Part4.Signature = xca_User.UserLogon, xca_Entity WHERE xca_Part4.Tix = '"&P4Tix&"' AND xca_Part4.Signature = '"&Signature&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
	GetP4UserEntityName.setSQLText(sql)
	GetP4UserEntityName.Open
	P4EntityName=GetP4UserEntityName.fields.getValue("EntityName")
	P4UserName=GetP4UserEntityName.fields.getValue("UserName")
	P4EntityAddress=GetP4UserEntityName.fields.getValue("EntityAddress")
	P4EntityCity=GetP4UserEntityName.fields.getValue("EntityCity")
	P4EntityProvince=GetP4UserEntityName.fields.getValue("EntityProvince")
	P4EntityPostalCode=GetP4UserEntityName.fields.getValue("EntityPostalCode")            
	P4UserEmail=GetP4UserEntityName.fields.getValue("UserEmail")
	P4UserFax=GetP4UserEntityName.fields.getValue("UserFax")
	P4UserTelephone=GetP4UserEntityName.fields.getValue("UserTelephone")
	P4UserExtension=GetP4UserEntityName.fields.getValue("UserExtension")

	'Response.Write "P4EntityAddress is " & P4EntityAddress
	'Response.Write "P4EntityName is " & P4EntityName
	'Response.Write "Signature is " & Signature

elseif RequestStatusValue="CU" or RequestStatusValue="CD" or RequestStatusValue="CI" or RequestStatusValue="CP" or RequestStatusValue="CA" or RequestStatusValue="CC" then

''sql gets user entity recordset of user
'sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserName = '"&UserName&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserID = '"&P1UserID&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
'sql = "Select * from xca_Entity where EntityID = '"&UserEntityID&"'"
	GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open
''''''This sql statement is getting the correct information from Tix, P1UserI, and UserEntityID.
	'Response.Write "P1UserID is " & P1UserID
	'Response.Write "UserEntityID is " & UserEntityID
	Tix=GetUserEntityName.fields.getValue("Tix")
	'Response.Write "Tix is " & Tix
	AppEntityName=GetUserEntityName.fields.getValue("EntityName")
	'Response.Write "AppEntityName is " & AppEntityName
	UserName=GetUserEntityName.fields.getValue("UserName")
	'Response.Write "UserName is " &UserName
	UserEmail=GetUserEntityName.fields.getValue("UserEmail")
	'Response.Write "UserEmail is " & UserEmail
	UserFax=GetUserEntityName.fields.getValue("UserFax")
	'Response.Write "UserFax is " & UserFax
	UserTelephone=GetUserEntityName.fields.getValue("UserTelephone")
	'Response.Write "UserTelephone is " & UserTelephone
	UserExtension=GetUserEntityName.fields.getValue("UserExtension")
	'Response.Write "UserExtension is " & UserExtension
	EntityPostalCode=GetUserEntityName.fields.getValue("EntityPostalCode")
	'Response.Write "EntityPostalCode is " & EntityPostalCode
	EntityAddress=GetUserEntityName.fields.getValue("EntityAddress")
	EntityCity=GetUserEntityName.fields.getValue("EntityCity")
	EntityProvince=GetUserEntityName.fields.getValue("EntityProvince")



sqlrsv="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'R'"
	ToRRsvRec.setSQLText(sqlrsv)
	ToRRsvRec.open

sqlass="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'A'"
	ToRAssRec.setSQLText(sqlass)
	ToRAssRec.open


''sql3 gets the NPASplitID of the NPA-preferred NXX that was requested from the COCode table
sql3="SELECT  xca_COCode.NPASplitID, xca_Part1.NXX1preferred FROM xca_COCode FULL OUTER JOIN xca_Part1 ON xca_COCode.Tix = xca_Part1.Tix WHERE xca_Part1.Tix = '"&Tix&"'"
	GetCOCodeData.setSQLText(sql3)
	GetCOCodeData.Open
	NPASplitIDValue=GetCOCodeData.fields.getValue("NPASplitID")
	Select Case NPASplitIDValue
	Case "Included"
		NPASplitIDValue1="Included"
	Case "Excluded"
		NPASplitIDValue1="Excluded"
	End Select



CodeReqNew=GetPart1Data.fields.getValue("CodeRequestNew")
Select Case CodeReqNew
Case "c"
	CodeReqNewchar1="**"
Case "o"
	CodeReqNewchar2="**"
End Select


CertReqValue=GetPart1Data.fields.getValue("CertificationRequired")
Select Case CertReqValue
Case "Y"
	CertReqchar="YES"
Case "N"
	CertReqchar="NO"
End Select


ReqCertReadyValue=GetPart1Data.fields.getValue("RequiredCertificationReady")
Select Case ReqCertReadyValue
Case "Y"
	ReqCertReadychar="YES"
Case "N"
	ReqCertReadychar="NO"
End Select


TypeEntityValue=GetPart1Data.fields.getValue("CarrierType")
Select Case TypeEntityValue
Case "l"
	TypeEntitychar="Local Exchange Carrier"
Case "w"
	TypeEntitychar="Wireless Service Provider"
Case "o"
	TypeEntitychar="Other(Specify)"
End Select


AuthPart2Value=GetPart1Data.fields.getValue("AuthorizationPart2")
Select Case AuthPart2Value
Case "y"
	AuthPart2char1="**"
Case "n"
	AuthPart2char2="**"
End Select


TyReqvalue=GetPart1Data.fields.getValue("TypeOfRequest")
Select Case TyReqvalue
Case "A"
	TyReqchar1="**"
Case "U"
	TyReqchar2="**"
Case "R"
	TyReqchar3="**"
End Select


Reas4ReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case Reas4ReqValue
Case "aic"
	Reas4Reqchar="a) Initial Code for new Switching Entity or new Point of Interconnection (Complete Part 2)"
Case "aau"
	Reas4Reqchar="b) Code request for New Application for existing switching entity or point of interconnection (Code Aplicant must complete Section 1.7)"
Case "aag"
	Reas4Reqchar="c) Additional Code for Growth (Code Applicant must complete Section 1.6)"
End Select


ReasForReqValue=GetPart1Data.fields.getValue("ReasonForRequest")
Select Case ReasForReqValue
Case "ric"
	ReasForReqchar="a) Initial Code"
Case "rau"
	ReasForReqchar="b) New Application (Complete Section 1.7)"
Case "rag"
	ReasForReqchar="c) Growth (Complete Section 1.6)"
End Select


JeopardyValue = GetPart1Data.fields.getValue("NPAinJeopardy")
select case JeopardyValue
case "y"
	JeopardyName1="YES"
	P3Jeopardychar="YES"
case "n"
	JeopardyName2="NO"
	P3Jeopardychar="NO"
case else
	P3Jeopardychar="NO"
end select


NXXGrowthCalValue = GetPart1Data.fields.getValue("NXXGrowthCal")
TNsValue = GetPart1Data.fields.getValue("TNs")
Prev6Month1Value = GetPart1Data.fields.getValue("Prev6Month1")
Prev6Month2Value = GetPart1Data.fields.getValue("Prev6Month2")
Prev6Month3Value = GetPart1Data.fields.getValue("Prev6Month3")
Prev6Month4Value = GetPart1Data.fields.getValue("Prev6Month4")
Prev6Month5Value = GetPart1Data.fields.getValue("Prev6Month5")
Prev6Month6Value = GetPart1Data.fields.getValue("Prev6Month6")
ProjGrowth16Month1Value = GetPart1Data.fields.getValue("ProjGrowth16Month1")
ProjGrowth16Month2Value = GetPart1Data.fields.getValue("ProjGrowth16Month2")
ProjGrowth16Month3Value = GetPart1Data.fields.getValue("ProjGrowth16Month3")
ProjGrowth16Month4Value = GetPart1Data.fields.getValue("ProjGrowth16Month4")
ProjGrowth16Month5Value = GetPart1Data.fields.getValue("ProjGrowth16Month5")
ProjGrowth16Month6Value = GetPart1Data.fields.getValue("ProjGrowth16Month6")
ProjGrowth712Month1Value = GetPart1Data.fields.getValue("ProjGrowth712Month1")
ProjGrowth712Month2Value = GetPart1Data.fields.getValue("ProjGrowth712Month2")
ProjGrowth712Month3Value = GetPart1Data.fields.getValue("ProjGrowth712Month3")
ProjGrowth712Month4Value = GetPart1Data.fields.getValue("ProjGrowth712Month4")
ProjGrowth712Month5Value = GetPart1Data.fields.getValue("ProjGrowth712Month5")
ProjGrowth712Month6Value = GetPart1Data.fields.getValue("ProjGrowth712Month6")
AvgMonGrowthRateValue = GetPart1Data.fields.getValue("AvgMonGrowthRate")
MonthsToExhaustValue = GetPart1Data.fields.getValue("MonthsToExhaust")
AppendixBExplanationValue = GetPart1Data.fields.getValue("AppendixBExplanation")

''sql3 gets the p3 recordset of the P3 request using the Tix from the input form.
sql3="Select * from xca_Part3 where Tix= '"&Tix&"'"
	GetPart3Data.setSQLText(sql3)
	GetPart3Data.Open

Part3ResultsValue=GetPart3Data.fields.getValue("Part3Result")
Select Case Part3ResultsValue
Case "a"
	Part3ResultsChar1="**"
Case "r"
	Part3ResultsChar2="**"
Case "u"
	Part3ResultsChar3="**"
Case "i"
	Part3ResultsChar4="**"
Case "d"
	Part3ResultsChar5="**"
Case "s"
	Part3ResultsChar6="**"
End Select


sql6="SELECT xca_Part3.*, xca_Part1.SwitchID, xca_Part1.RateCenter FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE xca_Part3.Tix = '"&Tix&"' AND xca_Part3.Part3Result = 'a'"
	CodeAssignRec.setSQLText(sql6)
	CodeAssignRec.Open

RRCompleteValue=CodeAssignRec.fields.getValue("RRComplete")
Select Case RRCompleteValue
Case "Y"
	RRCompletechar="YES"
Case "N"
	RRCompletechar="NO"
End Select


CNAResponsibleValue=CodeAssignRec.fields.getValue("CNAResponsible")
Select Case CNAResponsibleValue
Case "Y"
	CNAResponsiblechar1="IS"
Case "N"
	CNAResponsiblechar1="IS NOT"
end Select



sql7="SELECT xca_Part3.*, xca_Part1.SwitchID FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE xca_Part3.Tix = '"&Tix&"' AND xca_Part3.Part3Result = 'r'"
	CodeResvRec.setSQLText(sql7)
	CodeResvRec.Open

ExtDateValue=CodeResvRec.fields.getValue("ExtentionDate")
If ExtDateValue="1/1/1900" Then
	ExtDateValue=""
Else
	ExtDateValue="'"&ExtentionDate&"'"
End If

end if



	

%>	

<body bgColor="#d7c7a4" bgProperties="fixed" text="black" leftmargin=20 rightmargin=20>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserID\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserID\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityName()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = ? AND xca_User.UserID = ? AND xca_Entity.EntityID = ?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetUserEntityName.setRecordSource(rsTmp);
}
function _GetUserEntityName_ctor()
{
	CreateRecordset('GetUserEntityName', _initGetUserEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q35\q,CReq=0)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetAdminEntityName()
{
}
function _initGetAdminEntityName()
{
	GetAdminEntityName.advise(RS_ONBEFOREOPEN, _setParametersGetAdminEntityName);
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_Parms WHERE (xca_Entity.EntityName = ?)';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetAdminEntityName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetAdminEntityName') != null)
		GetAdminEntityName.setBookmark(thisPage.getState('pb_GetAdminEntityName'));
}
function _GetAdminEntityName_ctor()
{
	CreateRecordset('GetAdminEntityName', _initGetAdminEntityName, null);
}
function _GetAdminEntityName_dtor()
{
	GetAdminEntityName._preserveState();
	thisPage.setState('pb_GetAdminEntityName', GetAdminEntityName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminUserName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_User\sWHERE\sUserLogon\s=\s?\q,TCControlID_Unmatched=\qGetAdminUserName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_User\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_User\sWHERE\sUserLogon\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminUserName()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_User WHERE UserLogon = ?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetAdminUserName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetAdminUserName') != null)
		GetAdminUserName.setBookmark(thisPage.getState('pb_GetAdminUserName'));
}
function _GetAdminUserName_ctor()
{
	CreateRecordset('GetAdminUserName', _initGetAdminUserName, null);
}
function _GetAdminUserName_dtor()
{
	GetAdminUserName._preserveState();
	thisPage.setState('pb_GetAdminUserName', GetAdminUserName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1Data 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxca_Part1.Tix\s=?\q,TCControlID_Unmatched=\qGetPart1Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxca_Part1.Tix\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetPart1Data()
{
}
function _initGetPart1Data()
{
	GetPart1Data.advise(RS_ONBEFOREOPEN, _setParametersGetPart1Data);
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
	cmdTmp.CommandText = 'Select * from xca_part1 where xca_Part1.Tix =?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart1Data.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPart1Data') != null)
		GetPart1Data.setBookmark(thisPage.getState('pb_GetPart1Data'));
}
function _GetPart1Data_ctor()
{
	CreateRecordset('GetPart1Data', _initGetPart1Data, null);
}
function _GetPart1Data_dtor()
{
	GetPart1Data._preserveState();
	thisPage.setState('pb_GetPart1Data', GetPart1Data.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart3Data 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Part3\swhere\sTix=?\q,TCControlID_Unmatched=\qGetPart3Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Part3\swhere\sTix=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetPart3Data()
{
}
function _initGetPart3Data()
{
	GetPart3Data.advise(RS_ONBEFOREOPEN, _setParametersGetPart3Data);
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
	cmdTmp.CommandText = 'Select * from xca_Part3 where Tix=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart3Data.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPart3Data') != null)
		GetPart3Data.setBookmark(thisPage.getState('pb_GetPart3Data'));
}
function _GetPart3Data_ctor()
{
	CreateRecordset('GetPart3Data', _initGetPart3Data, null);
}
function _GetPart3Data_dtor()
{
	GetPart3Data._preserveState();
	thisPage.setState('pb_GetPart3Data', GetPart3Data.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetCOCodeData 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Part1.NXX1preferred,\sxca_COCode.NPASplitID\sFROM\sxca_COCode\sFULL\sOUTER\sJOIN\sxca_Part1\sON\sxca_COCode.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part1.Tix\s=\s?)\q,TCControlID_Unmatched=\qGetCOCodeData\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Part1.NXX1preferred,\sxca_COCode.NPASplitID\sFROM\sxca_COCode\sFULL\sOUTER\sJOIN\sxca_Part1\sON\sxca_COCode.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part1.Tix\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetCOCodeData()
{
}
function _initGetCOCodeData()
{
	GetCOCodeData.advise(RS_ONBEFOREOPEN, _setParametersGetCOCodeData);
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
	cmdTmp.CommandText = 'SELECT xca_Part1.NXX1preferred, xca_COCode.NPASplitID FROM xca_COCode FULL OUTER JOIN xca_Part1 ON xca_COCode.Tix = xca_Part1.Tix WHERE (xca_Part1.Tix = ?)';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetCOCodeData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetCOCodeData') != null)
		GetCOCodeData.setBookmark(thisPage.getState('pb_GetCOCodeData'));
}
function _GetCOCodeData_ctor()
{
	CreateRecordset('GetCOCodeData', _initGetCOCodeData, null);
}
function _GetCOCodeData_dtor()
{
	GetCOCodeData._preserveState();
	thisPage.setState('pb_GetCOCodeData', GetCOCodeData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart4Data style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part4\sWHERE\s(xca_Part4.Tix\s=\s?)\sAND\s(xca_User.UserLogon\s=\s?)AND\sxca_Entity.EntityID\s=\s?\q,TCControlID_Unmatched=\qGetPart4Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part4\sWHERE\s(xca_Part4.Tix\s=\s?)\sAND\s(xca_User.UserLogon\s=\s?)AND\sxca_Entity.EntityID\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart4Data()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_User, xca_Part4 WHERE (xca_Part4.Tix = ?) AND (xca_User.UserLogon = ?)AND xca_Entity.EntityID = ?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart4Data.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPart4Data') != null)
		GetPart4Data.setBookmark(thisPage.getState('pb_GetPart4Data'));
}
function _GetPart4Data_ctor()
{
	CreateRecordset('GetPart4Data', _initGetPart4Data, null);
}
function _GetPart4Data_dtor()
{
	GetPart4Data._preserveState();
	thisPage.setState('pb_GetPart4Data', GetPart4Data.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=CodeAssignRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID,\sxca_Part1.RateCenter\sAS\sRateCenter\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'a')\q,TCControlID_Unmatched=\qCodeAssignRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID,\sxca_Part1.RateCenter\sAS\sRateCenter\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'a')\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersCodeAssignRec()
{
}
function _initCodeAssignRec()
{
	CodeAssignRec.advise(RS_ONBEFOREOPEN, _setParametersCodeAssignRec);
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
	cmdTmp.CommandText = 'SELECT xca_Part3.*, xca_Part1.SwitchID AS SwitchID, xca_Part1.RateCenter AS RateCenter FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE (xca_Part3.Tix = ?) AND (xca_Part3.Part3Result = \'a\')';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	CodeAssignRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_CodeAssignRec') != null)
		CodeAssignRec.setBookmark(thisPage.getState('pb_CodeAssignRec'));
}
function _CodeAssignRec_ctor()
{
	CreateRecordset('CodeAssignRec', _initCodeAssignRec, null);
}
function _CodeAssignRec_dtor()
{
	CodeAssignRec._preserveState();
	thisPage.setState('pb_CodeAssignRec', CodeAssignRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=CodeResvRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'r')\q,TCControlID_Unmatched=\qCodeResvRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'r')\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersCodeResvRec()
{
}
function _initCodeResvRec()
{
	CodeResvRec.advise(RS_ONBEFOREOPEN, _setParametersCodeResvRec);
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
	cmdTmp.CommandText = 'SELECT xca_Part3.*, xca_Part1.SwitchID AS SwitchID FROM xca_Part3 INNER JOIN xca_Part1 ON xca_Part3.Tix = xca_Part1.Tix WHERE (xca_Part3.Tix = ?) AND (xca_Part3.Part3Result = \'r\')';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	CodeResvRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_CodeResvRec') != null)
		CodeResvRec.setBookmark(thisPage.getState('pb_CodeResvRec'));
}
function _CodeResvRec_ctor()
{
	CreateRecordset('CodeResvRec', _initCodeResvRec, null);
}
function _CodeResvRec_dtor()
{
	CodeResvRec._preserveState();
	thisPage.setState('pb_CodeResvRec', CodeResvRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ToRAssRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'A'\q,TCControlID_Unmatched=\qToRAssRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'A'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initToRAssRec()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Part1 WHERE TypeOfRequest = \'A\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ToRAssRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ToRAssRec') != null)
		ToRAssRec.setBookmark(thisPage.getState('pb_ToRAssRec'));
}
function _ToRAssRec_ctor()
{
	CreateRecordset('ToRAssRec', _initToRAssRec, null);
}
function _ToRAssRec_dtor()
{
	ToRAssRec._preserveState();
	thisPage.setState('pb_ToRAssRec', ToRAssRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 
id=ToRRsvRec style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" width=461>
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'R'\q,TCControlID_Unmatched=\qToRRsvRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part1\sWHERE\sTypeOfRequest\s=\s'R'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initToRRsvRec()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Part1 WHERE TypeOfRequest = \'R\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ToRRsvRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ToRRsvRec') != null)
		ToRRsvRec.setBookmark(thisPage.getState('pb_ToRRsvRec'));
}
function _ToRRsvRec_ctor()
{
	CreateRecordset('ToRRsvRec', _initToRRsvRec, null);
}
function _ToRRsvRec_dtor()
{
	ToRRsvRec._preserveState();
	thisPage.setState('pb_ToRRsvRec', ToRRsvRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 id=GetP4UserEntityName 
	style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part4\sINNER\sJOIN\sxca_User\sON\sxca_Part4.Signature\s=\sxca_User.UserLogon,\sxca_Entity\sWHERE\s(xca_Part4.Tix\s=\s?)\sAND\s(xca_Part4.Signature\s=\s?)\sAND\sxca_Entity.EntityID\s=\s?\q,TCControlID_Unmatched=\qGetP4UserEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part4\sINNER\sJOIN\sxca_User\sON\sxca_Part4.Signature\s=\sxca_User.UserLogon,\sxca_Entity\sWHERE\s(xca_Part4.Tix\s=\s?)\sAND\s(xca_Part4.Signature\s=\s?)\sAND\sxca_Entity.EntityID\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetP4UserEntityName()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Part4 INNER JOIN xca_User ON xca_Part4.Signature = xca_User.UserLogon, xca_Entity WHERE (xca_Part4.Tix = ?) AND (xca_Part4.Signature = ?) AND xca_Entity.EntityID = ?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetP4UserEntityName.setRecordSource(rsTmp);
}
function _GetP4UserEntityName_ctor()
{
	CreateRecordset('GetP4UserEntityName', _initGetP4UserEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<table align=center border="0" cellpadding="2"><tr>
  <td nowrap><font color="maroon" face="Arial Black" size="5"><strong>
	Request Forms 
            Report: Part1, Part3, Part4</strong></font> 
    
		</td>
	</tr>
</table>

<p>&nbsp;</p>

<table align="left" border="0" cellPadding="0" cellSpacing="0" width="174" background ="" height="48" style  ="HEIGHT: 48px; WIDTH: 174px">
    
    <tr>
        <td align="left" noWrap>&nbsp;
        <td align="left" nowrap><strong><font face="Arial" color="maroon" size="4">
	CNA 
            Ticket #:&nbsp;&nbsp;
            <% Response.Write(Tix) %></font></strong></td>
		</td>
	</tr>
	<TR>
		<td align="left" noWrap>&nbsp;
		<td align="left" noWrap><font color="maroon" size="4"><strong>
	Request 
            Status:&nbsp;&nbsp;
            <% Response.write RequestStatuschar %></strong></font>
		</td>
	</TR>
</table>&nbsp;

<TABLE align=right border=0 cellPadding=1 cellSpacing=1 height=29 
style="HEIGHT: 29px; WIDTH: 310px" width=32.56%>
    <TR>
        <TD align=right><font face="Arial" size="4" color="maroon"><strong>Created:</strong></font>
        </TD>
        <TD align=left><font face="Arial" size="4" color="maroon">
            <% Response.write "" & Date() %></font>
        </TD>
	</TR>
</TABLE>

<p>&nbsp;</p>

<table border="0" cellpadding="0"><tr>
        <td noWrap>
	<td><font color="maroon" face="Arial" size="4"><strong>
Part 1 - Canadian 
            Central Office Code (NXX) Assignment Request Form</strong></font> 
    
		</td>
	</tr>
</table>
<font face=arial size=2>

<p>Please complete the following form. Use one form per NXX 
code request. Mail, fax, or submit online the completed form to the Code 
Administrator.</p>
<p>The Code Applicants are granted subject to the condition 
that all code holders are subject to the assignment guidelines which are 
published and available from the appropriate Code Administrator. A code assigned 
to an entity, either directly by the Code Administrator or through transfer from 
another entity, should be placed in service within 6 months after the initially 
published effective date.</p>
<p>These guidelines may be modified from time-to-time. The 
assignment guidelines in effect shall apply equally to all Code Applicants and 
all existing code holders.</p> 
<p>The Code Applicant and the Code Administrator acknowledge 
that the information contained on this request form is sensitive and will be 
treated as confidential. Prior to confirmation the information in this form will 
only be shared with the appropriate administrator and/or regulators. Information 
requested for RDBS and BRIDS will become available to the public upon input into 
those systems.</p>
<p>I hereby certify that the following information 
requesting an NXX code is true and accurate to the best of my knowledge and that 
this application has been prepared in accordance with the Canadian Central 
Office Code (NXX) Assignment Guidelines dated October 23, 1997 which were 
adopted by the CSCN on April 2, 1998.</p>
<p>It is understood that the Code Applicant will return the 
CO Code to the administrator for reassignment if the resource is no longer in 
use by the Code Applicant, no longer required for the service for which it was 
intended, not activated within the time frame specified in these guidelines (an 
extension can be applied for), or not used in conformance with these assignment 
guidelines.</p></font>
<p>
<br>
<table align="left" border="0" cellPadding="0" cellSpacing="0">
<tr>
<td wrap>
<strong><font   
            size="2" face=arial><strong>Code Applicants are required to retain a copy of all 
            application forms, appendices and supporting data in the event of an 
            audit.</strong></font>
            </strong></td></tr>
</table>
<br>
<br>
<br>

<table align=center border="0" cellPadding="0" cellSpacing="0">
<tr>
<td align="right" wrap><label><font face=arial size="2"><strong>Authorized Representative 
            Name:&nbsp;&nbsp;</strong></font></label></td>
<td align="left" wrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AuthorizedRep 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 98px" width=98>
	<PARAM NAME="_ExtentX" VALUE="2593">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AuthorizedRep">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRep">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRep()
{
	AuthorizedRep.setDataSource(GetPart1Data);
	AuthorizedRep.setDataField('AuthorizedRep');
}
function _AuthorizedRep_ctor()
{
	CreateLabel('AuthorizedRep', _initAuthorizedRep, null);
}
</script>
<% AuthorizedRep.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
<tr>
<td align="right" wrap><label><font face=arial size="2"><strong>Title:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AuthorizedRepTitle 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 126px" width=126>
	<PARAM NAME="_ExtentX" VALUE="3334">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AuthorizedRepTitle">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="AuthorizedRepTitle">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRepTitle()
{
	AuthorizedRepTitle.setDataSource(GetPart1Data);
	AuthorizedRepTitle.setDataField('AuthorizedRepTitle');
}
function _AuthorizedRepTitle_ctor()
{
	CreateLabel('AuthorizedRepTitle', _initAuthorizedRepTitle, null);
}
</script>
<% AuthorizedRepTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
<tr>
<td align="right" wrap><label><font face=arial size="2"><strong>Date of 
            Receipt:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=DateofReceipt1 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofReceipt1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofReceipt1()
{
	DateofReceipt1.setDataSource(GetPart1Data);
	DateofReceipt1.setDataField('DateofReceipt');
}
function _DateofReceipt1_ctor()
{
	CreateLabel('DateofReceipt1', _initDateofReceipt1, null);
}
</script>
<% DateofReceipt1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
</table>
<br><br>
<br><br>
<strong><center><font  size="4" face=arial color="#993300">General Information</font></strong></CENTER>
<table align="left" border="0" cellPadding="0" cellSpacing="1">
<tr>
        <td wrap style="FONT-WEIGHT: bold"><label><strong><font    
            size="3" face=arial color="#993300">1.1 Contact 
            Information:</font></strong></label> 
 
 </td></tr>
 
 </table>
 <br>
 <br>


<table align="center" border="0" cellPadding="1" cellSpacing="1" >
    <tbody>
    
    <tr>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code Applicant 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" wrap><font face="Arial"> </font>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Company</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <% Response.Write AppEntityName%></font></STRONG>
		</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Entity Name </STRONG>
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityName 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1042px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityName">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityName()
{
	AdminEntityName.setDataSource(GetAdminEntityName);
	AdminEntityName.setDataField('EntityName');
}
function _AdminEntityName_ctor()
{
	CreateLabel('AdminEntityName', _initAdminEntityName, null);
}
</script>
<% AdminEntityName.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>

</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Contact 
            Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <% Response.Write UserName%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Contact 
            Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityContact 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1062px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityContact()
{
	AdminEntityContact.setDataSource(GetAdminEntityName);
	AdminEntityContact.setDataField('EntityContact');
}
function _AdminEntityContact_ctor()
{
	CreateLabel('AdminEntityContact', _initAdminEntityContact, null);
}
</script>
<% AdminEntityContact.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Street 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write EntityAddress%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Street 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityAddress 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1082px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityAddress()
{
	AdminEntityAddress.setDataSource(GetAdminEntityName);
	AdminEntityAddress.setDataField('EntityAddress');
}
function _AdminEntityAddress_ctor()
{
	CreateLabel('AdminEntityAddress', _initAdminEntityAddress, null);
}
</script>
<% AdminEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>City</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write EntityCity%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>City</STRONG> 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityCity 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1102px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityCity()
{
	AdminEntityCity.setDataSource(GetAdminEntityName);
	AdminEntityCity.setDataField('EntityCity');
}
function _AdminEntityCity_ctor()
{
	CreateLabel('AdminEntityCity', _initAdminEntityCity, null);
}
</script>
<% AdminEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Province</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write EntityProvince%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Province</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityProvince 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1122px; WIDTH: 95px" width=95>
	<PARAM NAME="_ExtentX" VALUE="2514">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityProvince()
{
	AdminEntityProvince.setDataSource(GetAdminEntityName);
	AdminEntityProvince.setDataField('EntityProvince');
}
function _AdminEntityProvince_ctor()
{
	CreateLabel('AdminEntityProvince', _initAdminEntityProvince, null);
}
</script>
<% AdminEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
         
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Postal 
            Code</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write EntityPostalCode%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Postal Code</STRONG> 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityPostalCode 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1142px; WIDTH: 111px" width=111>
	<PARAM NAME="_ExtentX" VALUE="2937">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityPostalCode()
{
	AdminEntityPostalCode.setDataSource(GetAdminEntityName);
	AdminEntityPostalCode.setDataField('EntityPostalCode');
}
function _AdminEntityPostalCode_ctor()
{
	CreateLabel('AdminEntityPostalCode', _initAdminEntityPostalCode, null);
}
</script>
<% AdminEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
           
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>E-Mail Address</STRONG> 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <% Response.Write UserEmail%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>E-Mail 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityEmail 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1162px; WIDTH: 75px" width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityEmail()
{
	AdminEntityEmail.setDataSource(GetAdminEntityName);
	AdminEntityEmail.setDataField('EntityEmail');
}
function _AdminEntityEmail_ctor()
{
	CreateLabel('AdminEntityEmail', _initAdminEntityEmail, null);
}
</script>
<% AdminEntityEmail.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>    
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Facsimile</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write UserFax%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Facsimile</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityFax 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1182px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityFax()
{
	AdminEntityFax.setDataSource(GetAdminEntityName);
	AdminEntityFax.setDataField('EntityFax');
}
function _AdminEntityFax_ctor()
{
	CreateLabel('AdminEntityFax', _initAdminEntityFax, null);
}
</script>
<% AdminEntityFax.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
          
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Telephone</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <%Response.Write UserTelephone%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Telephone</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityTelephone 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1202px; WIDTH: 107px" width=107>
	<PARAM NAME="_ExtentX" VALUE="2831">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityTelephone()
{
	AdminEntityTelephone.setDataSource(GetAdminEntityName);
	AdminEntityTelephone.setDataField('EntityTelephone');
}
function _AdminEntityTelephone_ctor()
{
	CreateLabel('AdminEntityTelephone', _initAdminEntityTelephone, null);
}
</script>
<% AdminEntityTelephone.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
            
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><STRONG>
        <% Response.Write UserExtension%></font></STRONG>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityExtension 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1222px; WIDTH: 101px" width=101>
	<PARAM NAME="_ExtentX" VALUE="2672">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityExtension()
{
	AdminEntityExtension.setDataSource(GetAdminEntityName);
	AdminEntityExtension.setDataField('EntityExtension');
}
function _AdminEntityExtension_ctor()
{
	CreateLabel('AdminEntityExtension', _initAdminEntityExtension, null);
}
</script>
<% AdminEntityExtension.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
           
</td></tr></tbody>

</table>


<br><br>

<table align="left" border="0" cellPadding="0" cellSpacing="0">
    
    <TR>
        <TD align=left colSpan=8><strong><font face=arial color="#993300" size="3">
	1.2 
            CO Code Information:</font></strong>
    <TR>
        <TD align=right colSpan=8>
            <DIV align=left>&nbsp; </DIV>
    <TR>
        <TD align=left  colSpan=2 width=100><strong><font  face=arial size="2">&nbsp;NPA:&nbsp;</font></strong></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 31px" 
            width=31>
	<PARAM NAME="_ExtentX" VALUE="820">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPA()
{
	NPA.setDataSource(GetPart1Data);
	NPA.setDataField('NPA');
}
function _NPA_ctor()
{
	CreateLabel('NPA', _initNPA, null);
}
</script>
<% NPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD align=left colSpan=2 width=100><strong><font   
            face=arial size="2">&nbsp;LATA:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=LATA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="LATA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="LATA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLATA()
{
	LATA.setDataSource(GetPart1Data);
	LATA.setDataField('LATA');
}
function _LATA_ctor()
{
	CreateLabel('LATA', _initLATA, null);
}
</script>
<% LATA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
        <TD align=left colSpan=4 width=100><strong><font   
            face=arial size="2">&nbsp;OCN:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=OCN style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 32px" 
            width=32>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="OCN">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OCN">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOCN()
{
	OCN.setDataSource(GetPart1Data);
	OCN.setDataField('OCN');
}
function _OCN_ctor()
{
	CreateLabel('OCN', _initOCN, null);
}
</script>
<% OCN.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            

    <TR>
        <TD align=left colSpan=7><strong><font   
            face=arial size="2">Switch 
            Identification (Switching Entity / POI):&nbsp;&nbsp;</strong></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=SwitchID style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
            width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="SwitchID">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSwitchID()
{
	SwitchID.setDataSource(GetPart1Data);
	SwitchID.setDataField('SwitchID');
}
function _SwitchID_ctor()
{
	CreateLabel('SwitchID', _initSwitchID, null);
}
</script>
<% SwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
    <TR>
        <TD align=left colSpan=5>
        <TD align=left colSpan=2><font face="Arial" size="2">This is an eleven-character descriptor of the 
            switch provided by the owning entity for the purpose of routing 
            calls. This is the 11 character COMMON LANGUAGE Location 
            Identification - (CLLI) of the switch or POI.</font>
    <TR>
        <TD align=left colSpan=7><strong><font   
            face=arial size="2">
	City or Wire 
            Center:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=WireCenter style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 76px" 
            width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="WireCenter">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="WireCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initWireCenter()
{
	WireCenter.setDataSource(GetPart1Data);
	WireCenter.setDataField('WireCenter');
}
function _WireCenter_ctor()
{
	CreateLabel('WireCenter', _initWireCenter, null);
}
</script>
<% WireCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
    <TR>
        <TD align=left colSpan=7><strong><font   
            face=arial size="2">Rate 
            Center:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RateCenter style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 75px" 
            width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RateCenter">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RateCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRateCenter()
{
	RateCenter.setDataSource(GetPart1Data);
	RateCenter.setDataField('RateCenter');
}
function _RateCenter_ctor()
{
	CreateLabel('RateCenter', _initRateCenter, null);
}
</script>
<% RateCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font face="Arial" size="2">Rate Center Name must be a 
            tariffed Rate Center associated with toll billing.</font>
    <TR>
        <TD align=left colSpan=7><strong><font  face=arial size="2">Route Same 
            as<strong><font  face=arial size="2">&nbsp;NPA:&nbsp;&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RouteNPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 68px" 
            width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RouteNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNPA()
{
	RouteNPA.setDataSource(GetPart1Data);
	RouteNPA.setDataField('RouteNPA');
}
function _RouteNPA_ctor()
{
	CreateLabel('RouteNPA', _initRouteNPA, null);
}
</script>
<% RouteNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RouteNXX style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 68px" 
            width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RouteNXX">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RouteNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNXX()
{
	RouteNXX.setDataSource(GetPart1Data);
	RouteNXX.setDataField('RouteNXX');
}
function _RouteNXX_ctor()
{
	CreateLabel('RouteNXX', _initRouteNXX, null);
}
</script>
<% RouteNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;<strong><font face="Arial" size="2">Use 
            Same Rate Center as<strong><font face="Arial" size="2">&nbsp;NPA:&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=CenterNPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CenterNPA">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNPA()
{
	CenterNPA.setDataSource(GetPart1Data);
	CenterNPA.setDataField('CenterNPA');
}
function _CenterNPA_ctor()
{
	CreateLabel('CenterNPA', _initCenterNPA, null);
}
</script>
<% CenterNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<strong><font face="Arial" size="2">&nbsp; NXX:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=CenterNXX style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CenterNXX">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CenterNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNXX()
{
	CenterNXX.setDataSource(GetPart1Data);
	CenterNXX.setDataField('CenterNXX');
}
function _CenterNXX_ctor()
{
	CreateLabel('CenterNXX', _initCenterNXX, null);
}
</script>
<% CenterNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font></strong></font></strong></font></strong></font></strong></font> </strong>
    <TR>
        <TD align=left colSpan=7>&nbsp;&nbsp;
    <TR>
        <TD align=left colSpan=7><strong><font  face=arial size="3" color="#993300" style="FONT-WEIGHT: bold">
1.3 Dates:</font></strong>
    <TR>
        <TD align=left colSpan=7>&nbsp;
    <TR>
        <TD align=left colSpan=7><strong><font face="Arial" size="2">Application 
Date:&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ApplicationDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ApplicationDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initApplicationDate()
{
	ApplicationDate.setDataSource(GetPart1Data);
	ApplicationDate.setDataField('ApplicationDate');
}
function _ApplicationDate_ctor()
{
	CreateLabel('ApplicationDate', _initApplicationDate, null);
}
</script>
<% ApplicationDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font  face=arial size=1>dd/mm/ccyy</font></font></strong> 
    <TR>
        <TD align=left colSpan=7><strong><font face="Arial" size="2"><strong>Requested Effective Date:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequestedEffDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 116px" width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestedEffDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestedEffDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedEffDate()
{
	RequestedEffDate.setDataSource(GetPart1Data);
	RequestedEffDate.setDataField('RequestedEffDate');
}
function _RequestedEffDate_ctor()
{
	CreateLabel('RequestedEffDate', _initRequestedEffDate, null);
}
</script>
<% RequestedEffDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
<font  face=arial size=1>dd/mm/ccyy</font> 
            
</strong></font></strong>
    <TR>
        <TD align=left colSpan=7>&nbsp;
    <TR>
        <TD align=left colSpan=7>

<p><font face="Arial" size="2">The nationwide cut-over is a minimum of 45 days after the NXX 
            code request is input to RDBS and BRIDS. To the extent possible, 
            code applicants should avoid requesting an effective date that is an 
            interval less than 66 calendar days from the submission of this 
            form. It should be noted that interconnection arrangements and 
            facilities need to be in place prior to activation of a code. Such 
            arrangements are outside the scope of these guidelines.</font></p>
    <TR>
        <TD align=left colSpan=7>&nbsp;
    <TR>
        <TD align=left colSpan=7>
<p><font face="Arial" size="2">Requests for code assignment should not be made more than 6 
            months prior to the requested effective date.</font></p>
    <TR>
        <TD align=left colSpan=7>&nbsp;
    <TR>
        <TD align=left colSpan=7>
<p><font face="Arial" size="2">Acknowledgment and indication of disposition of this 
            application will be provided to applicant as noted in Section 1.2 
            within ten working days from the date of receipt of this 
            application.</font></p>

<tr>
	<td align="left" wrap colSpan=7>
</td></tr>
</table>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<table align="left" background ="" border="0" cellPadding="0" cellSpacing="0">
    
    <TR>
        <TD align=left colSpan=3><strong><font color=maroon face="Arial" Size=2><font color=maroon face="Arial" Size=2>
	1.4 Type of 
            Entity Requesting the Code:</font></strong> </FONT> 
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
<tr>
<td  wrap align="left" colSpan=3><strong><font face="Arial" size="2"> A)&nbsp;&nbsp;</font><font color=maroon face="Arial" Size=2>
            <% Response.Write TypeEntitychar %></font></strong>&nbsp; 
<strong><font face="Arial" size="2">&nbsp; Other Explained:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=OtherCarrierType 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 114px" width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="OtherCarrierType">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="OtherCarrierType">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOtherCarrierType()
{
	OtherCarrierType.setDataSource(GetPart1Data);
	OtherCarrierType.setDataField('OtherCarrierType');
}
function _OtherCarrierType_ctor()
{
	CreateLabel('OtherCarrierType', _initOtherCarrierType, null);
}
</script>
<% OtherCarrierType.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font></strong>     

</td><tr></tr>
    <TR>
        <TD align=left colSpan=3 vAlign=top>&nbsp;


<tr>
        <TD align=left colSpan=3 vAlign=top><font face=arial size="2"><strong>B)&nbsp; Type of Service for which code is being 
            requested:</strong></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=TypeOfService 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="TypeOfService">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="TypeOfService">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTypeOfService()
{
	TypeOfService.setDataSource(GetPart1Data);
	TypeOfService.setDataField('TypeOfService');
}
function _TypeOfService_ctor()
{
	CreateLabel('TypeOfService', _initTypeOfService, null);
}
</script>
<% TypeOfService.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</TD></tr>
    <TR>
        <TD align=left colSpan=3>&nbsp;


<tr>
<td wrap align="left" colSpan=3><strong><font face="Arial" size="2">C)&nbsp; Is certification or authorization required to provide 
            this type of service in the relevant geographic 
            area?&nbsp;</strong></FONT><font face="Arial" size=2 color=maroon><strong>
            <% Response.Write CertReqchar %></strong></font>
		</td>
	</tr>
	<tr>
<td  wrap width=25></td>
        <TD colSpan=2><STRONG><FONT face=Arial 
            size=2>(1)&nbsp; If no, explain:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=CertificationNoExplained 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 162px" width=162>
	<PARAM NAME="_ExtentX" VALUE="4286">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CertificationNoExplained">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="CertificationNoExplained">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCertificationNoExplained()
{
	CertificationNoExplained.setDataSource(GetPart1Data);
	CertificationNoExplained.setDataField('CertificationNoExplained');
}
function _CertificationNoExplained_ctor()
{
	CreateLabel('CertificationNoExplained', _initCertificationNoExplained, null);
}
</script>
<% CertificationNoExplained.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></STRONG>
		</TD>
	</tr>
	<tr>
		<td align="left"  wrap ><font face="Arial" size="2"><strong>&nbsp;&nbsp;&nbsp;</strong></font></td>
		<TD align=left colSpan=2><FONT face=Arial size=2><STRONG>(2)&nbsp; If yes, 
            does your company have such certification or 
            authorization?</STRONG></FONT><font face="Arial" size=2 color=maroon><strong>
            <% Response.write ReqCertReadychar %></strong></font>
		</TD>
    <TR>
        <TD align=left></TD>
        <TD align=left colSpan=2>

<tr>
<td align="left" wrap >&nbsp;</td>
        <TD align=left width=35></TD>
        <TD align=left><strong><font face="Arial" size="2">(i)&nbsp;&nbsp;If yes, 
            indicate type and date of certification or authorization(e.g. letter 
            of authorization, license, Certificate of Public Convenience &amp; 
            Necessity (CPCN), tarriff, etc.):
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequiredYesExplanation 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 160px" width=160>
	<PARAM NAME="_ExtentX" VALUE="4233">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequiredYesExplanation">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredYesExplanation">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredYesExplanation()
{
	RequiredYesExplanation.setDataSource(GetPart1Data);
	RequiredYesExplanation.setDataField('RequiredYesExplanation');
}
function _RequiredYesExplanation_ctor()
{
	CreateLabel('RequiredYesExplanation', _initRequiredYesExplanation, null);
}
</script>
<% RequiredYesExplanation.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </font></strong>
            

<tr>
<td align="left" wrap >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        <TD align=left >
</TD>
        <TD align=left><font face="Arial" size="2"><strong>(ii)&nbsp; If no, 
            explain:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequiredNoExplanationel1 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 155px" width=155>
	<PARAM NAME="_ExtentX" VALUE="4101">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequiredNoExplanationel1">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequiredNoExplanation">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequiredNoExplanationel1()
{
	RequiredNoExplanationel1.setDataSource(GetPart1Data);
	RequiredNoExplanationel1.setDataField('RequiredNoExplanation');
}
function _RequiredNoExplanationel1_ctor()
{
	CreateLabel('RequiredNoExplanationel1', _initRequiredNoExplanationel1, null);
}
</script>
<% RequiredNoExplanationel1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong></font>
    <TR>
        <TD align=left colSpan=3>&nbsp; 
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;&nbsp; 
    <TR>
        <TD align=left colSpan=3><strong><font face="Arial" size="3" color="#993300" >1.5&nbsp; Type of Request: 
    
	</font></strong>
    <TR>
        <TD align=left colSpan=3>&nbsp;
    <TR>
        <TD align=left colSpan=3><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write TyReqchar1 %></font></STRONG>
		<font face="Arial" size="2"><strong>&nbsp;1)&nbsp; Code Assignment - Requested NXX:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXXAssign 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 1703px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXAssign">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX1preferred">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXAssign()
{
	NXXAssign.setDataSource(ToRAssRec);
	NXXAssign.setDataField('NXX1preferred');
}
function _NXXAssign_ctor()
{
	CreateLabel('NXXAssign', _initNXXAssign, null);
}
</script>
<% NXXAssign.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong></font>
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <P><FONT face=Arial size=2><STRONG>Secondary NXXs if requested becomes 
            unavailable (optional, you can identify 2 
            NXXs):</STRONG></FONT></FONT></P>
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2A()
{
	NXX2A.setDataSource(ToRAssRec);
	NXX2A.setDataField('NXX2');
}
function _NXX2A_ctor()
{
	CreateLabel('NXX2A', _initNXX2A, null);
}
</script>
<% NXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3A()
{
	NXX3A.setDataSource(ToRAssRec);
	NXX3A.setDataField('NXX3');
}
function _NXX3A_ctor()
{
	CreateLabel('NXX3A', _initNXX3A, null);
}
</script>
<% NXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD align=left colSpan=2>
        <TD align=left><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX1A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1A()
{
	NoNXX1A.setDataSource(ToRAssRec);
	NoNXX1A.setDataField('NoNXX1');
}
function _NoNXX1A_ctor()
{
	CreateLabel('NoNXX1A', _initNoNXX1A, null);
}
</script>
<% NoNXX1A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2A()
{
	NoNXX2A.setDataSource(ToRAssRec);
	NoNXX2A.setDataField('NoNXX2');
}
function _NoNXX2A_ctor()
{
	CreateLabel('NoNXX2A', _initNoNXX2A, null);
}
</script>
<% NoNXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3A()
{
	NoNXX3A.setDataSource(ToRAssRec);
	NoNXX3A.setDataField('NoNXX3');
}
function _NoNXX3A_ctor()
{
	CreateLabel('NoNXX3A', _initNoNXX3A, null);
}
</script>
<% NoNXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX4A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4A()
{
	NoNXX4A.setDataSource(ToRAssRec);
	NoNXX4A.setDataField('NoNXX4');
}
function _NoNXX4A_ctor()
{
	CreateLabel('NoNXX4A', _initNoNXX4A, null);
}
</script>
<% NoNXX4A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX5A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5A()
{
	NoNXX5A.setDataSource(ToRAssRec);
	NoNXX5A.setDataField('NoNXX5');
}
function _NoNXX5A_ctor()
{
	CreateLabel('NoNXX5A', _initNoNXX5A, null);
}
</script>
<% NoNXX5A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> 
        <TR>
        <TD align=left colSpan=2>
        <td nowrap align=left><font face="Arial" color=maroon size="2"><STRONG>
            <% Response.Write Reas4Reqchar %></STRONG></font>            
		</td>    <TR>
        <TD align=left colSpan=3>&nbsp;
    <TR>
        <TD align=left colSpan=3><strong><font face="Arial" color=maroon size="4">&nbsp;
            <% Response.Write TyReqchar2 %></font></strong>&nbsp;<FONT face=Arial 
            size=2><STRONG>2)&nbsp;Update Information 
            (Complete Section 2).&nbsp;&nbsp; NXX requiring 
            update:</STRONG></FONT><font face="Arial" size=2><font face="Arial">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=NXXUpdate style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 93px" 
            width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="NXXUpdate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="NXXUpdate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXUpdate()
{
	NXXUpdate.setDataSource(GetPart1Data);
	NXXUpdate.setDataField('NXXUpdate');
}
function _NXXUpdate_ctor()
{
	CreateLabel('NXXUpdate', _initNXXUpdate, null);
}
</script>
<% NXXUpdate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></font></STRONG>
    <TR>
        <TD align=left colSpan=3>&nbsp;
    <TR>
        <TD align=left colSpan=3><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write TyReqchar3 %></strong></font>&nbsp;
        <FONT face=Arial size=2><STRONG>3)&nbsp; Code Reservation Only - 
            Requested NXX:&nbsp;</STRONG></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXXReserve 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 1926px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXReserve">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX1preferred">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXReserve()
{
	NXXReserve.setDataSource(ToRRsvRec);
	NXXReserve.setDataField('NXX1preferred');
}
function _NXXReserve_ctor()
{
	CreateLabel('NXXReserve', _initNXXReserve, null);
}
</script>
<% NXXReserve.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <P><FONT face=Arial size=2><STRONG>Secondary NXXs if requested becomes 
            unavailable (optional, you can identify 2 
            NXXs):</STRONG></FONT></FONT></P>
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2R()
{
	NXX2R.setDataSource(ToRRsvRec);
	NXX2R.setDataField('NXX2');
}
function _NXX2R_ctor()
{
	CreateLabel('NXX2R', _initNXX2R, null);
}
</script>
<% NXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3R()
{
	NXX3R.setDataSource(ToRRsvRec);
	NXX3R.setDataField('NXX3');
}
function _NXX3R_ctor()
{
	CreateLabel('NXX3R', _initNXX3R, null);
}
</script>
<% NXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD align=left colSpan=2>
        <TD align=left><font face="Arial" size="2"><strong>Undesirable NXXs 
            (optional, you can identify 5 NXXs):</strong></font> 
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX1R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1R()
{
	NoNXX1R.setDataSource(ToRRsvRec);
	NoNXX1R.setDataField('NoNXX1');
}
function _NoNXX1R_ctor()
{
	CreateLabel('NoNXX1R', _initNoNXX1R, null);
}
</script>
<% NoNXX1R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2R()
{
	NoNXX2R.setDataSource(ToRRsvRec);
	NoNXX2R.setDataField('NoNXX2');
}
function _NoNXX2R_ctor()
{
	CreateLabel('NoNXX2R', _initNoNXX2R, null);
}
</script>
<% NoNXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3R()
{
	NoNXX3R.setDataSource(ToRRsvRec);
	NoNXX3R.setDataField('NoNXX3');
}
function _NoNXX3R_ctor()
{
	CreateLabel('NoNXX3R', _initNoNXX3R, null);
}
</script>
<% NoNXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX4R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4R()
{
	NoNXX4R.setDataSource(ToRRsvRec);
	NoNXX4R.setDataField('NoNXX4');
}
function _NoNXX4R_ctor()
{
	CreateLabel('NoNXX4R', _initNoNXX4R, null);
}
</script>
<% NoNXX4R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX5R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5R()
{
	NoNXX5R.setDataSource(ToRRsvRec);
	NoNXX5R.setDataField('NoNXX5');
}
function _NoNXX5R_ctor()
{
	CreateLabel('NoNXX5R', _initNoNXX5R, null);
}
</script>
<% NoNXX5R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD align=left colSpan=2>
		<td align=left><font face="Arial" color=maroon size="2"><STRONG>
            <% Response.Write ReasForReqchar %></STRONG></font>
		</td>
	</TR>
    <TR>
        <TD align=left colSpan=3>
            <P><font face=arial size=2>
            When the Code Applicant desires to change the status of a CO 
            Code from reserved to assigned within the time frame contained 
            within the guidelines, the Code Applicant should complete and submit 
            a new Canadian Central Office Code (NXX) Assignment Request 
            Form.&nbsp;</font></P>
    <TR>
        <TD align=left colSpan=3>&nbsp;
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
    <TR>
        <TD align=left colSpan=3><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
	<strong>1.6 Additional Code Request For 
            Growth:</strong></font> 
    <TR>
        <TD align=left colSpan=3>&nbsp;
    <TR>
        <TD align=left colSpan=2>
<p>&nbsp;</p>
        <TD align=left>
<p><FONT face=Arial size=2>Basis of eligibility for an additional code for growth assigned 
            to the switching entity/POI assumes the following: the initial code 
            or the code previously assigned to a new application meets the 
            exhaust criteria, as specified in the Central Office Code (NXX) 
            Assignment Guidelines, depending on whether the NPA is in a 
            non-jeopardy situation as described in Section 7.3 of the 
            guidelines. The appropriate situation shall be indicated below 
            (select one).</FONT></p>
    <TR>
        <td align="left" colSpan=3><font face="Arial" size="2" color=maroon><strong>&nbsp;
            <% Response.Write JeopardyName2 %>
             &nbsp;</font></STRONG>
        <font face="Arial" size="2"><strong>Non-Jeopardy NPA 
            Situation</strong></font> 
    <TR>
        <TD align=left colSpan=2>
        <TD align=left><FONT face=Arial size=2>I hereby certify that the existing CO Code(s) 
            (NXX) at this Switching Entity/POI is/(are) projected to exhaust 
            within 12 months of the date of this application. This fact is 
            documented on Appendix B and will be supplied to an auditor when 
            requested to do so per Appendix A of the Guidelines.</FONT>
    <TR>
        <td align="left" colSpan=3><font face="Arial" size="2" color=maroon><strong>&nbsp;
            <% Response.Write JeopardyName1 %>
             &nbsp;</font></STRONG>
        <font face="Arial" size="2"><strong>Jeopardy NPA Situation (see Section 
            7.4(c) of the Guidelines) 
        </strong></font>
	<TR>
        <TD align=left colSpan=2><FONT face=Arial></FONT>
        <TD align=left><p><FONT face=Arial size=2>I 
            hereby certify that the existing CO Code(s) (NXX) at this Switching 
            Entity/POI is/(are) projected to exhaust within 6 months of the date 
            of this application. This fact is documented on Appendix B and will 
            be supplied to an auditor when requested to do so per Appendix A of 
            the Guidelines.</FONT></p><FONT face="" size=2></FONT>
    <TR>
        <TD align=left colSpan=3>&nbsp;
<P><P>
            <TABLE background="" border=0 height=280 
            style="HEIGHT: 280px; WIDTH: 969px" width=969>
                
                <TR>
                    <TD align=left colSpan=12><STRONG><FONT color=#993300 
                        face=Arial size=3>APPENDIX B:</FONT></STRONG> 
                <TR>
                    <TD align=left colSpan=12>
                <TR>
                    <TD align=left colSpan=12><FONT face=Arial 
                        size=2><STRONG>NXXs included in growth 
                        calculation:</STRONG></FONT>
                        <FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write NXXGrowthCalValue%></FONT></STRONG>
                <TR>
                    <TD align=left colSpan=12><STRONG><FONT face=Arial 
                        size=2>A.&nbsp; Telephone Numbers (TNs) Available for 
                        Assignment (See Glossary):</FONT></STRONG>
                        <FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write TNsValue%></FONT></STRONG>
                <TR>
                    <TD align=left colSpan=12><FONT face=Arial 
                        size=2>Definitions of 
                        terms may be found in the Glossary section of the 
                        Central Office Code (NXX) Assignment Guidelines.</FONT> 
                <TR>
                    <TD align=left colSpan=6>
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=Month1 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth1()
{
	Month1.setCaption('Month #1');
}
function _Month1_ctor()
{
	CreateLabel('Month1', _initMonth1, null);
}
</script>
<% Month1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=Month2 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth2()
{
	Month2.setCaption('Month #2');
}
function _Month2_ctor()
{
	CreateLabel('Month2', _initMonth2, null);
}
</script>
<% Month2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=Month3 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth3()
{
	Month3.setCaption('Month #3');
}
function _Month3_ctor()
{
	CreateLabel('Month3', _initMonth3, null);
}
</script>
<% Month3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=Month4 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month4">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth4()
{
	Month4.setCaption('Month #4');
}
function _Month4_ctor()
{
	CreateLabel('Month4', _initMonth4, null);
}
</script>
<% Month4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=MOnth5 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="MOnth5">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMOnth5()
{
	MOnth5.setCaption('Month #5');
}
function _MOnth5_ctor()
{
	CreateLabel('MOnth5', _initMOnth5, null);
}
</script>
<% MOnth5.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
                    <TD align=middle>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=20 id=Month6 
                        style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
                        width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Month6">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Month #6">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
                         </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initMonth6()
{
	Month6.setCaption('Month #6');
}
function _Month6_ctor()
{
	CreateLabel('Month6', _initMonth6, null);
}
</script>
<% Month6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD align=left colSpan=6><STRONG><FONT face=Arial 
                        size=2>B.&nbsp; Previous 6-month growth 
                        history:</FONT></STRONG></TD>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month1Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month2Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month3Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month4Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month5Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write Prev6Month6Value%></FONT></STRONG>
</TD></TR>
                <TR>
                    <TD align=left colSpan=12><FONT face=Arial 
                        size=2>Telephone Numbers 
                        (TNs) assigned in each previous month, starting with the 
                        most distant month as Month #1, and Month #6 as the 
                        current month.</FONT></TD></TR>
                <TR>
                    <TD align=left colSpan=6><STRONG><FONT face=Arial 
                        size=2>C.&nbsp; Projected growth - Months&nbsp;&nbsp; 
                        1-6:</FONT></STRONG></TD>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month1Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month2Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month3Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month4Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month5Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth16Month6Value%></FONT></STRONG>
</TD></TR>
                <TR>
                    <TD align=left colSpan=6>&nbsp;&nbsp;&nbsp;&nbsp; 
                        <STRONG><FONT face=Arial size=2>Projected growth - Months&nbsp; 
                        7-12:</FONT></STRONG></TD>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month1Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month2Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month3Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month4Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month5Value%></FONT></STRONG>
                    <TD align=middle><FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write ProjGrowth712Month6Value%></FONT></STRONG>
</TD></TR>
                <TR>
                    <TD align=left colSpan=12><FONT face=Arial size=2>TNs assigned in 
                        each following month, starting with the most recent 
                        month as Month #1.&nbsp; In a jeopardy situation, only 6 
                        months growth porjection is required.</FONT></TD></TR>
                <TR>
                    <TD align=left colSpan=12><STRONG><FONT face=Arial 
                        size=2>D.&nbsp; Average Monthly Growth Rate (From Part C 
                        above):</FONT></STRONG>&nbsp;
                   <FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write AvgMonGrowthRateValue%></FONT></STRONG>
</TD></TR>
                <TR>
                    <TD align=left colSpan=12><STRONG><FONT face=Arial 
                        size=2>E.&nbsp; Months to Exhaust = TNs Available for 
                        Assignment (A) / Average Monthly Growth Rate (D) 
                        =</STRONG></FONT>&nbsp;
                    <FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write MonthsToExhaustValue%></FONT></STRONG>
</TD></TR>
                <TR>
                    <TD align=left colSpan=12><FONT face=Arial size=2>To be assigned an 
                        additional CO Code for growth, &quot;Months to 
                        Exhaust&quot; must be less than or equal to 12 month for 
                        a non -jeopardy NPA (See Section 4.2.1 of the 
                        Guidelines), or less than or equal to 6 months for a 
                        jeopardy NPA (See Section 8.4(c) of the 
                        Guidelines).</FONT></TD></TR>
                <TR>
                    <TD align=left colSpan=12><STRONG><FONT face=Arial 
                        size=2>Explanation:</FONT></STRONG>&nbsp;
                    <FONT face=Arial size=2 color=maroon><strong>
                        <%Response.Write AppendixBExplanationValue%></FONT></STRONG>
		</TD>
	</TR>
</TABLE>
<P><P></P>
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
    <TR>
        <TD align=left colSpan=3><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
		<strong>1.7 Code Request for New 
            Application(see Section 4.2 of the Guidelines)</strong></font>
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
    <TR>
        <TD align=left colSpan=2>
        <TD align=left><font face="Arial" size="2">
	Basis of eligibility for an additional code 
            means that there has not been a code assigned to this switching 
            entity/point of interconnection for this purpose. (Check the 
            applicable space and, if applicable, provide the requested 
            information). If eligibility is based on a category that requires 
            additional explanation or documentation and the code administrator 
            denies a request, the applicant has the option to pursue an appeals 
            process.</font>
    <TR>
        <TD align=left colSpan=3>
			 <dd><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write CodeReqNewchar1 %>
             &nbsp;</font></STRONG><strong><font face="Arial" size="2"> Code is necessary for distinct routing, 
            rating or billing purposes.<font face="Arial" Size="2"><strong> Any additional information that 
            can be provided by the Code Applicant may facilitate the processing 
            of that application.</strong></font></strong></FONT> 
			</dd>
		</TD>
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <STRONG><FONT face=Arial size=2>Description:</FONT></STRONG>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequestNewNecessary 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 147px" width=147>
	<PARAM NAME="_ExtentX" VALUE="3889">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestNewNecessary">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewNecessary">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewNecessary()
{
	RequestNewNecessary.setDataSource(GetPart1Data);
	RequestNewNecessary.setDataField('RequestNewNecessary');
}
function _RequestNewNecessary_ctor()
{
	CreateLabel('RequestNewNecessary', _initRequestNewNecessary, null);
}
</script>
<% RequestNewNecessary.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>            
    <TR>
        <TD align=left colSpan=3>
			<dd><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write CodeReqNewchar2 %>
             &nbsp;</font></STRONG>
		<font face="Arial" size="2"><strong>Other <font size="2">The Code Applicant must provide an explanation of why existing 
            resources assigned to that entity cannot satisfy this 
            requirement.</strong></font></FONT> 
			</dd>
		</TD>
    <TR>
        <TD align=left colSpan=2>
        <TD align=left>
            <FONT face=Arial size=2><strong>Description:</FONT></STRONG>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequestNewOther 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 118px" width=118>
	<PARAM NAME="_ExtentX" VALUE="3122">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestNewOther">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestNewOther">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestNewOther()
{
	RequestNewOther.setDataSource(GetPart1Data);
	RequestNewOther.setDataField('RequestNewOther');
}
function _RequestNewOther_ctor()
{
	CreateLabel('RequestNewOther', _initRequestNewOther, null);
}
</script>
<% RequestNewOther.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>            
    <TR>
        <TD align=left colSpan=3>
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
    <TR>
        <TD align=left colSpan=3><strong><font face="Arial" size="3" color="#993300" style="FONT-WEIGHT: bold">
	1.8 Authorization for entry of Part 2 
            Information into Bellcore databases (Check applicable 
            space):</font></strong>
    <TR>
        <TD align=left colSpan=3>&nbsp;&nbsp;
    <TR>
        <TD align=right valign=top colSpan=2><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write AuthPart2char1 %>
             &nbsp;</font></STRONG>
        <TD align=left><font face="Arial" size="2"><strong>Yes - </strong>I 
            have attached a completed Part 2 of this form. This is the Code 
            Administrator's authorization to input/revise the indicated RDBS 
            and/or BRIDS data. Further, I understand that the Code Administrator 
            may not be the authorized party to input the data. The authorization 
            and/or data input responsibilities are determined on an Operating 
            Company Number level. If the Code Administrator advises me that said 
            Code Administrator does not have Administrative Operating Company 
            Number (AOCN) responsibility for my data inputs, I will contact 
            Bellcore-TRA to determine the correct AOCN company. Upon that 
            determination, I will submit Part 2 directly to the AOCN company for 
            input to RDBS and BRIDS.</font></FONT></STRONG> 
    <TR>
        <TD align=right valign=top colSpan=2><font face="Arial" color=maroon size="4"><strong>&nbsp;
            <% Response.Write AuthPart2char2 %>
             &nbsp;</font></STRONG></TD>
        <TD align=left><font face=arial size="2"><strong>No - </strong>Part 2 
            of this form is not attached. RDBS and BRIDS input will be the 
            responsibility of the Code Applicant. The 66 calendar day 
            nation-wide minimum interval cut-over for RDBS and BRIDS will not 
            begin until input into RDBS and BRIDS has been 
            completed.</font></FONT> 
		</TD>
	<tr>
		<TD align=left colSpan=3>&nbsp;&nbsp;&nbsp;</TD>
	<tr>
		<TD align=left colSpan=3></TD>
	</tr>
	<tr>
		<td align="left" colSpan="3" wrap>
	</td></tr>
</table>


<p>&nbsp;</p><br>		

<h5>&nbsp;</h5></FORM>

<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p>

<hr>
<hr size=4 align=left color=maroon width=71.58% style="HEIGHT: 4px; WIDTH: 680px">
<table align="left" cellPadding="0" cellSpacing="0" width="75%" border=0>
    <TR>
		<TD align=left><strong><font face="Arial" color=maroon size="5">Part 3: Canadian 
            CNA's Response/Confirmation Form</font></strong> 
    
		</TD>
	</TR>
</table>
<P><BR></P>
<hr size=4 align=left color=maroon width=71.68% style="HEIGHT: 4px; WIDTH: 681px">


<p></p>


<p></p>

<p><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>




<br><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
<br></font></FONT>
<table align="center" background ="" border="0" cellPadding="1" cellSpacing="1" style ="WIDTH: 75%" width="75%">
    
    <tr>
        <td align="right" noWrap><strong><font face="Arial" size="4">Applicant Requested 
            Dates</font></strong>
        <td>
        <td align="right" noWrap>
        <td align="left" noWrap>
    
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Date of Requested 
            Application:</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=DateofApp style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" 
            width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofApp">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofApp()
{
	DateofApp.setDataSource(GetPart1Data);
	DateofApp.setDataField('ApplicationDate');
}
function _DateofApp_ctor()
{
	CreateLabel('DateofApp', _initDateofApp, null);
}
</script>
<% DateofApp.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Date of 
            Receipt:</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=DateofReceipt 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofReceipt">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofReceipt()
{
	DateofReceipt.setDataSource(GetPart1Data);
	DateofReceipt.setDataField('DateofReceipt');
}
function _DateofReceipt_ctor()
{
	CreateLabel('DateofReceipt', _initDateofReceipt, null);
}
</script>
<% DateofReceipt.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Date Response Due from CNA 
            Admin:</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=DateofResponse 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 119px" width=119>
	<PARAM NAME="_ExtentX" VALUE="3149">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofResponse">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateResponseDue">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofResponse()
{
	DateofResponse.setDataSource(GetPart1Data);
	DateofResponse.setDataField('DateResponseDue');
}
function _DateofResponse_ctor()
{
	CreateLabel('DateofResponse', _initDateofResponse, null);
}
</script>
<% DateofResponse.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Requested Effective Date of CO 
            Code:</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequestedEffDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 116px" width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestedEffDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="RequestedEffDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedEffDate()
{
	RequestedEffDate.setDataSource(GetPart1Data);
	RequestedEffDate.setDataField('RequestedEffDate');
}
function _RequestedEffDate_ctor()
{
	CreateLabel('RequestedEffDate', _initRequestedEffDate, null);
}
</script>
<% RequestedEffDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right" noWrap></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="right" noWrap>
           <div align="right"><font face ="" size="2"> 
           <font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>The Preferred NPA-NXX Split 
            ID:</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font> </div></td>
        <td align="left" noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NPASplitID style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NPASplitID">
	<PARAM NAME="DataSource" VALUE="GetCOCodeData">
	<PARAM NAME="DataField" VALUE="NPASplitID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPASplitID()
{
	NPASplitID.setDataSource(GetCOCodeData);
	NPASplitID.setDataField('NPASplitID');
}
function _NPASplitID_ctor()
{
	CreateLabel('NPASplitID', _initNPASplitID, null);
}
</script>
<% NPASplitID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="right" noWrap><font face="Arial" size="2"><STRONG>Administrator who is 
            Approving Part 3:</STRONG></font><strong></FONT></strong></FONT>
           <td align="left" noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Label1 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
            width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="CNAUserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setDataSource(GetPart3Data);
	Label1.setDataField('CNAUserName');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>
		<td align="left" noWrap>
            <div align="right">&nbsp;</div>
        <td align="left" noWrap>
    <tr>
        <td align="right" noWrap>
        <td align="right" noWrap>
        <td align="left" noWrap>
        <td align="left" noWrap>
    <tr>
        <td align="right" noWrap>
            <div align="left">&nbsp;</div>
        <td align="right" noWrap>
        <td align="left" noWrap>
        <td align="left" noWrap>
    <tr>
        <td align="right" noWrap>
            <div align="left">&nbsp;</div>
        <td align="right" noWrap>
        <td align="left" noWrap>
        <td align="left" noWrap>
    <tr>
        <td align="right" noWrap>
            <p><font face ="" size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension 
            Date:</STRONG></font></font></font></p>
        <td align="middle" noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ExtensionDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ExtensionDate">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="ExtentionDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initExtensionDate()
{
	ExtensionDate.setDataSource(CodeResvRec);
	ExtensionDate.setDataField('ExtentionDate');
}
function _ExtensionDate_ctor()
{
	CreateLabel('ExtensionDate', _initExtensionDate, null);
}
</script>
<% ExtensionDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        <td align="left" noWrap>
        <td align="left" noWrap>
    <tr>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="middle" noWrap vAlign=top><FONT 
            face=Arial size=1>dd/mm/ccyy </FONT></td>
        <td align="left" noWrap>
            <div align="right">&nbsp; </div>
        <td align="left" noWrap>
</td></tr></table>



<hr>
<form action="xca_Part3int2.asp" method="post" id="formP3" name="formP3">&nbsp; 
<table align=left background=xca_Part3infield.asp#d7c7a4 border=0 cellPadding=1 
cellSpacing=1 style="WIDTH: 100%" width=100%>
 
    <tr>
        <td align=left colSpan=12><strong><FONT face=Arial size=4>SELECT THE 
            APPROPRIATE PART 3 ACTION:</strong> </FONT>
    <TR>
        <TD align=left colSpan=12>&nbsp; 
    <tr>
        <td align=left colSpan=12>&nbsp;</td>
    <TR>
        <TD colSpan=12 noWrap><FONT color=maroon 
            face="" size=4><strong>
            <% Response.Write Part3ResultsChar1 %>
            </FONT></FONT></STRONG></FONT><FONT><FONT face=Arial><strong>Approve 
            Part 1 Request</strong></FONT></FONT> 
    <TR>
        <TD colSpan=12>&nbsp; 
    <TR>
        <TD colSpan=12 noWrap><font face=Arial 
            ><STRONG>-Code 
            Reserved-</STRONG> </font>
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <p align=left><strong><font 
            face=Arial size=2>Requested 
            NPA:</font></strong></p>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ReservedNPA 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 90px" width=90>
	<PARAM NAME="_ExtentX" VALUE="2381">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedNPA">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="ReservedNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedNPA()
{
	ReservedNPA.setDataSource(CodeResvRec);
	ReservedNPA.setDataField('ReservedNPA');
}
function _ReservedNPA_ctor()
{
	CreateLabel('ReservedNPA', _initReservedNPA, null);
}
</script>
<% ReservedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <DIV align=left><font face=Arial size=2 
            ><strong>Reserved NXX: 
            </strong></font></DIV>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=reservedNXX 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 90px" width=90>
	<PARAM NAME="_ExtentX" VALUE="2381">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="reservedNXX">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="ReservedNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initreservedNXX()
{
	reservedNXX.setDataSource(CodeResvRec);
	reservedNXX.setDataField('ReservedNXX');
}
function _reservedNXX_ctor()
{
	CreateLabel('reservedNXX', _initreservedNXX, null);
}
</script>
<% reservedNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->      
    <TR>
        <TD>
        <TD align=left colSpan=2 noWrap><FONT face=Arial 
            size=2><STRONG>Secondary 
            NXXs chosen that are sill available:</STRONG></FONT> 
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Green">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Green"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2R()
{
	NXX2R.setDataSource(ToRRsvRec);
	NXX2R.setDataField('NXX2');
}
function _NXX2R_ctor()
{
	CreateLabel('NXX2R', _initNXX2R, null);
}
</script>
<% NXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Green">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Green"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3R()
{
	NXX3R.setDataSource(ToRRsvRec);
	NXX3R.setDataField('NXX3');
}
function _NXX3R_ctor()
{
	CreateLabel('NXX3R', _initNXX3R, null);
}
</script>
<% NXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD align=left colSpan=2 noWrap><FONT face=Arial 
            size=2><STRONG>Undesirable 
            NXXs:</STRONG></FONT> 
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX1R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1R()
{
	NoNXX1R.setDataSource(ToRRsvRec);
	NoNXX1R.setDataField('NoNXX1');
}
function _NoNXX1R_ctor()
{
	CreateLabel('NoNXX1R', _initNoNXX1R, null);
}
</script>
<% NoNXX1R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX2R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2R()
{
	NoNXX2R.setDataSource(ToRRsvRec);
	NoNXX2R.setDataField('NoNXX2');
}
function _NoNXX2R_ctor()
{
	CreateLabel('NoNXX2R', _initNoNXX2R, null);
}
</script>
<% NoNXX2R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX3R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3R()
{
	NoNXX3R.setDataSource(ToRRsvRec);
	NoNXX3R.setDataField('NoNXX3');
}
function _NoNXX3R_ctor()
{
	CreateLabel('NoNXX3R', _initNoNXX3R, null);
}
</script>
<% NoNXX3R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX4R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4R()
{
	NoNXX4R.setDataSource(ToRRsvRec);
	NoNXX4R.setDataField('NoNXX4');
}
function _NoNXX4R_ctor()
{
	CreateLabel('NoNXX4R', _initNoNXX4R, null);
}
</script>
<% NoNXX4R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX5R style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5R">
	<PARAM NAME="DataSource" VALUE="ToRRsvRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5R()
{
	NoNXX5R.setDataSource(ToRRsvRec);
	NoNXX5R.setDataField('NoNXX5');
}
function _NoNXX5R_ctor()
{
	CreateLabel('NoNXX5R', _initNoNXX5R, null);
}
</script>
<% NoNXX5R.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->      
    <TR>
        <TD>
        <TD align=left colSpan=2 noWrap><font face=Arial 
            size=2><strong>Date of 
            Reservation:</strong></font> 
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ReservedNXXDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 146px" width=146>
	<PARAM NAME="_ExtentX" VALUE="3863">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedNXXDate">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="ReservedNPANXXDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedNXXDate()
{
	ReservedNXXDate.setDataSource(CodeResvRec);
	ReservedNXXDate.setDataField('ReservedNPANXXDate');
}
function _ReservedNXXDate_ctor()
{
	CreateLabel('ReservedNXXDate', _initReservedNXXDate, null);
}
</script>
<% ReservedNXXDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <p align=left><font face=Arial size=2 
            ><STRONG>Your code 
            reservation will be honored until:</STRONG></font></p>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ReservedNXXHonorDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 184px" width=184>
	<PARAM NAME="_ExtentX" VALUE="4868">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedNXXHonorDate">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="ReservedNPANXXHonorDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedNXXHonorDate()
{
	ReservedNXXHonorDate.setDataSource(CodeResvRec);
	ReservedNXXHonorDate.setDataField('ReservedNPANXXHonorDate');
}
function _ReservedNXXHonorDate_ctor()
{
	CreateLabel('ReservedNXXHonorDate', _initReservedNXXHonorDate, null);
}
</script>
<% ReservedNXXHonorDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <p align=left><font face=Arial size=2 
            ><STRONG>Switch 
            Identification (Switching Entity/POI):</font> </STRONG></p>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ReservedSwitchID 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedSwitchID">
	<PARAM NAME="DataSource" VALUE="CodeResvRec">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initReservedSwitchID()
{
	ReservedSwitchID.setDataSource(CodeResvRec);
	ReservedSwitchID.setDataField('SwitchID');
}
function _ReservedSwitchID_ctor()
{
	CreateLabel('ReservedSwitchID', _initReservedSwitchID, null);
}
</script>
<% ReservedSwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD colSpan=12>&nbsp; 
    <TR>
        <TD colSpan=12 noWrap><STRONG><font face=Arial>-Code Update- 
            </font></STRONG>
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <p align=left><strong><font 
            face=Arial size=2>Requested 
            NPA:</font></strong></p>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=UpdatedNPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 84px" 
            width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="UpdatedNPA">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="UpdatedNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUpdatedNPA()
{
	UpdatedNPA.setDataSource(GetPart3Data);
	UpdatedNPA.setDataField('UpdatedNPA');
}
function _UpdatedNPA_ctor()
{
	CreateLabel('UpdatedNPA', _initUpdatedNPA, null);
}
</script>
<% UpdatedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> 
    <TR>
        <TD>
        <TD align=right colSpan=2 noWrap>
            <DIV align=left><FONT face=Arial><STRONG><FONT face="" size=2>Updated NXX:</FONT> </STRONG></FONT></DIV>
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXXUpdate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 84px" 
            width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXUpdate">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="UpdatedNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXUpdate()
{
	NXXUpdate.setDataSource(GetPart3Data);
	NXXUpdate.setDataField('UpdatedNXX');
}
function _NXXUpdate_ctor()
{
	CreateLabel('NXXUpdate', _initNXXUpdate, null);
}
</script>
<% NXXUpdate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD colSpan=12>&nbsp; 
    <TR>
        <TD colSpan=12><FONT face=Arial><STRONG>-Code Assigned-</STRONG></FONT> 
    <tr>
        <td>
        <td align=right colSpan=2 noWrap>
            <p align=left><strong><font 
            face=Arial size=2>Requested NPA:</font></strong> 
            </p>
        <td align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AssignedNPA 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AssignedNPA">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="AssignedNPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNPA()
{
	AssignedNPA.setDataSource(CodeAssignRec);
	AssignedNPA.setDataField('AssignedNPA');
}
function _AssignedNPA_ctor()
{
	CreateLabel('AssignedNPA', _initAssignedNPA, null);
}
</script>
<% AssignedNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</td>
    <tr>
        <td></td>
        <td align=left colSpan=2 noWrap>
            <p><font face=Arial size=2><strong>Assigned NXX:</strong></font> 
            </p>
        <td align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AssignedNXX 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AssignedNXX">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="AssignedNXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNXX()
{
	AssignedNXX.setDataSource(CodeAssignRec);
	AssignedNXX.setDataField('AssignedNXX');
}
function _AssignedNXX_ctor()
{
	CreateLabel('AssignedNXX', _initAssignedNXX, null);
}
</script>
<% AssignedNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
         </td>
    <TR>
        <TD>
        <TD align=left colSpan=2 noWrap><FONT face=Arial 
            size=2><STRONG>Secondary 
            NXXs chosen that are sill available:</STRONG></FONT> 
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Green">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Green"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX2A()
{
	NXX2A.setDataSource(ToRAssRec);
	NXX2A.setDataField('NXX2');
}
function _NXX2A_ctor()
{
	CreateLabel('NXX2A', _initNXX2A, null);
}
</script>
<% NXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Green">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Green"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXX3A()
{
	NXX3A.setDataSource(ToRAssRec);
	NXX3A.setDataField('NXX3');
}
function _NXX3A_ctor()
{
	CreateLabel('NXX3A', _initNXX3A, null);
}
</script>
<% NXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD align=left colSpan=2 noWrap><FONT face=Arial 
            size=2><STRONG>Undesirable 
            NXXs:</STRONG></FONT> 
        <TD align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX1A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX1A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX1">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX1A()
{
	NoNXX1A.setDataSource(ToRAssRec);
	NoNXX1A.setDataField('NoNXX1');
}
function _NoNXX1A_ctor()
{
	CreateLabel('NoNXX1A', _initNoNXX1A, null);
}
</script>
<% NoNXX1A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX2A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX2A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX2">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX2A()
{
	NoNXX2A.setDataSource(ToRAssRec);
	NoNXX2A.setDataField('NoNXX2');
}
function _NoNXX2A_ctor()
{
	CreateLabel('NoNXX2A', _initNoNXX2A, null);
}
</script>
<% NoNXX2A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX3A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX3A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX3">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX3A()
{
	NoNXX3A.setDataSource(ToRAssRec);
	NoNXX3A.setDataField('NoNXX3');
}
function _NoNXX3A_ctor()
{
	CreateLabel('NoNXX3A', _initNoNXX3A, null);
}
</script>
<% NoNXX3A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX4A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX4A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX4">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX4A()
{
	NoNXX4A.setDataSource(ToRAssRec);
	NoNXX4A.setDataField('NoNXX4');
}
function _NoNXX4A_ctor()
{
	CreateLabel('NoNXX4A', _initNoNXX4A, null);
}
</script>
<% NoNXX4A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NoNXX5A style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 55px" 
            width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NoNXX5A">
	<PARAM NAME="DataSource" VALUE="ToRAssRec">
	<PARAM NAME="DataField" VALUE="NoNXX5">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNoNXX5A()
{
	NoNXX5A.setDataSource(ToRAssRec);
	NoNXX5A.setDataField('NoNXX5');
}
function _NoNXX5A_ctor()
{
	CreateLabel('NoNXX5A', _initNoNXX5A, null);
}
</script>
<% NoNXX5A.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->

	<tr>
		<td>
        <td align=left colSpan=2 noWrap>
            <p><font face=Arial size=2><strong>NXX Effective 
            Date:</strong></font></p>
        <td align=left colSpan=9><font face=arial size=1 
            >
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AssignedNPANXXDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 144px" width=144>
	<PARAM NAME="_ExtentX" VALUE="3810">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AssignedNPANXXDate">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="AssignedNPANXXDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNPANXXDate()
{
	AssignedNPANXXDate.setDataSource(CodeAssignRec);
	AssignedNPANXXDate.setDataField('AssignedNPANXXDate');
}
function _AssignedNPANXXDate_ctor()
{
	CreateLabel('AssignedNPANXXDate', _initAssignedNPANXXDate, null);
}
</script>
<% AssignedNPANXXDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            <font face=arial size=1>dd/mm/ccyy 
            </font></font></td>
    <tr>
        <td><strong></strong>
        <td align=left colSpan=2 noWrap><font face=Arial 
            size=2><STRONG>Switch 
            Identification (Switching Entity/POI):</STRONG> 
 </font></td>
        <td align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=SwitchID style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
            width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="SwitchID">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSwitchID()
{
	SwitchID.setDataSource(CodeAssignRec);
	SwitchID.setDataField('SwitchID');
}
function _SwitchID_ctor()
{
	CreateLabel('SwitchID', _initSwitchID, null);
}
</script>
<% SwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
	<tr>
        <td></td>
        <td align=left colSpan=2><font face=Arial size=2 
            >&nbsp;&nbsp;&nbsp; <STRONG>Rate Center: </STRONG></font>
        <td align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RateCenter style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 75px" 
            width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RateCenter">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="RateCenter">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRateCenter()
{
	RateCenter.setDataSource(CodeAssignRec);
	RateCenter.setDataField('RateCenter');
}
function _RateCenter_ctor()
{
	CreateLabel('RateCenter', _initRateCenter, null);
}
</script>
<% RateCenter.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->	
</td>
    <tr>
        <td></td>
        <td align=left colSpan=2><font face=Arial size=2 
            ><STRONG>Routing and Rating 
            information is complete:</STRONG></font></td>
        <td align=left colSpan=9><font color=maroon 
            face=Arial size=2><strong>
            <% Response.Write RRCompletechar %>
            </strong></font></td>
    <tr>
        <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
        <td align=left colSpan=2><FONT face=Arial size=2 
            ><STRONG>Additional RDBS and 
            BRIDS information is required as follows:</STRONG></FONT> 
        <td align=left colSpan=9>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RRDescription 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RRDescription">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="RRDescription">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRRDescription()
{
	RRDescription.setDataSource(CodeAssignRec);
	RRDescription.setDataField('RRDescription');
}
function _RRDescription_ctor()
{
	CreateLabel('RRDescription', _initRRDescription, null);
}
</script>
<% RRDescription.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
	<tr>
        <td></td>
        <td colSpan=11><font face=Arial size=2 
            ><STRONG>The Code 
            Administrator</STRONG></font>&nbsp;<font color=maroon 
            face=Arial size=2><strong>
            <% Response.Write CNAResponsiblechar1 %>
             &nbsp;</strong></font><font face=Arial size=2><STRONG>responsible for inputting Part 2 
            Information into RDBS and BRIDS.</font> </STRONG></td>
    <tr>
        <td></td>
        <td colSpan=11><font face=Arial size=2 
            ><STRONG>To be published in 
            the LERG and TPM by:</STRONG><strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=LERGDate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 68px" 
            width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="LERGDate">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="LERGDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLERGDate()
{
	LERGDate.setDataSource(CodeAssignRec);
	LERGDate.setDataField('LERGDate');
}
function _LERGDate_ctor()
{
	CreateLabel('LERGDate', _initLERGDate, null);
}
</script>
<% LERGDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong><font face=arial size=1>(dd/mm/ccyy),&nbsp;<font face=Arial size=2> 
            <STRONG>additional RDBS and BRIDS information needs to be received by 
            the </STRONG><STRONG>Code Administrator no later 
            than:</STRONG>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=LERGDate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 93px" 
            width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="LERGDate">
	<PARAM NAME="DataSource" VALUE="CodeAssignRec">
	<PARAM NAME="DataField" VALUE="RRReturnDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLERGDate()
{
	LERGDate.setDataSource(CodeAssignRec);
	LERGDate.setDataField('RRReturnDate');
}
function _LERGDate_ctor()
{
	CreateLabel('LERGDate', _initLERGDate, null);
}
</script>
<% LERGDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
<font face=arial size=1>(dd/mm/ccyy).<font face=Arial size=2></font></font></font></font></font> </td></tr>
    <tr>
        <td align=left colSpan=12>&nbsp;</td></tr>
    <tr>
        <td align=left colSpan=12>&nbsp;</td></tr>
    <tr>
        <td colSpan=12><font color=maroon face=Arial 
            size=4><STRONG>
            <% Response.Write Part3ResultsChar4 %>
            </font><FONT face=Arial>Part 1 </FONT><font 
            face=Arial>Form</STRONG></STRONG><strong 
            > Incomplete.</strong></font> </td></tr>
    <tr>
        <td></td>
        <td align=left colSpan=11>
            <p><font face=Arial size=2>Additional information required in the following 
            section(s):</font></p></td></td>
    <tr>
        <td></td>
        <td colSpan=11>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part3IncompleteDescription 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 179px" width=179>
	<PARAM NAME="_ExtentX" VALUE="4736">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part3IncompleteDescription">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="Part3IncompleteDescription">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart3IncompleteDescription()
{
	Part3IncompleteDescription.setDataSource(GetPart3Data);
	Part3IncompleteDescription.setDataField('Part3IncompleteDescription');
}
function _Part3IncompleteDescription_ctor()
{
	CreateLabel('Part3IncompleteDescription', _initPart3IncompleteDescription, null);
}
</script>
<% Part3IncompleteDescription.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td colSpan=12>&nbsp; 
    <tr>
        <td colSpan=12><font color=maroon face=Arial 
            size=4><strong>
            <% Response.Write Part3ResultsChar5 %>
            </strong></font><FONT face=Arial><strong 
            >Part 1 Form completed, Code request denied. 
            </strong></FONT></td></tr>
<tr>
        <td>
        <td align=left colSpan=11><font face=Arial 
            size=2>Explanation is: </font>
	<tr>
        <td></td>
        <td align=left colSpan=11>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part3DenialDescription 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 149px" width=149>
	<PARAM NAME="_ExtentX" VALUE="3942">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part3DenialDescription">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="Part3DenialDescription">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart3DenialDescription()
{
	Part3DenialDescription.setDataSource(GetPart3Data);
	Part3DenialDescription.setDataField('Part3DenialDescription');
}
function _Part3DenialDescription_ctor()
{
	CreateLabel('Part3DenialDescription', _initPart3DenialDescription, null);
}
</script>
<% Part3DenialDescription.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->           
</td>
    <tr>
        <td align=right colSpan=12>
            <DIV align=left>&nbsp;&nbsp;</DIV></td></tr>
    <tr>
        <td colSpan=12><font color=maroon face=Arial 
            size=4><strong>
            <% Response.Write Part3ResultsChar6 %>
            </strong></font><FONT face=Arial><strong 
            >Part 1 Assignment Activity Suspended by the 
            Administrator.</strong></FONT> </td></tr>
    <tr>
        <td>
        <td align=left colSpan=11><font face=Arial 
            size=2>Explanation is:</font> 
    <tr>
        <td>
        <td align=left colSpan=11>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part3SuspendedDescription 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 178px" width=178>
	<PARAM NAME="_ExtentX" VALUE="4710">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part3SuspendedDescription">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="part3SuspendedDescription">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart3SuspendedDescription()
{
	Part3SuspendedDescription.setDataSource(GetPart3Data);
	Part3SuspendedDescription.setDataField('part3SuspendedDescription');
}
function _Part3SuspendedDescription_ctor()
{
	CreateLabel('Part3SuspendedDescription', _initPart3SuspendedDescription, null);
}
</script>
<% Part3SuspendedDescription.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</td>
	<tr>
        <td>
        <td align=left colSpan=11><font face=arial 
            size=3><strong>Further 
            Action:</strong></font> </td>
    <tr>
        <td>
        <td align=left colSpan=11>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part3SuspendedFurtherAction 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 194px" width=194>
	<PARAM NAME="_ExtentX" VALUE="5133">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part3SuspendedFurtherAction">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="Part3SuspendedFurtherAction">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart3SuspendedFurtherAction()
{
	Part3SuspendedFurtherAction.setDataSource(GetPart3Data);
	Part3SuspendedFurtherAction.setDataField('Part3SuspendedFurtherAction');
}
function _Part3SuspendedFurtherAction_ctor()
{
	CreateLabel('Part3SuspendedFurtherAction', _initPart3SuspendedFurtherAction, null);
}
</script>
<% Part3SuspendedFurtherAction.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
    <tr>
        <td align=left colSpan=12><font face=Arial 
            size=2><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td></tr>
    <tr>
        <td align=left colSpan=12><font face=Arial 
            size=2><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td></tr>
    <tr>
        <td align=left colSpan=12><font face=Arial 
            size=2><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; NPA Jeopardy = <font 
            color=maroon face=Arial size=3>
            <%Response.write P3Jeopardychar%>
            </font></strong></font>
    <tr>
        <td></td>
        <td align=left colSpan=12><font face=Arial 
            size=1>If YES, refer to Section 7 of the 
            assignment guidelines</font> </td></tr>
    <tr>
        <td align=left colSpan=12><font face=Arial 
            size=2><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td></tr>
    <tr>
        <td align=left colSpan=12><font face=Arial 
            size=2><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font></td></tr>
    <tr>
        <td colSpan=12>&nbsp;&nbsp; <font face=Arial 
            size=3><strong>Remarks:
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Remarks style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 59px" 
            width=59>
	<PARAM NAME="_ExtentX" VALUE="1561">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Remarks">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="Remarks">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRemarks()
{
	Remarks.setDataSource(GetPart3Data);
	Remarks.setDataField('Remarks');
}
function _Remarks_ctor()
{
	CreateLabel('Remarks', _initRemarks, null);
}
</script>
<% Remarks.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</strong></font></td></tr>
</table> 

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<hr>

<hr size=4 align=left color=maroon width=65%>
<table align="left" cellPadding="0" cellSpacing="0" width="75%" border=0>
    <TR>
		<TD align=left><strong><font face="Arial" color=maroon size="5"> 
Part 4: 
            Confirmation of Code Activation</font></strong> 
</TD>
	</TR>
</table>

<p><BR></p>
<hr size=4 align=left color=maroon width=65%>

<p><br></p>

<table align="left" border="0" cellPadding="0" cellSpacing="0">
	<tr>
		<td><strong><font size="2" face=arial>Code Applicants are required to retain a 
            copy of all application forms, appendices and supporting data in the 
            event of an audit.</strong></FONT> 
		</td>
	</tr>
    <TR>
        <TD style="FONT-WEIGHT: bold">&nbsp;
	<tr>
        <td wrap style="FONT-WEIGHT: bold"><label><strong><font size="3" face=Arial color="#993300">
        Contact 
            Information:</font></strong></label> 
		</td>
	</tr>
</table>

<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>

<table align="center" border="0" cellPadding="1" cellSpacing="1" >
    <tbody>
    
    <tr>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code Applicant 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" wrap><font face="Arial"> </font>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Entity 
            Name</STRONG></STRONG></font></font></font><STRONG><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
 </font></font></STRONG> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4EntityName%></font></Strong>
		</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Entity Name</STRONG> 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityName 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 2905px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityName">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityName()
{
	AdminEntityName.setDataSource(GetAdminEntityName);
	AdminEntityName.setDataField('EntityName');
}
function _AdminEntityName_ctor()
{
	CreateLabel('AdminEntityName', _initAdminEntityName, null);
}
</script>
<% AdminEntityName.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Contact 
            Name</STRONG></font></font></font><STRONG><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font></STRONG> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4UserName%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Contact Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityContact 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 2925px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityContact()
{
	AdminEntityContact.setDataSource(GetAdminEntityName);
	AdminEntityContact.setDataField('EntityContact');
}
function _AdminEntityContact_ctor()
{
	CreateLabel('AdminEntityContact', _initAdminEntityContact, null);
}
</script>
<% AdminEntityContact.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Street 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4EntityAddress%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Street 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityAddress 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 2945px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityAddress()
{
	AdminEntityAddress.setDataSource(GetAdminEntityName);
	AdminEntityAddress.setDataField('EntityAddress');
}
function _AdminEntityAddress_ctor()
{
	CreateLabel('AdminEntityAddress', _initAdminEntityAddress, null);
}
</script>
<% AdminEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>City</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4EntityCity%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>City</STRONG> 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityCity 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 2965px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityCity()
{
	AdminEntityCity.setDataSource(GetAdminEntityName);
	AdminEntityCity.setDataField('EntityCity');
}
function _AdminEntityCity_ctor()
{
	CreateLabel('AdminEntityCity', _initAdminEntityCity, null);
}
</script>
<% AdminEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Province</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4EntityProvince%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Province</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityProvince 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 2985px; WIDTH: 95px" width=95>
	<PARAM NAME="_ExtentX" VALUE="2514">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityProvince()
{
	AdminEntityProvince.setDataSource(GetAdminEntityName);
	AdminEntityProvince.setDataField('EntityProvince');
}
function _AdminEntityProvince_ctor()
{
	CreateLabel('AdminEntityProvince', _initAdminEntityProvince, null);
}
</script>
<% AdminEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Postal Code</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4EntityPostalCode%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Postal Code</STRONG> 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityPostalCode 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 3005px; WIDTH: 111px" width=111>
	<PARAM NAME="_ExtentX" VALUE="2937">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityPostalCode()
{
	AdminEntityPostalCode.setDataSource(GetAdminEntityName);
	AdminEntityPostalCode.setDataField('EntityPostalCode');
}
function _AdminEntityPostalCode_ctor()
{
	CreateLabel('AdminEntityPostalCode', _initAdminEntityPostalCode, null);
}
</script>
<% AdminEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>E-Mail Address</STRONG> 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4UserEmail%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>E-Mail 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityEmail 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 3025px; WIDTH: 75px" width=75>
	<PARAM NAME="_ExtentX" VALUE="1984">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityEmail()
{
	AdminEntityEmail.setDataSource(GetAdminEntityName);
	AdminEntityEmail.setDataField('EntityEmail');
}
function _AdminEntityEmail_ctor()
{
	CreateLabel('AdminEntityEmail', _initAdminEntityEmail, null);
}
</script>
<% AdminEntityEmail.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>    
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2"> 
        <font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Facsimile</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4UserFax%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Facsimile</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityFax 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 3045px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityFax()
{
	AdminEntityFax.setDataSource(GetAdminEntityName);
	AdminEntityFax.setDataField('EntityFax');
}
function _AdminEntityFax_ctor()
{
	CreateLabel('AdminEntityFax', _initAdminEntityFax, null);
}
</script>
<% AdminEntityFax.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</font>
		</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Telephone</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4UserTelephone%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Telephone</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityTelephone 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 3065px; WIDTH: 107px" width=107>
	<PARAM NAME="_ExtentX" VALUE="2831">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityTelephone()
{
	AdminEntityTelephone.setDataSource(GetAdminEntityName);
	AdminEntityTelephone.setDataField('EntityTelephone');
}
function _AdminEntityTelephone_ctor()
{
	CreateLabel('AdminEntityTelephone', _initAdminEntityTelephone, null);
}
</script>
<% AdminEntityTelephone.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension</STRONG></STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" size=2 color=maroon><Strong>
        <%Response.Write P4UserExtension%></font></Strong>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension</STRONG></STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AdminEntityExtension 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 3085px; WIDTH: 101px" width=101>
	<PARAM NAME="_ExtentX" VALUE="2672">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityExtension()
{
	AdminEntityExtension.setDataSource(GetAdminEntityName);
	AdminEntityExtension.setDataField('EntityExtension');
}
function _AdminEntityExtension_ctor()
{
	CreateLabel('AdminEntityExtension', _initAdminEntityExtension, null);
}
</script>
<% AdminEntityExtension.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr></tbody>

</table>
<P>

<BR><BR></P>
<P>
<BR></P>


<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 align=left height=216 style="HEIGHT: 216px; WIDTH: 480px" width=480>
    
    <TR>
        <TD colSpan=3><font size="3" face=arial color="#993300"><strong>
		1.1 CO Code 
            Information:</strong></font>
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;
	<TR>
		<TD><font size="2" face=arial><STRONG>1)&nbsp;&nbsp;</STRONG></font></TD>
		<TD><STRONG>CO 
            Code:</STRONG></TD>
		<TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part4NPA style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 64px" 
            width=64>
	<PARAM NAME="_ExtentX" VALUE="1693">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part4NPA">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="Part4NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart4NPA()
{
	Part4NPA.setDataSource(GetPart4Data);
	Part4NPA.setDataField('Part4NPA');
}
function _Part4NPA_ctor()
{
	CreateLabel('Part4NPA', _initPart4NPA, null);
}
</script>
<% Part4NPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
-
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part4NXX style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 64px" 
            width=64>
	<PARAM NAME="_ExtentX" VALUE="1693">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part4NXX">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="Part4NXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart4NXX()
{
	Part4NXX.setDataSource(GetPart4Data);
	Part4NXX.setDataField('Part4NXX');
}
function _Part4NXX_ctor()
{
	CreateLabel('Part4NXX', _initPart4NXX, null);
}
</script>
<% Part4NXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
            
</TD>
	</TR>
	<TR>
		<TD valign=top><FONT face=Arial size=2><STRONG>2)</STRONG></FONT></TD>
		<TD><STRONG>Switch 
            Identification <BR>(Switching 
            Entity/POI):</STRONG></FONT></TD>
		<TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=p4SwitchID style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
            width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="p4SwitchID">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="SwitchID">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initp4SwitchID()
{
	p4SwitchID.setDataSource(GetPart4Data);
	p4SwitchID.setDataField('SwitchID');
}
function _p4SwitchID_ctor()
{
	CreateLabel('p4SwitchID', _initp4SwitchID, null);
}
</script>
<% p4SwitchID.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
		</TD>
	</TR>
	<TR>
		<TD><font size="2" face=arial><STRONG>3)</STRONG></font></TD>
		<TD><STRONG>Part 1 
            Application Date:</STRONG></FONT></TD>
		<TD align=left><font size=3>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=p1ApplicationDate 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3126px; WIDTH: 105px" 
            width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="p1ApplicationDate">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initp1ApplicationDate()
{
	p1ApplicationDate.setDataSource(GetPart1Data);
	p1ApplicationDate.setDataField('ApplicationDate');
}
function _p1ApplicationDate_ctor()
{
	CreateLabel('p1ApplicationDate', _initp1ApplicationDate, null);
}
</script>
<% p1ApplicationDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
	<TR>
		<TD></TD>
		<TD><font size="2" face=arial><STRONG>Part 1 Date of 
            Receipt:</STRONG></font></TD>
		<TD align=left><font size=3>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=DateofP1Application 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3126px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="DateofP1Application">
	<PARAM NAME="DataSource" VALUE="GetPart1Data">
	<PARAM NAME="DataField" VALUE="DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initDateofP1Application()
{
	DateofP1Application.setDataSource(GetPart1Data);
	DateofP1Application.setDataField('DateofReceipt');
}
function _DateofP1Application_ctor()
{
	CreateLabel('DateofP1Application', _initDateofP1Application, null);
}
</script>
<% DateofP1Application.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
		</TD>
	</TR>
    <TR>
        <TD>
        <TD><font size="2" face=arial><STRONG>Part 3 Date of 
            Receipt:</STRONG></font>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=P3DateofResponse 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3146px; WIDTH: 109px" 
            width=109>
	<PARAM NAME="_ExtentX" VALUE="2884">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="P3DateofResponse">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="P3DateofReceipt">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initP3DateofResponse()
{
	P3DateofResponse.setDataSource(GetPart3Data);
	P3DateofResponse.setDataField('P3DateofReceipt');
}
function _P3DateofResponse_ctor()
{
	CreateLabel('P3DateofResponse', _initP3DateofResponse, null);
}
</script>
<% P3DateofResponse.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD>
        <TD><font size="2" face=arial><STRONG>Part 4 Date of 
            Receipt:</STRONG></font>
        <TD align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=P4DateOfReceipt 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3166px; WIDTH: 105px" 
            width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="P4DateOfReceipt">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initP4DateOfReceipt()
{
	P4DateOfReceipt.setDataSource(GetPart4Data);
	P4DateOfReceipt.setDataField('ApplicationDate');
}
function _P4DateOfReceipt_ctor()
{
	CreateLabel('P4DateOfReceipt', _initP4DateOfReceipt, null);
}
</script>
<% P4DateOfReceipt.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
    <TR>
        <TD>
        <TD><font size="2" face=arial><STRONG>NXX Effective Date:</STRONG></font>
        <TD align=left><font size=3>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=P4EffDate 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3186px; WIDTH: 49px" width=49>
	<PARAM NAME="_ExtentX" VALUE="1296">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="P4EffDate">
	<PARAM NAME="DataSource" VALUE="GetPart3Data">
	<PARAM NAME="DataField" VALUE="EffDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initP4EffDate()
{
	P4EffDate.setDataSource(GetPart3Data);
	P4EffDate.setDataField('EffDate');
}
function _P4EffDate_ctor()
{
	CreateLabel('P4EffDate', _initP4EffDate, null);
}
</script>
<% P4EffDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</font>
		</TD>
    <TR>
        <TD>&nbsp;
        <TD>&nbsp;&nbsp;&nbsp;&nbsp;
        <TD align=left></TD></TR>
</TABLE>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<TABLE ALIGN=left BORDER=0 CELLSPACING=0 CELLPADDING=0 background="" height=316 style="HEIGHT: 316px; WIDTH: 709px" width=709>
    
    <TR>
        <TD colSpan=2><font size="3" face=arial color="#993300"><STRONG>1.2 
            Confirmation Information:</STRONG></font>
    <TR>
        <TD colSpan=2>&nbsp; 
    <TR>
        <TD colSpan=2>
<P><FONT face=arial size=2><strong>By submiting a Part 4, I certify 
            that the CO Code(NXX) specified below is in service and that the CO 
            Code (NXX) is being used for purpose specified in the original 
            application (See Section 6.3.3). </FONT></P></STRONG>
    <TR>
        <TD colSpan=2>&nbsp;
    
    <TR>
        <TD><FONT face=arial size=2><STRONG>In-Service Date:
            <%Response.write "(Date must be " & GetP4DAYS() &  " days after the Part1 date of receipt)" %>
 
        </STRONG></FONT>
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=InServiceDate 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 3206px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="InServiceDate">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="InServiceDate">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initInServiceDate()
{
	InServiceDate.setDataSource(GetPart4Data);
	InServiceDate.setDataField('InServiceDate');
}
function _InServiceDate_ctor()
{
	CreateLabel('InServiceDate', _initInServiceDate, null);
}
</script>
<% InServiceDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
    <TR>
        <TD> 
        <TD>
	<TR>
		<TD><FONT face="Arial" size=2><STRONG>Authorized Representative of Code 
            Applicant:&nbsp;&nbsp; </STRONG></FONT>

</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=p4AuthorizedRep 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 74px" width=74>
	<PARAM NAME="_ExtentX" VALUE="1958">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="p4AuthorizedRep">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="Part4Name">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initp4AuthorizedRep()
{
	p4AuthorizedRep.setDataSource(GetPart4Data);
	p4AuthorizedRep.setDataField('Part4Name');
}
function _p4AuthorizedRep_ctor()
{
	CreateLabel('p4AuthorizedRep', _initp4AuthorizedRep, null);
}
</script>
<% p4AuthorizedRep.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</TD>
	</TR>
	<TR>
		<TD><FONT face="Arial" size=2><STRONG>Title</STRONG>:</FONT> 
</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=p4AuthorizedRepTitle 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="p4AuthorizedRepTitle">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="Part4Title">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initp4AuthorizedRepTitle()
{
	p4AuthorizedRepTitle.setDataSource(GetPart4Data);
	p4AuthorizedRepTitle.setDataField('Part4Title');
}
function _p4AuthorizedRepTitle_ctor()
{
	CreateLabel('p4AuthorizedRepTitle', _initp4AuthorizedRepTitle, null);
}
</script>
<% p4AuthorizedRepTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->            
</TD></TR>
    <TR>
        <TD>&nbsp; 
        <TD>
	<TR>
		<TD><FONT face="Arial" size=2><STRONG>Logon that created Part 
            4:</STRONG></FONT> 
		</TD>
		<TD><FONT face="Arial" size=2 color=maroon><STRONG>
		<%Response.Write Signature%></STRONG></FONT>
		</TD>
	</TR>
	<TR>
		<TD><FONT face="Arial" size=2><STRONG>Date:</STRONG></FONT> 
		</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=Part4Date style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 66px" 
            width=66>
	<PARAM NAME="_ExtentX" VALUE="1746">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part4Date">
	<PARAM NAME="DataSource" VALUE="GetPart4Data">
	<PARAM NAME="DataField" VALUE="Part4Date">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart4Date()
{
	Part4Date.setDataSource(GetPart4Data);
	Part4Date.setDataField('Part4Date');
}
function _Part4Date_ctor()
{
	CreateLabel('Part4Date', _initPart4Date, null);
}
</script>
<% Part4Date.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
	</TR>
    <TR>
        <TD>&nbsp; 
        <TD>
    <TR>
        <TD>&nbsp; 
        <TD>
	<TR>
		<TD><a HREF="#top"> Back to 
            Top</a> 
		</TD>
		<TD align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnReturnToMenu 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 3306px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMenu">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
    </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturnToMenu()
{
	btnReturnToMenu.value = 'Return';
	btnReturnToMenu.setStyle(0);
}
function _btnReturnToMenu_ctor()
{
	CreateButton('btnReturnToMenu', _initbtnReturnToMenu, null);
}
</script>
<% btnReturnToMenu.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>&nbsp;
<P>

<p></p>
</form> 
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

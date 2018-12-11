<%@ Language=VBScript %>

<%If session("UserEntityType") <> "u" and session("UserEntityType") <> "a" then
		Response.Redirect("..\default.htm")
end if%>

<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
</form>
<!--#include file="xca_CNASLib.inc"-->
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">



<title>Part 3: Canadian CNA's Response/Confirmation Form</title>
</head>
<%
Sub btnReturnToMenu_onclick()
Response.Redirect "xca_MenuSubPost.asp"
End Sub
'check to see where data coming from
ViewP3=session("ViewP3")

Tix=int(Request.Form("P3ViewTix"))
UserEntityID=int(session("UserEntityID"))
session("P3ViewTix")=Tix
session("P1EntityID")=AppEntityID
'AdminUserName=session("UserUserName")

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
			'Check for invalid tix  Mike put this in
	if checkTix="" then	
		session("NoTixSent")="DidNotSend"
		Response.Redirect session("Here")
	end if 

	
	'Check for invalid tix  Mike put this in
	if checkTix="" then	
		session("NoTixSent")="DidNotSend"
		Response.Redirect session("Here")
	end if 
	
''sql 1 gets the entity recordset of the ADMIN
sql1="Select EntityID from xca_Entity where EntityName= '"&AdminEntityID&"'"
	
''sql2 gets the p1 recordset of the P1 request using the Tix from the input form.
sql2="Select * from xca_Part1 where Tix= '"&Tix&"'"
	GetPart1Data.setSQLText(sql2)
	GetPart1Data.Open	
	P1UserID = GetPart1Data.fields.getValue("UserID")

session("NoTixSent")=""
''sql gets user entity recordset of user
sql = "SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = '"&Tix&"' AND xca_User.UserID = '"&P1UserID&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
'sql = "Select * from xca_Entity where EntityID = '"&UserEntityID&"'"
	GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open


AdminData=session("ADMIN")

'get Admin info for top of form
sqlADMIN="Select * from xca_Entity where EntityName ='"&AdminData&"'"
	GetAdminEntityName.setSQLText(sqlADMIN)
	GetAdminEntityName.Open


''sql3 gets the NPASplitID of the NPA-preferred NXX that was requested frpom the COCode table
sql33="SELECT xca_COCode.*, xca_Part1.NXX1preferred FROM xca_COCode INNER JOIN xca_Part1 ON xca_COCode.NPA = xca_Part1.NPA WHERE xca_Part1.Tix = '"&Tix&"' AND xca_COCode.NPASplitID <> 'Excluded'"
	GetCOCodeData.setSQLText(sql33)
	GetCOCodeData.Open
	
''sql3 gets the p3 recordset of the P3 request using the Tix from the input form.
sql3="Select * from xca_Part3 where Tix= '"&Tix&"'"
	GetPart3Data.setSQLText(sql3)
	GetPart3Data.Open

Part3ResultsValue=GetPart3Data.fields.getValue("Part3Result")
Select Case Part3ResultsValue
Case "a"
	Part3ResultsChar1="**"
Case "r"
	Part3ResultsChar1="**"
Case "u"
	Part3ResultsChar1="**"
Case "i"
	Part3ResultsChar4="**"
Case "d"
	Part3ResultsChar5="**"
Case "s"
	Part3ResultsChar6="**"
End Select


RequestStatusValue=GetPart1Data.fields.getValue("RequestStatus")
Select Case RequestStatusValue
Case "NW"
	RequestStatuschar="Pending - New Code"
Case "UP"
	RequestStatuschar="Pending - Update"
Case "AS"
	RequestStatuschar="Pending - Assigned"
Case "RS"
	RequestStatuschar="Pending - Reserved"
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

P3JeopardyValue = GetPart1Data.fields.getValue("NPAinJeopardy")

select case P3JeopardyValue
case "y"
	P3Jeopardychar="YES"
case "n"
	P3Jeopardychar="NO"
case else
	P3Jeopardychar="NO"
end select
	

sqlrsv="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'R'"
	ToRRsvRec.setSQLText(sqlrsv)
	ToRRsvRec.open

sqlass="Select * From xca_Part1 Where Tix= '"&Tix&"' And TypeOfRequest = 'A'"
	ToRAssRec.setSQLText(sqlass)
	ToRAssRec.open

%>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black" leftmargin=15 rightmargin=20>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserName\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\r\n\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part1\sWHERE\sxca_Part1.Tix\s=\s?\sAND\sxca_User.UserName\s=\s?\sAND\sxca_Entity.EntityID\s=\s?\r\n\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q1\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qServer\s(ASP)\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=3,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q35\q,CReq=0),Row3=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam3\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetUserEntityName()
{
}
function _initGetUserEntityName()
{
	GetUserEntityName.advise(RS_ONBEFOREOPEN, _setParametersGetUserEntityName);
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_User, xca_Part1 WHERE xca_Part1.Tix = ? AND xca_User.UserName = ? AND xca_Entity.EntityID = ? ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetUserEntityName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetUserEntityName') != null)
		GetUserEntityName.setBookmark(thisPage.getState('pb_GetUserEntityName'));
}
function _GetUserEntityName_ctor()
{
	CreateRecordset('GetUserEntityName', _initGetUserEntityName, null);
}
function _GetUserEntityName_dtor()
{
	GetUserEntityName._preserveState();
	thisPage.setState('pb_GetUserEntityName', GetUserEntityName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_Parms\sWHERE\s(xca_Entity.EntityName\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminEntityName()
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminUserName 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sUserName\sfrom\sxca_User\swhere\sEntityID=?\q,TCControlID_Unmatched=\qGetAdminUserName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_User\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sUserName\sfrom\sxca_User\swhere\sEntityID=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
	cmdTmp.CommandText = 'Select UserName from xca_User where EntityID=?';
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
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxca_Part1.Tix\s=?\q,TCControlID_Unmatched=\qGetPart1Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_part1\swhere\sxca_Part1.Tix\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1Data()
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
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Part3\swhere\sTix=?\q,TCControlID_Unmatched=\qGetPart3Data\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Part1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Part3\swhere\sTix=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart3Data()
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 
id=GetCOCodeData style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" 
width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_COCode\sWHERE\sTix\s=\s?\sAND\sNPASplitID\s=\s'Included'\q,TCControlID_Unmatched=\qGetCOCodeData\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_COCode\sWHERE\sTix\s=\s?\sAND\sNPASplitID\s=\s'Included'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetCOCodeData()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_COCode WHERE Tix = ? AND NPASplitID = \'Included\'';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 
id=CodeAssignRec style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" 
width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID,\sxca_Part1.RateCenter\sAS\sRateCenter\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'a')\q,TCControlID_Unmatched=\qCodeAssignRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID,\sxca_Part1.RateCenter\sAS\sRateCenter\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'a')\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCodeAssignRec()
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 
id=CodeResvRec style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" 
width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'r')\q,TCControlID_Unmatched=\qCodeResvRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Part3.*,\sxca_Part1.SwitchID\sAS\sSwitchID\sFROM\sxca_Part3\sINNER\sJOIN\sxca_Part1\sON\sxca_Part3.Tix\s=\sxca_Part1.Tix\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_Part3.Part3Result\s=\s'r')\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCodeResvRec()
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
	<PARAM NAME="ExtentX" VALUE="12197">
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ToRRsvRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
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

<hr size=4 align=left color=maroon width=71.58% style="HEIGHT: 4px; WIDTH: 680px">
<table align="left" cellPadding="0" cellSpacing="0" width="70.51%" border=0 height=60 style="HEIGHT: 60px; WIDTH: 617px">
    <TR>
		<TD align=left><strong><font face="Arial Black" color=maroon size="4">Part 
            3: Canadian CNA's Response/Confirmation Form</font></strong>
	<tr>
        <TD align=left><strong><font face="Arial" color=maroon size="4">
        CNA Ticket #:&nbsp;&nbsp;
        <%Response.Write(Tix)%></font></strong>
        </td>
	<tr>
        <TD align=left><font color="maroon" face="Arial" size="4"><strong>
		Request Status:&nbsp;&nbsp;
        <% Response.Write RequestStatuschar %></font></strong>
		</td>
	</tr>
</table>&nbsp;
<br><br>
<P></P>
<P>&nbsp;</P>
<P>
<hr size=4 align=left color=maroon width=71.68% style="HEIGHT: 4px; WIDTH: 681px">

<P></P>
<table align="center" background ="" border="0" cellPadding="1" cellSpacing="1" height="272" style="HEIGHT: 272px">
    <tbody>
    
    <tr>
        <td align="left" colSpan="2" noWrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code Applicant 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" noWrap><font face="Arial"> </font></td>
        <td align="left" colSpan="2" noWrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity 
            Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AppEntityname 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityname">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityname()
{
	AppEntityname.setDataSource(GetUserEntityName);
	AppEntityname.setDataField('EntityName');
}
function _AppEntityname_ctor()
{
	CreateLabel('AppEntityname', _initAppEntityname, null);
}
</script>
<% AppEntityname.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"> 
            </font></font>CNA Admin</font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><strong><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CO Code Manager</font></strong><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>

</FONT></B>


<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityContact 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 864px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityContact">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityContact()
{
	AppEntityContact.setDataSource(GetUserEntityName);
	AppEntityContact.setDataField('UserName');
}
function _AppEntityContact_ctor()
{
	CreateLabel('AppEntityContact', _initAppEntityContact, null);
}
</script>
<% AppEntityContact.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->

        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityContact 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 83px" width=83>
	<PARAM NAME="_ExtentX" VALUE="2196">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->

</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AppEntityAddress 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityAddress()
{
	AppEntityAddress.setDataSource(GetUserEntityName);
	AppEntityAddress.setDataField('EntityAddress');
}
function _AppEntityAddress_ctor()
{
	CreateLabel('AppEntityAddress', _initAppEntityAddress, null);
}
</script>
<% AppEntityAddress.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityAddress 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 86px" width=86>
	<PARAM NAME="_ExtentX" VALUE="2275">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
</OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AppEntityCity 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityCity">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityCity()
{
	AppEntityCity.setDataSource(GetUserEntityName);
	AppEntityCity.setDataField('EntityCity');
}
function _AppEntityCity_ctor()
{
	CreateLabel('AppEntityCity', _initAppEntityCity, null);
}
</script>
<% AppEntityCity.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityCity 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AppEntityProvince 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityProvince()
{
	AppEntityProvince.setDataSource(GetUserEntityName);
	AppEntityProvince.setDataField('EntityProvince');
}
function _AppEntityProvince_ctor()
{
	CreateLabel('AppEntityProvince', _initAppEntityProvince, null);
}
</script>
<% AppEntityProvince.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityProvince 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 87px" width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
    </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
         
            
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal 
            Code</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AppEntityPostalCode 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
</OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityPostalCode()
{
	AppEntityPostalCode.setDataSource(GetUserEntityName);
	AppEntityPostalCode.setDataField('EntityPostalCode');
}
function _AppEntityPostalCode_ctor()
{
	CreateLabel('AppEntityPostalCode', _initAppEntityPostalCode, null);
}
</script>
<% AppEntityPostalCode.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal Code 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityPostalCode 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail Address 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityEmail 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1064px; WIDTH: 64px" width=64>
	<PARAM NAME="_ExtentX" VALUE="1693">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityEmail()
{
	AppEntityEmail.setDataSource(GetUserEntityName);
	AppEntityEmail.setDataField('UserEmail');
}
function _AppEntityEmail_ctor()
{
	CreateLabel('AppEntityEmail', _initAppEntityEmail, null);
}
</script>
<% AppEntityEmail.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityEmail 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 71px" width=71>
	<PARAM NAME="_ExtentX" VALUE="1879">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>    
            
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityFax 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1104px; WIDTH: 53px" width=53>
	<PARAM NAME="_ExtentX" VALUE="1402">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityFax">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityFax()
{
	AppEntityFax.setDataSource(GetUserEntityName);
	AppEntityFax.setDataField('UserFax');
}
function _AppEntityFax_ctor()
{
	CreateLabel('AppEntityFax', _initAppEntityFax, null);
}
</script>
<% AppEntityFax.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityFax 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
          
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityTelephone 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1144px; WIDTH: 90px" width=90>
	<PARAM NAME="_ExtentX" VALUE="2381">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityTelephone()
{
	AppEntityTelephone.setDataSource(GetUserEntityName);
	AppEntityTelephone.setDataField('UserTelephone');
}
function _AppEntityTelephone_ctor()
{
	CreateLabel('AppEntityTelephone', _initAppEntityTelephone, null);
}
</script>
<% AppEntityTelephone.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityTelephone 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
            
</td></tr>
    <tr>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=AppEntityExtension 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 1184px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AppEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityExtension()
{
	AppEntityExtension.setDataSource(GetUserEntityName);
	AppEntityExtension.setDataField('UserExtension');
}
function _AppEntityExtension_ctor()
{
	CreateLabel('AppEntityExtension', _initAppEntityExtension, null);
}
</script>
<% AppEntityExtension.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" noWrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AdminEntityExtension 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Maroon">
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
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr></tbody>
</table><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
<p></p>

<hr>

<p><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
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
<% 
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("ApplicationDate"),vbShortDate)
'DateofApp.display
 %>
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
<% 
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("DateofReceipt"),vbShortDate)
'DateofReceipt.display
%>
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
<% 
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("DateResponseDue"),vbShortDate)
'DateofResponse.display
 %>
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
<% 
' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format
response.write FormatDateTime(GetPart1Data.fields.getValue("RequestedEffDate"),vbShortDate)
'RequestedEffDate.display
 %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right" noWrap></td>
        <td align="left" noWrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="right" noWrap>
            <div align="right"><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>The Preferred NPA-NXX Split 
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
            id=Label1 style="HEIGHT: 20px; LEFT: 10px; TOP: 1324px; WIDTH: 96px" 
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
	<PARAM NAME="DataSource" VALUE="">
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
<% ' KT CHANGED 2013-06-12:  Skip databound control display and just write out date in spec format; 
     ' THIS COULD BE BLANK (formatter doesn't like blank so do if statement
if CodeResvRec.fields.getValue("ExtentionDate")<>"" then
response.write FormatDateTime(CodeResvRec.fields.getValue("ExtentionDate"),vbShortDate)
end if
'ExtensionDate.display 
%>
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
            Part1 Request</strong></FONT></FONT> 
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
            size=2><STRONG>Secondary NXXs 
            chosen that are sill available:</STRONG></FONT> 
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
            size=2><STRONG>Secondary NXXs 
            chosen that are sill available:</STRONG></FONT> 
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
<font face=arial size=1>(dd/mm/ccyy)<font face=Arial size=2></font></font></font></font></font> </td></tr>
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
            </strong></font><FONT face=Arial><strong> 
            Part 1 Form completed, Code request denied. 
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
            </strong></font><FONT face=Arial><strong> 
            Part 1 Assignment Activity Suspended by the 
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
    <tr>
        <td align=left colSpan=12>&nbsp;&nbsp; </td></tr>
    <TR>
        <TD align=left colSpan=12><a HREF="#top"> Back to Top</a>
        </TD>
    <tr>
        <td align=left colSpan=12>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMenu 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 2124px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMenu">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
		</td>
	</tr>
</table> 


</form> 
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

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

UserEntityID=session("UserEntityID")
UserID=session("UserUserID")
NPA=session("SelectedNPA")
SelectedNXX=session("SelectedNXX")
AuthorizedRep=Replace(Request.Form("AuthorizedRep"),"'","''")
' KT CHANGED 2013-06-14:  
ApplicationDate=FormatDateTime(Request.Form("ApplicationDate"),vbShortDate)
ESRD=Request.Form("ESRDAssign")
session("ESRD")=ESRD
SyncField=now()
SQLget= "Select Tix from xca_ESRD_Part1 where NPA='"&NPA&"' and NXX='"&SelectedNXX&"' and ESRD='"&ESRD&"'"
P1DataCon.setSQLText(SQLget)
P1DataCon.open
RetrieveTix = P1DataCon.fields.getValue("Tix")
session("Tix")=RetrieveTix
P1DataCon.close
''''''''''''''''''''''''''''''''''''''''''''''''''	
Set objConn=server.CreateObject("ADODB.Connection")
Set objCmd=server.CreateObject("ADODB.Command")
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn
SQLstmt = "INSERT INTO xca_ESRD_Part3 (Tix,AuthorizedRep,ApplicationDate,NPA,NXX,ESRD,SyncField, UserID, EntityID) VALUES ('"&RetrieveTix&"', '"&AuthorizedRep&"','"&ApplicationDate&"','"&NPA&"','"&SelectedNXX&"','"&ESRD&"','"&SyncField&"', '"&UserID&"', '"&UserEntityID &"')"
objCmd.CommandText=SQLstmt
objCmd.Execute
'Update ESRD with tix#
SQLstmt1="Update xca_ESRD Set Status='I' where Tix='"&RetrieveTix&"'"
response.write SQLstmt1
	objCmd.CommandText=SQLstmt1
	objCmd.Execute
SQLstmt2="Update xca_ESRD_Part1 Set Status='Closed' where Tix='"&RetrieveTix&"'"
	objCmd.CommandText=SQLstmt2
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

Response.Redirect "ESRD_Part3Confirm.asp"

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

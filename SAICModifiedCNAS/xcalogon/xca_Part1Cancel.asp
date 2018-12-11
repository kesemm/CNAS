<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="xca_CNASLib.inc"-->
<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</HEAD>
<%
Tix=Request.Form("P1CancelTix")

UserEntityID=session("UserEntityId")
UserUserType = session("UserUserType")
UserUserID = cint(session("UserUserID"))
'Check for invalid tix in COCode - only-status=Q
If session("UserEntityType")<> "a" then
  If  session("UserUserType") = "m" then
	sqlnoTix="Select * from xca_COCode where Tix= '"&Tix&"' and EntityID='"&UserEntityID&"' and Status='Q'"
	GetCOCodeData.setSQLText(sqlnoTix)
	GetCOCodeData.Open
	checkTIX= GetCOCodeData.fields.getValue("Tix")
	CanNPA = GetCOCodeData.fields.getValue("NPA")
	CanNXX = GetCOCodeData.fields.getValue("NXX")
	session("CanNPA")=CanNPA
	session("CanNXX")=CanNXX
	session("P1CanTix")=Tix
  Else
	sqlnoTix="Select * from xca_COCode,xca_Part1 where xca_COCode.Tix= '"&Tix&"' and xca_COCode.EntityID='"&UserEntityID&"'and xca_Part1.UserID='"&UserUserID&"' and xca_COCode.Status='Q' and xca_COCode.Tix=xca_Part1.Tix"
	GetCOCodeData.setSQLText(sqlnoTix)
	GetCOCodeData.Open
	checkTIX= GetCOCodeData.fields.getValue("Tix")
	CanNPA = GetCOCodeData.fields.getValue("NPA")
	CanNXX = GetCOCodeData.fields.getValue("NXX")
	session("CanNPA")=CanNPA
	session("CanNXX")=CanNXX
	session("P1CanTix")=Tix
  End IF
else 
  If  session("UserUserType") = "m" then
	sqlnoTix="Select * from xca_COCode where Tix= '"&Tix&"' and Status='Q'"
	GetCOCodeData.setSQLText(sqlnoTix)
	GetCOCodeData.Open
	checkTIX= GetCOCodeData.fields.getValue("Tix")
	CanNPA = GetCOCodeData.fields.getValue("NPA")
	CanNXX = GetCOCodeData.fields.getValue("NXX")
	session("CanNPA")=CanNPA
	session("CanNXX")=CanNXX
	session("P1CanTix")=Tix
  Else
    sqlnoTix="Select * from xca_COCode,xca_Part1 where xca_COCode.Tix= '"&Tix&"' and xca_Part1.UserID='"&UserUserID&"' and xca_COCode.Status='Q' and xca_COCode.Tix=xca_Part1.Tix"
	GetCOCodeData.setSQLText(sqlnoTix)
	GetCOCodeData.Open
	checkTIX= GetCOCodeData.fields.getValue("Tix")
	CanNPA = GetCOCodeData.fields.getValue("NPA")
	CanNXX = GetCOCodeData.fields.getValue("NXX")
	session("CanNPA")=CanNPA
	session("CanNXX")=CanNXX
	session("P1CanTix")=Tix
  End If
end if 	

if checkTix="" then	
session("NoTixSent")="DidNotSend"
Response.Redirect session("Here")

end if


session("NoTixSent")=""

%>
<script LANGUAGE="JavaScript">

        <!-- Hide code from non-js browsers
var doCancel 
	doCancel=confirm("Do you want to delete this ticket?????")
	if (doCancel)	{
		location.href = 'xca_Part1CancelPost.asp';
				}
	else {
		
		location.href = 'xca_MenuMainPost.asp';
	}

// end hiding -->
</script>	


<BODY>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetCOCodeData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject=\qTables\q,TCDBObjectName=\qxca_Entity\q,TCControlID_Unmatched=\qGetCOCodeData\q,TCPPConn=\qcnasadmin\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
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
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '"xca_Entity"';
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


</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>



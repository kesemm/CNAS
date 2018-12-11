<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>

<HTML>
<HEAD>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%

Dim Authenticated, UserStatus, EntityStatus
UserName = Replace(Request.Form("UID"),"'","''")
UserPass = Replace(Request.Form("PWD"),"'","''")


 
If UserName = "" or UserPass = "" then
		
        Response.Redirect("xca_Login2.asp")
  else
  
    Call Authenticate(UserName, UserPass)
    Call GetData(UserStatus, EntityStatus)    
    
     
end if

Sub Authenticate(UID, PWD)
    sql = "Select * from xca_User where UserLogon = '"&UID&"' and UserPassword = '"&PWD&"'"
    GetUserUserData.setSQLText(sql)
	GetUserUserData.open
	''CLose RS to to Test for DB Failure
	'GetUserUserData.close 
    IF (GetUserUserData.isOpen()) THEN 'Check for DB to be open
	
    
     If GetUserUserData.getCount() > 0 then
				UserStatus = GetUserUserData.fields.getValue("UserStatus")
				session("UserEntityID")=GetUserUserData.fields.getValue("EntityID")
				session("UserUserID")=GetUserUserData.fields.getValue("UserID")
				session("UserUserName")=GetUserUserData.fields.getValue("UserName")
				session("UserUserEmail")=GetUserUserData.fields.getValue("UserEmail")
				session("UserUserLogon")=GetUserUserData.fields.getValue("UserLogon")
				session("UserUserType")=GetUserUserData.fields.getValue("UserType")
				session("UserUserStatus")=UserStatus
				 EntityID=session("UserEntityID")
     'get Entity Data  
		sql1 = "Select * from xca_Entity where xca_Entity.EntityID = '" &EntityID&"'"
				 GetUserEntityData.setSQLText(sql1)
				 GetUserEntityData.open       
       					session("UserEntityType") = GetUserEntityData.fields.getValue("EntityType")
       					session("EntityUserEmail") = GetUserEntityData.fields.getValue("EntityEmail")
       					EntityStatus = GetUserEntityData.fields.getValue("EntityStatus")
	'get Administrator form data
		sqlparm= "select * from xca_Parms where Name = 'ADMIN'"
				GetParmData.setSQLText(sqlparm)
				GetParmData.open
				AdminData=GetParmData.fields.getValue("Value")
				session("ADMIN")=AdminData
	
		sql2 = "Select * from xca_Entity  where EntityName = '"&AdminData&"'"
				GetUserEntityData.setSQLText(sql2)
				GetUserEntityData.open       
       					session("AdminEntityEmail") = GetUserEntityData.fields.getValue("EntityEmail")
       								
    else
       UserStatus = "No"
       EntityStatus="No"
    end if
   Else
	Response.Redirect ("xca_SysFailure.asp")
	
   END IF 'End of check for DB
end sub

Sub GetData(UStat, EStat)
	IF (GetUserUserData.isOpen = false) THEN 'Check for DB to be open
		Response.Redirect ("xca_SysFailure.asp")
	Else
	If UStat = "a" and EStat="a" then
Response.Redirect("xcalogon/xca_MenuMain.asp")

	else
        Response.Redirect("xca_Login2.asp")
Response.Write UserStatus
Response.Write EntityStatus		   
       
    end if
   end if
end sub


%>


</HEAD>

<BODY>


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserUserData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12192">
	<PARAM NAME="ExtentY" VALUE="1799">
	<PARAM NAME="State" VALUE="(TCConn=\qcnaslogon\q,TCDBObject=\qTables\q,TCDBObjectName=\qxca_User\q,TCControlID_Unmatched=\qGetUserUserData\q,TCPPConn=\qcnaslogon\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_User\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<!--#INCLUDE FILE="_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserUserData()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnaslogon_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnaslogon_CommandTimeout');
	DBConn.CursorLocation = Application('cnaslogon_CursorLocation');
	DBConn.Open(Application('cnaslogon_ConnectionString'), Application('cnaslogon_RuntimeUserName'), Application('cnaslogon_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '"xca_User"';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetUserUserData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetUserUserData') != null)
		GetUserUserData.setBookmark(thisPage.getState('pb_GetUserUserData'));
}
function _GetUserUserData_ctor()
{
	CreateRecordset('GetUserUserData', _initGetUserUserData, null);
}
function _GetUserUserData_dtor()
{
	GetUserUserData._preserveState();
	thisPage.setState('pb_GetUserUserData', GetUserUserData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12192">
	<PARAM NAME="ExtentY" VALUE="1799">
	<PARAM NAME="State" VALUE="(TCConn=\qcnaslogon\q,TCDBObject=\qTables\q,TCDBObjectName=\qxca_Entity\q,TCControlID_Unmatched=\qGetUserEntityData\q,TCPPConn=\qcnaslogon\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityData()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnaslogon_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnaslogon_CommandTimeout');
	DBConn.CursorLocation = Application('cnaslogon_CursorLocation');
	DBConn.Open(Application('cnaslogon_ConnectionString'), Application('cnaslogon_RuntimeUserName'), Application('cnaslogon_RuntimePassword'));
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
	GetUserEntityData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetUserEntityData') != null)
		GetUserEntityData.setBookmark(thisPage.getState('pb_GetUserEntityData'));
}
function _GetUserEntityData_ctor()
{
	CreateRecordset('GetUserEntityData', _initGetUserEntityData, null);
}
function _GetUserEntityData_dtor()
{
	GetUserEntityData._preserveState();
	thisPage.setState('pb_GetUserEntityData', GetUserEntityData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetParmData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject=\qTables\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qGetParmData\q,TCPPConn=\qcnasadmin\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetParmData()
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
	GetParmData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetParmData') != null)
		GetParmData.setBookmark(thisPage.getState('pb_GetParmData'));
}
function _GetParmData_ctor()
{
	CreateRecordset('GetParmData', _initGetParmData, null);
}
function _GetParmData_dtor()
{
	GetParmData._preserveState();
	thisPage.setState('pb_GetParmData', GetParmData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

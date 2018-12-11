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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>



function CleanUp(Item,CutDate)
	
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	'on error resume next
	
	select Case Item
		case "Closed Ticket"
			objCmd.CommandText=	"CleanUpTickets '" & CutDate & "'"
		
		case "Completed Splits"
			objCmd.CommandText=	"CleanUpSplits '" & CutDate & "'"	
			
		case "Completed Transfers"
			objCmd.CommandText=	"CleanUpTransfers '" & CutDate & "'"
			
		case "Old Logs"
			objCmd.CommandText=	"CleanUpLogs '" & CutDate & "'"
	end select  
	
	objCmd.Execute
	if objConn.Errors.Count = 0 then 
		CleanUpTemp=true
	else
		CleanUpTemp=false	
	end if	
	objConn.close
	Set objConn=Nothing
	Set objCmd=Nothing
	CleanUp=CleanUpTemp
	
end function

Sub btnShow_onclick()
	if not IsDateReal(txtCleanUpDate.value) then
		session("CleanUpError")="InvalidDateFormat"
	else	
		session("CleanUpError")="Show"
		txt=lstCleanUpItem.selectedIndex
		Item = lstCleanUpItem.getText(txt)
	
		SQL1 = "select tix, RequestStatus, ApplicationDate  from xca_Part1 where CompletionDate <='" & txtCleanUpDate.Value & "' order by CompletionDate "
		SQL2 = "select * from xca_NPASplit where Status='C' And NPASplitComplete <='" & txtCleanUpDate.Value & "' order by NPASplitComplete "
		SQL3 = "select * from xca_Transfers where Status='C' And TransferDate <='" & txtCleanUpDate.Value & "' order by TransferDate "
		SQL4 = "SELECT a.LogType, a.NPA, a.NXX, b.UserLogon,a.Date1 AS Log_Date, a.Tix, a.Action, a.ActionText,a.Process FROM xca_Logs a INNER JOIN xca_User b ON a.UserID = b.UserID where a.Date1 <='" & txtCleanUpDate.Value & "' order by a.Date1 "

		if ClosedTicket.isopen() then ClosedTicket.close()
		if CompletedSplits.isopen() then CompletedSplits.close()
		if CompletedTransfers.isopen() then CompletedTransfers.close()
		if OldLogs.isopen() then OldLogs.close()
		select Case Item
			case "Closed Ticket"
				ClosedTicket.setSQLText(SQL1)
				ClosedTicket.open()
			case "Completed Splits"
				CompletedSplits.setSQLText(SQL2)
				CompletedSplits.open()
			case "Completed Transfers"
				CompletedTransfers.setSQLText(SQL3)
				CompletedTransfers.open()
			case "Old Logs"
				OldLogs.setSQLText(SQL4)
				OldLogs.open()
		end select  
	end if	
End Sub

sub btnReturnToMain_onclick()
	session("CleanUpDate")
	Response.Redirect "xca_MenuSecurityAdmin.asp"
end sub

Sub btnCleanUp_onclick()
	if not IsDateReal(txtCleanUpDate.value) then
		session("CleanUpError")="InvalidDateFormat"
	elseif date()- cDate(txtCleanUpDate.value) < 90 then
		session("CleanUpError")="CleaUpDateTooShort"
	'elseif date()- cDate(txtCleanUpDate.value) < 365 then
	'	session("CleanUpError")="CleaUpDateShort"	'for message ???
	else
		session("CleanUpError")="CleanUp"
		txt=lstCleanUpItem.selectedIndex
		Item = lstCleanUpItem.getText(txt)
		select Case Item
			case "Closed Ticket"
				Action="Ticket"
			case "Completed Splits"
				Action="Split"
			case "Completed Transfers"
				Action="Transfer"
			case "Old Logs"
				Action="Log"
		end select  
		
		'Response.write txtCleanUpDate.value
		
		if CleanUp(Item,txtCleanUpDate.value) then
		
			log "C","","",session("UserUserID"),Now,0,Action,txtCleanUpDate.value,"CleanUp" 
			'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
			'email session("AdminEntityEMail"),EMailTo,"","System Cleanup Performed", "System cleanup performed on " & date 
			session("CleanUp")="Yes"
			btnShow_onclick
		end if	
	end if	

End Sub

</SCRIPT>

<%	
	if session("CleanUpError")="InvalidDateFormat" then
		session("CleanUpError")=""
	%>
		<SCRIPT Language="JavaScript">
		alert("Invalid Clean Up Date format.")
		</SCRIPT>
 	<% 
	elseif session("CleanUpError")="CleaUpDateTooShort" then
		session("CleanUpError")=""
	%>
		<SCRIPT Language="JavaScript">
		alert("Clean Up Date can not be less than 90 days before the current date.")
		</SCRIPT>
 	<% 
 	elseif session("CleanUp")="Yes" then
 		session("CleanUp")=""
 	%>
		<SCRIPT Language="JavaScript">
		alert("Clean up performed successfully.")
		</SCRIPT>
 	<% 
	elseif (session("CleanUpError")="Show") or (session("CleanUpError")="CleanUp") then
		session("CleanUpError")=""
	elseif session("CleanUpDate")="" then	
		txtCleanUpDate.value=date()-365
		session("CleanUpDate")=txtCleanUpDate.value
	end if	
	session("CleanUp")=""		
%>

</HEAD>
<BODY  bgColor="#d7c7a4">

<TABLE WIDTH=75% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD align=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 id=lblTitle 
	style="HEIGHT: 37px; LEFT: 10px; TOP: 34px; WIDTH: 324px" width=324>
	<PARAM NAME="_ExtentX" VALUE="8573">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="System Clean Up Form">
	<PARAM NAME="FontFace" VALUE="Arial Black">
	<PARAM NAME="FontSize" VALUE="5">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial Black" SIZE="5" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTitle()
{
	lblTitle.setCaption('System Clean Up Form');
}
function _lblTitle_ctor()
{
	CreateLabel('lblTitle', _initlblTitle, null);
}
</script>
<% lblTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><BR>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=OldLogs1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sa.LogType,\sa.NPA,\sa.NXX,\sb.UserLogon,a.Date1\sAS\sLog_Date,\sa.Tix,\sa.Action,\sa.ActionText,a.Process\sFROM\sxca_Logs\sa\sINNER\sJOIN\sxca_User\sb\sON\sa.UserID\s=\sb.UserID\s\q,TCControlID_Unmatched=\qOldLogs1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Logs\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sa.LogType,\sa.NPA,\sa.NXX,\sb.UserLogon,a.Date1\sAS\sLog_Date,\sa.Tix,\sa.Action,\sa.ActionText,a.Process\sFROM\sxca_Logs\sa\sINNER\sJOIN\sxca_User\sb\sON\sa.UserID\s=\sb.UserID\s\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initOldLogs1()
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
	cmdTmp.CommandText = 'SELECT a.LogType, a.NPA, a.NXX, b.UserLogon,a.Date1 AS Log_Date, a.Tix, a.Action, a.ActionText,a.Process FROM xca_Logs a INNER JOIN xca_User b ON a.UserID = b.UserID ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	OldLogs1.setRecordSource(rsTmp);
	if (thisPage.getState('pb_OldLogs1') != null)
		OldLogs1.setBookmark(thisPage.getState('pb_OldLogs1'));
}
function _OldLogs1_ctor()
{
	CreateRecordset('OldLogs1', _initOldLogs1, null);
}
function _OldLogs1_dtor()
{
	OldLogs1._preserveState();
	thisPage.setState('pb_OldLogs1', OldLogs1.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=CompletedTransfers 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qCompletedTransfers\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Transfers\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCompletedTransfers()
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
//Recordset DTC error: Failed to get command text
	cmdTmp.CommandText = '';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	CompletedTransfers.setRecordSource(rsTmp);
	if (thisPage.getState('pb_CompletedTransfers') != null)
		CompletedTransfers.setBookmark(thisPage.getState('pb_CompletedTransfers'));
}
function _CompletedTransfers_ctor()
{
	CreateRecordset('CompletedTransfers', _initCompletedTransfers, null);
}
function _CompletedTransfers_dtor()
{
	CompletedTransfers._preserveState();
	thisPage.setState('pb_CompletedTransfers', CompletedTransfers.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=CompletedSplits style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qCompletedSplits\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCompletedSplits()
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
//Recordset DTC error: Failed to get command text
	cmdTmp.CommandText = '';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	CompletedSplits.setRecordSource(rsTmp);
	if (thisPage.getState('pb_CompletedSplits') != null)
		CompletedSplits.setBookmark(thisPage.getState('pb_CompletedSplits'));
}
function _CompletedSplits_ctor()
{
	CreateRecordset('CompletedSplits', _initCompletedSplits, null);
}
function _CompletedSplits_dtor()
{
	CompletedSplits._preserveState();
	thisPage.setState('pb_CompletedSplits', CompletedSplits.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ClosedTicket style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\q\r\n\q,TCControlID_Unmatched=\qClosedTicket\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\q\r\n\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initClosedTicket()
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
	cmdTmp.CommandText = ' ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ClosedTicket.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ClosedTicket') != null)
		ClosedTicket.setBookmark(thisPage.getState('pb_ClosedTicket'));
}
function _ClosedTicket_ctor()
{
	CreateRecordset('ClosedTicket', _initClosedTicket, null);
}
function _ClosedTicket_dtor()
{
	ClosedTicket._preserveState();
	thisPage.setState('pb_ClosedTicket', ClosedTicket.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=OldLogs style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sa.LogType,\sa.NPA,\sa.NXX,\sb.UserLogon,a.Date1\sAS\sLog_Date,\sa.Tix,\sa.Action,\sa.ActionText,a.Process\sFROM\sxca_Logs\sa\sINNER\sJOIN\sxca_User\sb\sON\sa.UserID\s=\sb.UserID\q,TCControlID_Unmatched=\qOldLogs\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sa.LogType,\sa.NPA,\sa.NXX,\sb.UserLogon,a.Date1\sAS\sLog_Date,\sa.Tix,\sa.Action,\sa.ActionText,a.Process\sFROM\sxca_Logs\sa\sINNER\sJOIN\sxca_User\sb\sON\sa.UserID\s=\sb.UserID\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initOldLogs()
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
	cmdTmp.CommandText = 'SELECT a.LogType, a.NPA, a.NXX, b.UserLogon,a.Date1 AS Log_Date, a.Tix, a.Action, a.ActionText,a.Process FROM xca_Logs a INNER JOIN xca_User b ON a.UserID = b.UserID';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	OldLogs.setRecordSource(rsTmp);
	if (thisPage.getState('pb_OldLogs') != null)
		OldLogs.setBookmark(thisPage.getState('pb_OldLogs'));
}
function _OldLogs_ctor()
{
	CreateRecordset('OldLogs', _initOldLogs, null);
}
function _OldLogs_dtor()
{
	OldLogs._preserveState();
	thisPage.setState('pb_OldLogs', OldLogs.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<TABLE WIDTH=50% BORDER=0 CELLSPACING=1 CELLPADDING=1 align=center>
<TR>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=lblCleanUpItem 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 466px; WIDTH: 94px" width=94>
	<PARAM NAME="_ExtentX" VALUE="2487">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="lblCleanUpItem">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Clean Up Item">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblCleanUpItem()
{
	lblCleanUpItem.setCaption('Clean Up Item');
}
function _lblCleanUpItem_ctor()
{
	CreateLabel('lblCleanUpItem', _initlblCleanUpItem, null);
}
</script>
<% lblCleanUpItem.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=lstCleanUpItem 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 139px" width=139>
	<PARAM NAME="_ExtentX" VALUE="3678">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="lstCleanUpItem">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="4">
	<PARAM NAME="CLED1" VALUE="Closed Ticket">
	<PARAM NAME="CLEV1" VALUE="1">
	<PARAM NAME="CLED2" VALUE="Completed Splits">
	<PARAM NAME="CLEV2" VALUE="2">
	<PARAM NAME="CLED3" VALUE="Completed Transfers">
	<PARAM NAME="CLEV3" VALUE="3">
	<PARAM NAME="CLED4" VALUE="Old Logs">
	<PARAM NAME="CLEV4" VALUE="4">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstCleanUpItem()
{
	lstCleanUpItem.addItem('Closed Ticket', '1');
	lstCleanUpItem.addItem('Completed Splits', '2');
	lstCleanUpItem.addItem('Completed Transfers', '3');
	lstCleanUpItem.addItem('Old Logs', '4');
}
function _lstCleanUpItem_ctor()
{
	CreateListbox('lstCleanUpItem', _initlstCleanUpItem, null);
}
</script>
<% lstCleanUpItem.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=lblCleanUpDate 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 507px; WIDTH: 95px" width=95>
	<PARAM NAME="_ExtentX" VALUE="2514">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="lblCleanUpDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Clean Up Date">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblCleanUpDate()
{
	lblCleanUpDate.setCaption('Clean Up Date');
}
function _lblCleanUpDate_ctor()
{
	CreateLabel('lblCleanUpDate', _initlblCleanUpDate, null);
}
</script>
<% lblCleanUpDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtCleanUpDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtCleanUpDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtCleanUpDate()
{
	txtCleanUpDate.setStyle(TXT_TEXTBOX);
	txtCleanUpDate.setMaxLength(10);
	txtCleanUpDate.setColumnCount(10);
}
function _txtCleanUpDate_ctor()
{
	CreateTextbox('txtCleanUpDate', _inittxtCleanUpDate, null);
}
</script>
<% txtCleanUpDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><BR>
<TABLE WIDTH=10% ALIGN=center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnShow style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 81px" 
	width=81>
	<PARAM NAME="_ExtentX" VALUE="2143">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnShow">
	<PARAM NAME="Caption" VALUE="Show List">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnShow()
{
	btnShow.value = 'Show List';
	btnShow.setStyle(0);
}
function _btnShow_ctor()
{
	CreateButton('btnShow', _initbtnShow, null);
}
</script>
<% btnShow.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD> <TD ALIGN=center>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnCleanUp 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 80px" width=80>
	<PARAM NAME="_ExtentX" VALUE="2117">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnCleanUp">
	<PARAM NAME="Caption" VALUE="Clean Up">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnCleanUp()
{
	btnCleanUp.value = 'Clean Up';
	btnCleanUp.setStyle(0);
}
function _btnCleanUp_ctor()
{
	CreateButton('btnCleanUp', _initbtnCleanUp, null);
}
</script>
<% btnCleanUp.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD ALIGN=center>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMain 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 600px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMain">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturnToMain()
{
	btnReturnToMain.value = 'Return';
	btnReturnToMain.setStyle(0);
}
function _btnReturnToMain_ctor()
{
	CreateButton('btnReturnToMain', _initbtnReturnToMain, null);
}
</script>
<% btnReturnToMain.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>

<BR>
<TABLE WIDTH=75% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid2 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 692px" 
	width=692>
	<PARAM NAME="_ExtentX" VALUE="18309">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="ClosedTicket">
	<PARAM NAME="CtrlName" VALUE="Grid2">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"tix","RequestStatus","ApplicationDate"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="218,229,241">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"tix","RequestStatus","ApplicationDate"'>
	<PARAM NAME="DetailAlignment" VALUE=",,">
	<PARAM NAME="HeaderAlignment" VALUE=",,">
	<PARAM NAME="DetailBackColor" VALUE=",,">
	<PARAM NAME="HeaderBackColor" VALUE=",,">
	<PARAM NAME="HeaderFont" VALUE=",,">
	<PARAM NAME="HeaderFontColor" VALUE=",,">
	<PARAM NAME="HeaderFontSize" VALUE=",,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,">
	<PARAM NAME="DetailFont" VALUE=",,">
	<PARAM NAME="DetailFontColor" VALUE=",,">
	<PARAM NAME="DetailFontSize" VALUE=",,">
	<PARAM NAME="DetailFontStyle" VALUE=",,">
	<PARAM NAME="ColumnCount" VALUE="3">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="2">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="10">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="692">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid2()
{
Grid2.pageSize = 10;
Grid2.setDataSource(ClosedTicket);
Grid2.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=692';
Grid2.headerAttributes = '   bgcolor=Maroon align=Center';
Grid2.headerWidth[0] = ' WIDTH=218';
Grid2.headerWidth[1] = ' WIDTH=229';
Grid2.headerWidth[2] = ' WIDTH=241';
Grid2.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid2.colHeader[0] = '\'tix\'';
Grid2.colHeader[1] = '\'RequestStatus\'';
Grid2.colHeader[2] = '\'ApplicationDate\'';
Grid2.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid2.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid2.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid2.colAttributes[0] = '  WIDTH=218';
Grid2.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid2.colData[0] = 'ClosedTicket.fields.getValue(\'tix\')';
Grid2.colAttributes[1] = '  WIDTH=229';
Grid2.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid2.colData[1] = 'ClosedTicket.fields.getValue(\'RequestStatus\')';
Grid2.colAttributes[2] = '  WIDTH=241';
Grid2.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid2.colData[2] = 'ClosedTicket.fields.getValue(\'ApplicationDate\')';
Grid2.navbarAlignment = 'Right';
var objPageNavbar = Grid2.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid2.hasPageNumber = true;
}
function _Grid2_ctor()
{
	CreateDataGrid('Grid2',_initGrid2);
}
</SCRIPT>

<%	Grid2.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 692px" 
	width=692>
	<PARAM NAME="_ExtentX" VALUE="18309">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="CompletedSplits">
	<PARAM NAME="CtrlName" VALUE="Grid1">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPASplitID","NPASplitComplete","PDPStartDate","PDPEndDate","CurrentNPA","NewNPA","Status"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4,5,6">
	<PARAM NAME="displayWidth" VALUE="97,89,83,89,94,122,120">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPASplitID","NPASplitComplete","PDPStartDate","PDPEndDate","CurrentNPA","NewNPA","Status"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,,,">
	<PARAM NAME="ColumnCount" VALUE="7">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="2">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="10">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="692">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 10;
Grid1.setDataSource(CompletedSplits);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=7 rules=ALL WIDTH=692';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=97';
Grid1.headerWidth[1] = ' WIDTH=89';
Grid1.headerWidth[2] = ' WIDTH=83';
Grid1.headerWidth[3] = ' WIDTH=89';
Grid1.headerWidth[4] = ' WIDTH=94';
Grid1.headerWidth[5] = ' WIDTH=122';
Grid1.headerWidth[6] = ' WIDTH=120';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPASplitID\'';
Grid1.colHeader[1] = '\'NPASplitComplete\'';
Grid1.colHeader[2] = '\'PDPStartDate\'';
Grid1.colHeader[3] = '\'PDPEndDate\'';
Grid1.colHeader[4] = '\'CurrentNPA\'';
Grid1.colHeader[5] = '\'NewNPA\'';
Grid1.colHeader[6] = '\'Status\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=97';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'CompletedSplits.fields.getValue(\'NPASplitID\')';
Grid1.colAttributes[1] = '  WIDTH=89';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'CompletedSplits.fields.getValue(\'NPASplitComplete\')';
Grid1.colAttributes[2] = '  WIDTH=83';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'CompletedSplits.fields.getValue(\'PDPStartDate\')';
Grid1.colAttributes[3] = '  WIDTH=89';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'CompletedSplits.fields.getValue(\'PDPEndDate\')';
Grid1.colAttributes[4] = '  WIDTH=94';
Grid1.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[4] = 'CompletedSplits.fields.getValue(\'CurrentNPA\')';
Grid1.colAttributes[5] = '  WIDTH=122';
Grid1.colFormat[5] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[5] = 'CompletedSplits.fields.getValue(\'NewNPA\')';
Grid1.colAttributes[6] = '  WIDTH=120';
Grid1.colFormat[6] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[6] = 'CompletedSplits.fields.getValue(\'Status\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid1.hasPageNumber = true;
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid3 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 692px" 
	width=692>
	<PARAM NAME="_ExtentX" VALUE="18309">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="CompletedTransfers">
	<PARAM NAME="CtrlName" VALUE="Grid3">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"TransferID","TransferDate","CurrentEntity","NewEntity","NPA","Status"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4,5">
	<PARAM NAME="displayWidth" VALUE="95,112,104,109,119,150">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"TransferID","TransferDate","CurrentEntity","NewEntity","NPA","Status"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,,">
	<PARAM NAME="ColumnCount" VALUE="6">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="2">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="10">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="692">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid3()
{
Grid3.pageSize = 10;
Grid3.setDataSource(CompletedTransfers);
Grid3.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=6 rules=ALL WIDTH=692';
Grid3.headerAttributes = '   bgcolor=Maroon align=Center';
Grid3.headerWidth[0] = ' WIDTH=95';
Grid3.headerWidth[1] = ' WIDTH=112';
Grid3.headerWidth[2] = ' WIDTH=104';
Grid3.headerWidth[3] = ' WIDTH=109';
Grid3.headerWidth[4] = ' WIDTH=119';
Grid3.headerWidth[5] = ' WIDTH=150';
Grid3.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid3.colHeader[0] = '\'TransferID\'';
Grid3.colHeader[1] = '\'TransferDate\'';
Grid3.colHeader[2] = '\'CurrentEntity\'';
Grid3.colHeader[3] = '\'NewEntity\'';
Grid3.colHeader[4] = '\'NPA\'';
Grid3.colHeader[5] = '\'Status\'';
Grid3.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid3.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid3.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid3.colAttributes[0] = '  WIDTH=95';
Grid3.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[0] = 'CompletedTransfers.fields.getValue(\'TransferID\')';
Grid3.colAttributes[1] = '  WIDTH=112';
Grid3.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[1] = 'CompletedTransfers.fields.getValue(\'TransferDate\')';
Grid3.colAttributes[2] = '  WIDTH=104';
Grid3.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[2] = 'CompletedTransfers.fields.getValue(\'CurrentEntity\')';
Grid3.colAttributes[3] = '  WIDTH=109';
Grid3.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[3] = 'CompletedTransfers.fields.getValue(\'NewEntity\')';
Grid3.colAttributes[4] = '  WIDTH=119';
Grid3.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[4] = 'CompletedTransfers.fields.getValue(\'NPA\')';
Grid3.colAttributes[5] = '  WIDTH=150';
Grid3.colFormat[5] = '<Font Size=2 Face="Arial" Color=Black >';
Grid3.colData[5] = 'CompletedTransfers.fields.getValue(\'Status\')';
Grid3.navbarAlignment = 'Right';
var objPageNavbar = Grid3.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid3.hasPageNumber = true;
}
function _Grid3_ctor()
{
	CreateDataGrid('Grid3',_initGrid3);
}
</SCRIPT>

<%	Grid3.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid5 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 691px" 
	width=691>
	<PARAM NAME="_ExtentX" VALUE="18283">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="OldLogs">
	<PARAM NAME="CtrlName" VALUE="Grid5">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPA","NXX","UserLogon","Log_Date","Process","Action","ActionText"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4,5,6">
	<PARAM NAME="displayWidth" VALUE="51,79,87,145,97,100,144">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPA","CO Code","Logon ID","Log Date/Time","Process","Action","ActionText"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,,,">
	<PARAM NAME="ColumnCount" VALUE="7">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="2">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="10">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="691">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid5()
{
Grid5.pageSize = 10;
Grid5.setDataSource(OldLogs);
Grid5.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=7 rules=ALL WIDTH=691';
Grid5.headerAttributes = '   bgcolor=Maroon align=Center';
Grid5.headerWidth[0] = ' WIDTH=51';
Grid5.headerWidth[1] = ' WIDTH=79';
Grid5.headerWidth[2] = ' WIDTH=87';
Grid5.headerWidth[3] = ' WIDTH=145';
Grid5.headerWidth[4] = ' WIDTH=97';
Grid5.headerWidth[5] = ' WIDTH=100';
Grid5.headerWidth[6] = ' WIDTH=144';
Grid5.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid5.colHeader[0] = '\'NPA\'';
Grid5.colHeader[1] = '\'CO Code\'';
Grid5.colHeader[2] = '\'Logon ID\'';
Grid5.colHeader[3] = '\'Log Date/Time\'';
Grid5.colHeader[4] = '\'Process\'';
Grid5.colHeader[5] = '\'Action\'';
Grid5.colHeader[6] = '\'ActionText\'';
Grid5.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid5.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid5.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid5.colAttributes[0] = '  WIDTH=51';
Grid5.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[0] = 'OldLogs.fields.getValue(\'NPA\')';
Grid5.colAttributes[1] = '  WIDTH=79';
Grid5.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[1] = 'OldLogs.fields.getValue(\'NXX\')';
Grid5.colAttributes[2] = '  WIDTH=87';
Grid5.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[2] = 'OldLogs.fields.getValue(\'UserLogon\')';
Grid5.colAttributes[3] = '  WIDTH=145';
Grid5.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[3] = 'OldLogs.fields.getValue(\'Log_Date\')';
Grid5.colAttributes[4] = '  WIDTH=97';
Grid5.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[4] = 'OldLogs.fields.getValue(\'Process\')';
Grid5.colAttributes[5] = '  WIDTH=100';
Grid5.colFormat[5] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[5] = 'OldLogs.fields.getValue(\'Action\')';
Grid5.colAttributes[6] = '  WIDTH=144';
Grid5.colFormat[6] = '<Font Size=2 Face="Arial" Color=Black >';
Grid5.colData[6] = 'OldLogs.fields.getValue(\'ActionText\')';
Grid5.navbarAlignment = 'Right';
var objPageNavbar = Grid5.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid5.hasPageNumber = true;
}
function _Grid5_ctor()
{
	CreateDataGrid('Grid5',_initGrid5);
}
</SCRIPT>

<%	Grid5.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>

</HTML>

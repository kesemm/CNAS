<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta name="VI60_DTCScriptingPlatform" content="Server (ASP)">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>Connection</title>
<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

function checkNull(c) 
	dim temp
	if isnull(c) then 
		temp=""
	else 
		temp=c 	
	end if	
	checkNull=temp
end function


Sub btnUpdate_onclick()

txt=lstStatus.selectedIndex
session("spStatus")=lstStatus.getValue(txt)
	
if session("sParamNPA")<>lstNPA.getText(lstNPA.selectedIndex) or session("sParamNXX")<>txtNXX.value	then
session("Error")="NPANXXChanged"
end if	
	session("COCodeAct")="Update"
	session("COCodeAct1")="Update"
End Sub

sub RestoreScreenValue
	txtOCN.value=session("spOCN")
	txtLATA.value=session("spLATA")
	txtSwitchID.value=session("spSwitchID")
	txtWireCenter.value=session("spWireCenter")
	txtRateCenter.value=session("spRateCenter")
end sub

Sub btnGetRecord_onclick()
	
if lstNPA.getText(lstNPA.selectedIndex)<>"" and txtNXX.value<>"" then
	session("COCodeAct")="GetRec" 
	session("sParamNPA")=lstNPA.getText(lstNPA.selectedIndex)
	session("sParamNXX")=txtNXX.value
	end if	
End Sub

Sub btnClearScreen_onclick()
	
	session("sParamNPA")=""
	session("sParamNXX")=""
	lstNPA.getText(0)
	txtNXX.value=""
	lstStatus.getText(0)
	lstEntity.getText(0)
	txtLATA.value=""
	txtOCN.value=""
	txtSwitchID.value=""
	txtWireCenter.value=""
	txtRateCenter.value=""
	txtTollHoming.value=""
	txtExchangeType.value=""
	txtRemarks.value=""
	txtEAS.value=""
	Response.Redirect "xca_Remarks.asp"
	
End Sub

function GetRecord(ParamNPA,ParamNXX)

	dim objConn
	dim objCmd
	dim objRec

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

	objCmd.CommandText="SelectRemarks '" & trim(ParamNPA) & "', '" & trim(ParamNXX) & "'"
			
	set objRec=objCmd.Execute
	GetRecordTemp=false
	if not objRec.EOF then
	
lstNPA.selectByText(checkNull(objRec("NPA")))		
txtNXX.value=checkNull(objRec("NXX"))
lstStatus.selectByValue(checkNull(objRec("Status")))
session("OldStatus")=checkNull(objRec("Status"))
lstEntity.selectByValue(checkNull(objRec("EntityID")))
session("EntityID")=checkNull(objRec("EntityID"))
if session("EntityID")="" then session("EntityID")="0"
txtLATA.value=trim(checkNull(objRec("LATA")))
txtOCN.value=trim(checkNull(objRec("OCN")))
txtSwitchID.value=trim(checkNull(objRec("SwitchID")))
txtWireCenter.value=trim(checkNull(objRec("WireCenter")))
txtRateCenter.value=trim(checkNull(objRec("RateCenter")))
txtTollHoming.value=trim(checkNull(objRec("TollHoming")))
txtExchangeType.value=trim(checkNull(objRec("ExchangeType")))
txtRemarks.value=trim(checkNull(objRec("Remarks")))
txtEAS.value=trim(checkNull(objRec("EAS")))
GetRecordTemp=true					
end if			
	objRec.close
	objConn.close
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	GetRecord=GetRecordTemp
	
end function

Sub btnReturnToMain_onclick()

	session("sParamNPA")=""
	session("sParamNXX")=""
	session("COCodeAct")=""
	session("Error")=""
	
	Response.Redirect "xca_MenuC0CAdmin.asp"
	
End Sub

</script>
</head>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<%		
txt=lstEntity.getText(0)
if txt<>"" then
lstEntity .addItem "","0",0
end if
lstEntity.selectByValue "0"
		
txt=lstStatus.getText(0)
if txt<>"" then
	lstStatus .addItem "","0",0
end if
lstStatus.selectByValue "0"

txt=lstNPA.getText(0)
if txt<>"" then
	lstNPA .addItem "","0",0
end if
lstNPA.selectByValue "0"
		
select case session("COCodeAct")
	case "GetRec" 
		
	ParamNPA=session("sParamNPA")
	ParamNXX=session("sParamNXX")
	if ParamNPA="" or ParamNXX="" then
		
	elseif (not isnumeric(ParamNPA)) or (not isnumeric(ParamNXX)) then%>
<script Language="JavaScript">
		alert("Both NPA and NXX must be numeric data type.")
		</script>
<% 
	elseif not GetRecord(ParamNPA,ParamNXX)then%>
<script Language="JavaScript">
		alert("No record exists for the specified NPA and NXX.")
		</script>
<% 
		end if
			 	
	case "Update" 
	
	ParamNPA=session("sParamNPA")
	ParamNXX=session("sParamNXX")
		
	%>
<%select case session("Error")%>
<%	case "NPANXXChanged"%>
<script
Language="JavaScript">
		alert("NPA-NXX value has been changed on the screen. Operation cancelled.")
		</script>
<%	case else	
	
	pStatus=session("spStatus")
	pEntityID=session("spEntityID")
	pLATA=session("spLATA")
	pOCN=session("spOCN")
	pSwitchID=session("spSwitchID")
	pWireCenter=session("spWireCenter")
	pRateCenter=session("spRateCenter")
	pEarliestInServiceDate=session("spEarliestInServiceDate")
	pInServiceDate=session("spInServiceDate")
	pTollHoming=session("spTollHoming")
	pExchangeType=session("spExchange")
	pRemarks=session("spRemarks")
	pEAS=session("spEAS")

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	objCmd.CommandText=	"UpdateRemarks'" & trim(ParamNPA) _
	& "', '" & trim(ParamNXX) _
	& "', '" & Replace(trim(txtTollHoming.value),"'","''") _
	& "', '" & Replace(trim(txtExchangeType.value),"'","''") _
	& "', '" & Replace(trim(txtRemarks.value),"'","''") _
	& "', '" & Replace(trim(txtEAS.value),"'","''") _
	& "'"
	objCmd.Execute
	objConn.close
	Set objCmd=Nothing
	Set objConn=Nothing
	Set objRec=Nothing
%>
<script Language="JavaScript">
alert("The record has been updated successfully.")
</script>
<%
end select	
	if session("Error")<>"" then RestoreScreenValue
		session("Error")=""
		call GetRecord(ParamNPA,ParamNXX)
	
	case else 
	end select
	session("COCodeAct")=""
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecNPA 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sdistinct\s\sNPA\sFROM\sxca_COCode\sORDER\sBY\sNPA\q,TCControlID_Unmatched=\qRecNPA\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sdistinct\s\sNPA\sFROM\sxca_COCode\sORDER\sBY\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
</OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecNPA()
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
	cmdTmp.CommandText = 'SELECT distinct  NPA FROM xca_COCode ORDER BY NPA';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecNPA.setRecordSource(rsTmp);
	RecNPA.open();
	if (thisPage.getState('pb_RecNPA') != null)
	RecNPA.setBookmark(thisPage.getState('pb_RecNPA'));
}
function _RecNPA_ctor()
{
	CreateRecordset('RecNPA', _initRecNPA, null);
}
function _RecNPA_dtor()
{
	RecNPA._preserveState();
	thisPage.setState('pb_RecNPA', RecNPA.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecStatus 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sCOStatus,COStatusDescription\sfrom\sxca_status_codes\sorder\sby\sCOStatusDescription\q,TCControlID_Unmatched=\qRecStatus\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sCOStatus,COStatusDescription\sfrom\sxca_status_codes\sorder\sby\sCOStatusDescription\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
												        </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecStatus()
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
	cmdTmp.CommandText = 'Select COStatus,COStatusDescription from xca_status_codes order by COStatusDescription';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecStatus.setRecordSource(rsTmp);
	RecStatus.open();
	if (thisPage.getState('pb_RecStatus') != null)
		RecStatus.setBookmark(thisPage.getState('pb_RecStatus'));
}
function _RecStatus_ctor()
{
	CreateRecordset('RecStatus', _initRecStatus, null);
}
function _RecStatus_dtor()
{
	RecStatus._preserveState();
	thisPage.setState('pb_RecStatus', RecStatus.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecEntity 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sEntityID,EntityName\sfrom\sxca_Entity\sorder\sby\sEntityName\q,TCControlID_Unmatched=\qRecEntity\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sEntityID,EntityName\sfrom\sxca_Entity\sorder\sby\sEntityName\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
																								          </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecEntity()
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
	cmdTmp.CommandText = 'select EntityID,EntityName from xca_Entity order by EntityName';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecEntity.setRecordSource(rsTmp);
	RecEntity.open();
	if (thisPage.getState('pb_RecEntity') != null)
		RecEntity.setBookmark(thisPage.getState('pb_RecEntity'));
}
function _RecEntity_ctor()
{
	CreateRecordset('RecEntity', _initRecEntity, null);
}
function _RecEntity_dtor()
{
	RecEntity._preserveState();
	thisPage.setState('pb_RecEntity', RecEntity.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->


<table WIDTH="75%" ALIGN="center" border="0" CELLSPACING="1" CELLPADDING="1">
  <tr>
    <td ALIGN="middle"><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 
      id=lblTitle style="HEIGHT: 37px; LEFT: 37px; TOP: 0px; WIDTH: 485px" 
      width=485>
	<PARAM NAME="_ExtentX" VALUE="12832">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="View and Update Remarks">
	<PARAM NAME="FontFace" VALUE="Arial Black">
	<PARAM NAME="FontSize" VALUE="5">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">

  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial Black" SIZE="5" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTitle()
{
	lblTitle.setCaption('View and Update Remarks');
}
function _lblTitle_ctor()
{
	CreateLabel('lblTitle', _initlblTitle, null);
}
</script>
<% lblTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
</table>

<p><br>
</p>

<table border="1" cellPadding="2" cellSpacing="2" cols="2" align="center">
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnGetRecord style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 109px" 
      width=109>
	<PARAM NAME="_ExtentX" VALUE="2884">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGetRecord">
	<PARAM NAME="Caption" VALUE="Get NPA-NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">

</OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGetRecord()
{
	btnGetRecord.value = 'Get NPA-NXX';
	btnGetRecord.setStyle(0);
}
function _btnGetRecord_ctor()
{
	CreateButton('btnGetRecord', _initbtnGetRecord, null);
}
</script>
<% btnGetRecord.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnUpdate style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 108px" 
      width=108>
	<PARAM NAME="_ExtentX" VALUE="2858">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="  Update Remarks  ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnUpdate()
{
	btnUpdate.value = '  Update Remarks  ';
	btnUpdate.setStyle(0);
}
function _btnUpdate_ctor()
{
	CreateButton('btnUpdate', _initbtnUpdate, null);
}
</script>
<% btnUpdate.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnClearScreen style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 106px" 
      width=106>
	<PARAM NAME="_ExtentX" VALUE="2805">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnClearScreen">
	<PARAM NAME="Caption" VALUE="Clear Screen">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnClearScreen()
{
	btnClearScreen.value = 'Clear Screen';
	btnClearScreen.setStyle(0);
}
function _btnClearScreen_ctor()
{
	CreateButton('btnClearScreen', _initbtnClearScreen, null);
}
</script>
<% btnClearScreen.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnReturnToMain style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" 
      width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMain">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
</table>

<p><!-- <TABLE border=1 cellPadding=1 cellSpacing=1  cellPadding=2 cellSpacing=2 width=400 style="WIDTH: 400px" 
>--> </p>

<table border="0" cellPadding="2" cellSpacing="2" cols="2" align="center">
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblNPA style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 66px" width=66>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblNPA()
{
	lblNPA.setCaption('NPA');
}
function _lblNPA_ctor()
{
	CreateLabel('lblNPA', _initlblNPA, null);
}
</script>
<% lblNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=lstNPA style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="lstNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecNPA">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstNPA()
{
	RecNPA.advise(RS_ONDATASETCOMPLETE, 'lstNPA.setRowSource(RecNPA, \'NPA\', \'NPA\');');
}
function _lstNPA_ctor()
{
	CreateListbox('lstNPA', _initlstNPA, null);
}
</script>
<% lstNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblNXX style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 23px" width=23>
	<PARAM NAME="_ExtentX" VALUE="609">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblNXX">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NXX">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblNXX()
{
	lblNXX.setCaption('NXX');
}
function _lblNXX_ctor()
{
	CreateLabel('lblNXX', _initlblNXX, null);
}
</script>
<% lblNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtNXX style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNXX()
{
	txtNXX.setStyle(TXT_TEXTBOX);
	txtNXX.setMaxLength(3);
	txtNXX.setColumnCount(3);
}
function _txtNXX_ctor()
{
	CreateTextbox('txtNXX', _inittxtNXX, null);
}
</script>
<% txtNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblStatus style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 35px" 
      width=35>
	<PARAM NAME="_ExtentX" VALUE="926">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblStatus">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Status">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblStatus()
{
	lblStatus.setCaption('Status');
}
function _lblStatus_ctor()
{
	CreateLabel('lblStatus', _initlblStatus, null);
}
</script>
<% lblStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=lstStatus style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1">
	<PARAM NAME="id" VALUE="lstStatus">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecStatus">
	<PARAM NAME="BoundColumn" VALUE="COStatus">
	<PARAM NAME="ListField" VALUE="COStatusDescription">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstStatus()
{
	RecStatus.advise(RS_ONDATASETCOMPLETE, 'lstStatus.setRowSource(RecStatus, \'COStatusDescription\', \'COStatus\');');
lstStatus.disabled=true;
}
function _lstStatus_ctor()
{
	CreateListbox('lstStatus', _initlstStatus, null);
}
</script>
<% lstStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblEntityID style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 46px" 
      width=46>
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblEntityID">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Entity ID">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblEntityID()
{
	lblEntityID.setCaption('Entity ID');
}
function _lblEntityID_ctor()
{
	CreateLabel('lblEntityID', _initlblEntityID, null);
}
</script>
<% lblEntityID.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=lstEntity style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 150px" 
      width=150>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="lstEntity">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecEntity">
	<PARAM NAME="BoundColumn" VALUE="EntityID">
	<PARAM NAME="ListField" VALUE="EntityName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstEntity()
{
	RecEntity.advise(RS_ONDATASETCOMPLETE, 'lstEntity.setRowSource(RecEntity, \'EntityName\', \'EntityID\');');
lstEntity.disabled=true;
}
function _lstEntity_ctor()
{
	CreateListbox('lstEntity', _initlstEntity, null);
}
</script>
<% lstEntity.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblLATA style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 29px" 
width=29>
	<PARAM NAME="_ExtentX" VALUE="767">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblLATA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="LATA">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblLATA()
{
	lblLATA.setCaption('LATA');
}
function _lblLATA_ctor()
{
	CreateLabel('lblLATA', _initlblLATA, null);
}
</script>
<% lblLATA.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtLATA style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 24px" 
width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtLATA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtLATA()
{
	txtLATA.setStyle(TXT_TEXTBOX);
	txtLATA.disabled=true;
	txtLATA.setMaxLength(4);
	txtLATA.setColumnCount(4);
}
function _txtLATA_ctor()
{
	CreateTextbox('txtLATA', _inittxtLATA, null);
}
</script>
<% txtLATA.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblOCN style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 26px" width=26>
	<PARAM NAME="_ExtentX" VALUE="688">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblOCN">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="OCN">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblOCN()
{
	lblOCN.setCaption('OCN');
}
function _lblOCN_ctor()
{
	CreateLabel('lblOCN', _initlblOCN, null);
}
</script>
<% lblOCN.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtOCN style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 24px" width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtOCN">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtOCN()
{
	txtOCN.setStyle(TXT_TEXTBOX);
	txtOCN.disabled=true;
	txtOCN.setMaxLength(4);
	txtOCN.setColumnCount(4);
}
function _txtOCN_ctor()
{
	CreateTextbox('txtOCN', _inittxtOCN, null);
}
</script>
<% txtOCN.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblSource style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 74px" 
      width=74>
	<PARAM NAME="_ExtentX" VALUE="1958">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblSource">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Source SE/POI">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblSource()
{
	lblSource.setCaption('Source SE/POI');
}
function _lblSource_ctor()
{
	CreateLabel('lblSource', _initlblSource, null);
}
</script>
<% lblSource.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtSwitchID style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 66px" 
      width=66>
	<PARAM NAME="_ExtentX" VALUE="1746">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtSwitchID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="11">
	<PARAM NAME="DisplayWidth" VALUE="11">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtSwitchID()
{
	txtSwitchID.setStyle(TXT_TEXTBOX);
	txtSwitchID.disabled=true;
	txtSwitchID.setMaxLength(11);
	txtSwitchID.setColumnCount(11);
}
function _txtSwitchID_ctor()
{
	CreateTextbox('txtSwitchID', _inittxtSwitchID, null);
}
</script>
<% txtSwitchID.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblWireCenter style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 62px" 
      width=62>
	<PARAM NAME="_ExtentX" VALUE="1640">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblWireCenter">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Wire Center">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblWireCenter()
{
	lblWireCenter.setCaption('Wire Center');
}
function _lblWireCenter_ctor()
{
	CreateLabel('lblWireCenter', _initlblWireCenter, null);
}
</script>
<% lblWireCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtWireCenter style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 150px" 
      width=150>
	<PARAM NAME="_ExtentX" VALUE="3969">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtWireCenter">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="40">
	<PARAM NAME="DisplayWidth" VALUE="25">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtWireCenter()
{
	txtWireCenter.setStyle(TXT_TEXTBOX);
	txtWireCenter.disabled=true;
	txtWireCenter.setMaxLength(30);
	txtWireCenter.setColumnCount(25);
}
function _txtWireCenter_ctor()
{
	CreateTextbox('txtWireCenter', _inittxtWireCenter, null);
}
</script>
<% txtWireCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=Label1 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Rate Center">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('Rate Center');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtRateCenter style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 150px" 
      width=150>
	<PARAM NAME="_ExtentX" VALUE="3969">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtRateCenter">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="40">
	<PARAM NAME="DisplayWidth" VALUE="25">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtRateCenter()
{
	txtRateCenter.setStyle(TXT_TEXTBOX);
	txtRateCenter.disabled=true;
	txtRateCenter.setMaxLength(30);
	txtRateCenter.setColumnCount(25);
}
function _txtRateCenter_ctor()
{
	CreateTextbox('txtRateCenter', _inittxtRateCenter, null);
}
</script>
<% txtRateCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblTollHoming style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
      width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblTollHoming">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="TollHoming">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">     
											              
      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTollHoming()
{
	lblTollHoming.setCaption('Toll Homing');
}
function _lblTollHoming_ctor()
{
	CreateLabel('lblTollHoming', _initlblTollHoming, null);
}
</script>
<% lblTollHoming.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtTollHoming style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
      width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtTollHoming">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
													      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtTollHoming()
{
	txtTollHoming.setStyle(TXT_TEXTBOX);
	txtTollHoming.setMaxLength(20);
	txtTollHoming.setColumnCount(20);
}
function _txtTollHoming_ctor()
{
	CreateTextbox('txtTollHoming', _inittxtTollHoming, null);
}
</script>
<% txtTollHoming.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblExchangeType style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
      width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblExchangeType">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="ExchangeType">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">     
											              
      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblExchangeType()
{
	lblExchangeType.setCaption('Exchange Type');
}
function _lblExchangeType_ctor()
{
	CreateLabel('lblExchangeType', _initlblExchangeType, null);
}
</script>
<% lblExchangeType.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtExchangeType style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
      width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtExchangeType">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
													      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtExchangeType()
{
	txtExchangeType.setStyle(TXT_TEXTBOX);
	txtExchangeType.setMaxLength(20);
	txtExchangeType.setColumnCount(20);
}
function _txtExchangeType_ctor()
{
	CreateTextbox('txtExchangeType', _inittxtExchangeType, null);
}
</script>
<% txtExchangeType.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
</table>

<p><br>
</p>

<table border="0" cellPadding="2" cellSpacing="2" cols="1" align="center">
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblRemarks style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
      width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblRemarks">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Remarks">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">     
											              
      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblRemarks()
{
	lblRemarks.setCaption('Remarks');
}
function _lblRemarks_ctor()
{
	CreateLabel('lblRemarks', _initlblRemarks, null);
}
</script>
<% lblRemarks.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtRemarks style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
      width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtRemarks">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
													      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtRemarks()
{
	txtRemarks.setStyle(TXT_TEXTBOX);
	txtRemarks.setMaxLength(80);
	txtRemarks.setColumnCount(80);
}
function _txtRemarks_ctor()
{
	CreateTextbox('txtRemarks', _inittxtRemarks, null);
}
</script>
<% txtRemarks.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
      id=lblEAS style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
      width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblEAS">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="EAS">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">     
											              
      </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblEAS()
{
	lblEAS.setCaption('EAS');
}
function _lblEAS_ctor()
{
	CreateLabel('lblEAS', _initlblEAS, null);
}
</script>
<% lblEAS.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=txtEAS style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
      width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEAS">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEAS()
{
	txtEAS.setStyle(TXT_TEXTBOX);
	txtEAS.setMaxLength(80);
	txtEAS.setColumnCount(80);
}
function _txtEAS_ctor()
{
	CreateTextbox('txtEAS', _inittxtEAS, null);
}
</script>
<% txtEAS.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
</table>

<p>&nbsp;</p>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

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
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Connection</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
</SCRIPT>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>
dim RecIndex


	
Sub btnInclude_onclick()
	Rec1.fields.setValue "TransferID",txtTransferID.value
	log "C",session("pNPA"),"",session("UserUserID"),Now,0,"Edit",txtTransferID.value,"Transfer" 
		
	Rec1.movenext
End Sub


Sub btnExclude_onclick()
	Rec1.fields.setValue "TransferID","Excluded"
	log "C",session("pNPA"),"",session("UserUserID"),Now,0,"Edit",txtTransferID.value,"Transfer" 
	
	Rec1.movenext
End Sub


sub getSessionParam()
	session("pTransferID")=txtTransferID.value	
	'session("pNPA")=txtNPA.value	
	txt=LSTNPA.selectedIndex
	session("pNPA")=LSTNPA.getText(txt)
	
	'session("pEntityID")=txtCurrentEntity.value
	txt=LSTCurrentEntity.selectedIndex
	session("pEntityID")=LSTCurrentEntity.getValue(txt)
	
	'session("pNewEntity")=txtNewEntity.value
	txt=LSTNewEntity.selectedIndex
	session("pNewEntity")=LSTNewEntity.getValue(txt)
		
	session("pTransferDate")=txtTransferDate.value
end sub


Sub btnAddNew_onclick()
	getSessionParam
	session("TransferAct")="Add"
	Response.Redirect "xca_Transfer.asp"
End Sub


Sub btnUpdate_onclick()
	getSessionParam
	session("TransferAct")="Update"
	Response.Redirect "xca_Transfer.asp"
End Sub


Sub btnComplete_onclick()
	getSessionParam
	session("TransferAct")="Transfer"
	Response.Redirect "xca_Transfer.asp"
End Sub


Sub btnDelete_onclick()
	session("pTransferID")=txtTransferID.value	
	session("pNPA")=""
	txt=LSTNPA.selectedIndex
	session("pNPA_log")=LSTNPA.getText(txt)
	
	session("pEntityID")=""
	session("pNewEntity")=""
	session("pTransferDate")=""
	session("TransferAct")="Delete"
	Response.Redirect "xca_Transfer.asp"
End Sub


sub ResetScreen()
	
	txtTransferID.value = ""
	
	LSTNPA.selectByText("")
	NPA1=""

	LSTCurrentEntity.selectByValue("")
	CurrentEntity=""

	LSTNewEntity.selectByValue("")
	NewEntity=""

	txtTransferDate.value=""
	
	session("pTransferID")=""
	session("pNPA")=""
	session("pNewEntity")=""
	session("pTransferDate")=""
	
end sub


Sub btnReturn_onclick()
	Response.Redirect "xca_TransferAdmin.asp"
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor=#d7c7a4>

<%	
txtTransferID.value = session("pTransferID")

'Response.Write session("pEntityID") & "|||"
'txtNPA.value=session("pNPA")
LSTNPA.selectByText(session("pNPA"))
NPA1=trim(session("pNPA"))

'txtCurrentEntity.value=trim(session("pEntityID"))
LSTCurrentEntity.selectByValue(session("pEntityID"))
CurrentEntity=trim(session("pEntityID"))


'txtNewEntity.value=trim(session("pNewEntity"))
LSTNewEntity.selectByValue(session("pNewEntity"))
NewEntity=trim(session("pNewEntity"))

txtTransferDate.value=trim(session("pTransferDate"))

dim objConn
dim objCmd
dim objRec

Select Case session("TransferAct")
Case "Add"
	if txtTransferID.value="" then
		session("Error")="Missing Transfer ID."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Transfer ID.")
		</SCRIPT>
		<%
	elseif NPA1="" then
		session("Error")="Missing NPA."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing NPA.")
		</SCRIPT>
		<%	
	elseif not isnumeric(NPA1) then
		session("Error")="NPA must be numeric data type."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("NPA must be numeric data type.")
		</SCRIPT>
		<%	
	elseif CurrentEntity="" then
		session("Error")="Missing Current Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Current Entity.")
		</SCRIPT>
		<%			
	elseif NewEntity="" then
		session("Error")="Missing New Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing New Entity.")
		</SCRIPT>
		<%	
	elseif CurrentEntity=NewEntity then
		session("Error")="Same Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Entities can not be the same.")
		</SCRIPT>
		<%				
	elseif txtTransferDate.value="" then
		session("Error")="Missing Transfer Date."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Transfer Date.")
		</SCRIPT>
		<%					
	elseif (not IsDateReal(txtTransferDate.value))  then
		session("Error")="Wrong Format in date field(s)."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Incorrect format in date field.")
		</SCRIPT>
		<%
		
	else
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set objConn=server.CreateObject("ADODB.Connection")
		Set objRec=server.CreateObject("ADODB.Recordset")
		Set objCmd=server.CreateObject("ADODB.Command")
	
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
	
		on error resume next
		objCmd.CommandText="CheckExistingTransfer '" & Replace(trim(txtTransferID.value),"'","''") & "'"
		set objRec=objCmd.Execute%>
	
		<%if not objRec.EOF then  %>
			<SCRIPT Language="VBSCRIPT">
			alert("The transfer ID is already in the data base. This record can not be added.")
			</SCRIPT>
		<%else
			
			objCmd.CommandText=	"AddTransfer '" & Replace(trim(txtTransferID.value),"'","''") _
												& "', '" & trim(txtTransferDate.value) _
												& "', " & trim(CurrentEntity) _
												& ", " & trim(NewEntity) _
												& ", '" & trim(NPA1) _
												& "', '" & "P" & "'"
			objCmd.Execute %>
							
			<%if objConn.Errors.Count <> 0 then  %>
				<SCRIPT Language="VBSCRIPT">
				alert("An error has occured while adding the transfer.")
				</SCRIPT>
			<%else
				log "C",trim(NPA1),"",session("UserUserID"),Now,0,"New",trim(txtTransferID.value),"Transfer" 
				'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
				'email session("AdminEntityEMail"),EMailTo,"","Transfer Added", "Transfer < " & trim(txtTransferID.value) & " > added on " & date 
			end if%>
				
		<% end if%>
	
	
	
		<%objConn.close
		Set objConn=Nothing
		Set objRec=Nothing
		Set objCmd=Nothing
		'session("TransferAct")=""
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	end if
	session("Error")=""
	session("TransferAct")=""
	
case "Update"
	if txtTransferID.value="" then
		session("Error")="Missing Transfer ID."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Transfer ID.")
		</SCRIPT>
		<%
	elseif NPA1="" then
		session("Error")="Missing NPA."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing NPA.")
		</SCRIPT>
		<%	
	elseif not isnumeric(NPA1) then
		session("Error")="NPA must be numeric data type."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("NPA must be numeric data type.")
		</SCRIPT>
		<%	
	elseif CurrentEntity="" then
		session("Error")="Missing Current Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Current Entity.")
		</SCRIPT>
		<%			
	elseif NewEntity="" then
		session("Error")="Missing New Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing New Entity.")
		</SCRIPT>
		<%	
	elseif CurrentEntity=NewEntity then
		session("Error")="Same Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Entities can not be the same.")
		</SCRIPT>
		<%		
	elseif txtTransferDate.value="" then
		session("Error")="Missing Transfer Date."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Missing Transfer Date.")
		</SCRIPT>
		<%					
	elseif (not IsDateReal(txtTransferDate.value))  then
		session("Error")="Wrong Format in date field(s)."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Incorrect format in date field.")
		</SCRIPT>
		<%
		
	else 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set objConn=server.CreateObject("ADODB.Connection")
		Set objRec=server.CreateObject("ADODB.Recordset")
		Set objCmd=server.CreateObject("ADODB.Command")
	
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
	
		on error resume next
		objCmd.CommandText="CheckExistingTransfer '" & Replace(trim(txtTransferID.value),"'","''") & "'"
		set objRec = objCmd.Execute%>
	
		<%if objRec.EOF then  %>
			<SCRIPT Language="VBSCRIPT">
			alert("The transfer ID does not exist in the data base. No record is updated.")
			</SCRIPT>
		<%else
		
			objCmd.CommandText=	"UpdateTransfer '" & Replace(trim(txtTransferID.value),"'","''") & "', '" & trim(txtTransferDate.value) _
											& "', " & trim(CurrentEntity)& ", " & trim(NewEntity) _
											& ", '" & trim(NPA1)& "', '" & "P" & "'"
			objCmd.Execute%>
			
			<%if objConn.Errors.Count <> 0 then  %>
				<SCRIPT Language="VBSCRIPT">
				alert("An error has occured while updating the transfer.")
				</SCRIPT>
			<%else
				log "C",trim(NPA1),"",session("UserUserID"),Now,0,"Edit",trim(txtTransferID.value),"Transfer" 
				'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
				'email session("AdminEntityEMail"),EMailTo,"","Transfer Updated", "Transfer < " & trim(txtTransferID.value) & " > updated on " & date 
			end if%>
			
		<%end if%>
	
		<%objConn.close
		Set objConn=Nothing
		Set objRec=Nothing
		Set objCmd=Nothing
	'	session("TransferAct")=""
	'''''''''''''''''''''''''''''''''''''''''''
	end if
	session("Error")=""
	session("TransferAct")=""
case "Transfer"	
	session("TransferAct")=""
	if CurrentEntity=NewEntity then
		session("Error")="Same Entity."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Entities can not be the same.")
		</SCRIPT>
		<%	
	elseif cdate(txtTransferDate.value)>cdate(date()) then 
	'Response.Write cdate(txtTransferDate.value) & "||" & cdate(date())
		session("Error")="Wrong Format in date field(s)."
		%>
		<SCRIPT Language="VBSCRIPT">
		alert("Transfer can not be completed before the specified transfer date.")
		</SCRIPT>
		<%
	else		
		Set objConn=server.CreateObject("ADODB.Connection")
		Set objCmd=server.CreateObject("ADODB.Command")
	
		on error resume next
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
		objCmd.CommandText=	"Transfer " & trim(NewEntity) & ", '" & Replace(trim(txtTransferID.value),"'","''")& "'"
		objCmd.Execute%>
	
		<%if objConn.Errors.Count <> 0 then  %>
			<SCRIPT Language="VBSCRIPT">
			alert("An error has occured.")
			</SCRIPT>
		<%else%>
			<%if objConn.Errors.Count <> 0 then  %>
				<SCRIPT Language="VBSCRIPT">
				alert("An error has occured while executing the transfer.")
				</SCRIPT>
			<%else
				log "C",trim(NPA1),"",session("UserUserID"),Now,0,"Complete",trim(txtTransferID.value),"Transfer" 
				ResetScreen
				'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
				'email session("AdminEntityEMail"),EMailTo,"","Transfer Completed", "Transfer < " & trim(txtTransferID.value) & " > completed on " & date 
			end if%>
		<%end if%>
	
		<%objConn.close
		Set objConn=Nothing
		Set objCmd=Nothing
		session("TransferAct")=""
	end if
	
case "Delete"	

	txtTransferID.value=""
	on error resume next
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objCmd=server.CreateObject("ADODB.Command")
	
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
		
	objCmd.CommandText=	"DeleteTransfer '" & Replace(trim(session("pTransferID")),"'","''") & "'"
	objCmd.Execute%>
		
	<%if objConn.Errors.Count <> 0 then  %>
		<SCRIPT Language="VBSCRIPT">
		alert("An error has occured while deleting the transfer.")
		</SCRIPT>
	<%else
		log "C",trim(session("pNPA_log")),"",session("UserUserID"),Now,0,"Delete",trim(session("pTransferID")),"Transfer" 
		'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
		'email session("AdminEntityEMail"),EMailTo,"","Transfer Deleted", "Transfer < " & trim(session("pTransferID")) & " > deleted on " & date 
	end if%>
	
	<%objConn.close
	Set objConn=Nothing
	Set objCmd=Nothing
	session("TransferAct")=""	
	session("pTransferID")=""
	session("pNPA_log")=""
end select	

dim NPA
dim TransferID
dim EntityID

TransferID = txtTransferID.value
txt=LSTCurrentEntity.selectedIndex
EntityID=LSTCurrentEntity.getValue(txt)
if EntityID="" then EntityID=0
NPA=NPA1
'Response.Write  EntityID & "||" & NPA
Rec1.open

%>


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_COCode\s\sWHERE\sNPA=?\sand\sEntityID=?\sand\s(TransferID=?\sor\sTransferID='Excluded')\sand\sStatus='I'\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_COCode\s\sWHERE\sNPA=?\sand\sEntityID=?\sand\s(TransferID=?\sor\sTransferID='Excluded')\sand\sStatus='I'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=3,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1,CValue_Unmatched=\qNPA\q),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q8\q,CReq=0,CValue_Unmatched=\qEntityID\q),Row3=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam3\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q10\q,CReq=0,CValue_Unmatched=\qTransferID\q)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersRec1()
{
	Rec1.setParameter(0,NPA);
	Rec1.setParameter(1,EntityID);
	Rec1.setParameter(2,TransferID);
}
function _initRec1()
{
	Rec1.advise(RS_ONBEFOREOPEN, _setParametersRec1);
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
	cmdTmp.CommandText = 'SELECT * FROM xca_COCode  WHERE NPA=? and EntityID=? and (TransferID=? or TransferID=\'Excluded\') and Status=\'I\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Rec1.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Rec1') != null)
		Rec1.setBookmark(thisPage.getState('pb_Rec1'));
}
function _Rec1_ctor()
{
	CreateRecordset('Rec1', _initRec1, null);
}
function _Rec1_dtor()
{
	Rec1._preserveState();
	thisPage.setState('pb_Rec1', Rec1.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecEntity style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sdistinct\sEntityID,\sEntityName\s\r\nfrom\sxca_Entity\q,TCControlID_Unmatched=\qRecEntity\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sdistinct\sEntityID,\sEntityName\s\r\nfrom\sxca_Entity\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
	cmdTmp.CommandText = 'select distinct EntityID, EntityName  from xca_Entity';
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecEntityNew style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sdistinct\sEntityName\s\r\nfrom\sxca_Entity\s\s\swhere\sEntityStatus='a'\q,TCControlID_Unmatched=\qRecEntityNew\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sdistinct\sEntityID,\sEntityName\s\r\nfrom\sxca_Entity\s\s\swhere\sEntityStatus='a'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecEntityNew()
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
	cmdTmp.CommandText = 'select distinct EntityID, EntityName  from xca_Entity   where EntityStatus=\'a\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecEntityNew.setRecordSource(rsTmp);
	RecEntityNew.open();
	if (thisPage.getState('pb_RecEntityNew') != null)
		RecEntityNew.setBookmark(thisPage.getState('pb_RecEntityNew'));
}
function _RecEntityNew_ctor()
{
	CreateRecordset('RecEntityNew', _initRecEntityNew, null);
}
function _RecEntityNew_dtor()
{
	RecEntityNew._preserveState();
	thisPage.setState('pb_RecEntityNew', RecEntityNew.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecNPA style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\q,TCControlID_Unmatched=\qRecNPA\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
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
	cmdTmp.CommandText = 'Select Distinct NPA from xca_COCode';
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

<TABLE width=65% ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 id=lblTitle 
	style="HEIGHT: 37px; LEFT: 10px; TOP: 350px; WIDTH: 253px" width=253>
	<PARAM NAME="_ExtentX" VALUE="6694">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="CO Code Transfer">
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
	lblTitle.setCaption('CO Code Transfer');
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


<TABLE width=65% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD  align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblTransferID 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 59px" width=59>
	<PARAM NAME="_ExtentX" VALUE="1561">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblTransferID">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Transfer ID">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTransferID()
{
	lblTransferID.setCaption('Transfer ID');
}
function _lblTransferID_ctor()
{
	CreateLabel('lblTransferID', _initlblTransferID, null);
}
</script>
<% lblTransferID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtTransferID 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtTransferID">
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
function _inittxtTransferID()
{
	txtTransferID.setStyle(TXT_TEXTBOX);
	txtTransferID.setMaxLength(10);
	txtTransferID.setColumnCount(10);
}
function _txtTransferID_ctor()
{
	CreateTextbox('txtTransferID', _inittxtTransferID, null);
}
</script>
<% txtTransferID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblNPA style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 24px" 
	width=24>
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
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=LSTNPA style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
	width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="LSTNPA">
	<PARAM NAME="DataSource" VALUE="RecNPA">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecNPA">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLSTNPA()
{
	RecNPA.advise(RS_ONDATASETCOMPLETE, 'LSTNPA.setRowSource(RecNPA, \'NPA\', \'NPA\');');
	LSTNPA.setDataSource(RecNPA);
}
function _LSTNPA_ctor()
{
	CreateListbox('LSTNPA', _initLSTNPA, null);
}
</script>
<% LSTNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
	<TR>
		<TD align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblCurrentEntity 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblCurrentEntity">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Current Entity">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblCurrentEntity()
{
	lblCurrentEntity.setCaption('Current Entity');
}
function _lblCurrentEntity_ctor()
{
	CreateLabel('lblCurrentEntity', _initlblCurrentEntity, null);
}
</script>
<% lblCurrentEntity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=LSTCurrentEntity 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="LSTCurrentEntity">
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
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLSTCurrentEntity()
{
	RecEntity.advise(RS_ONDATASETCOMPLETE, 'LSTCurrentEntity.setRowSource(RecEntity, \'EntityName\', \'EntityID\');');
}
function _LSTCurrentEntity_ctor()
{
	CreateListbox('LSTCurrentEntity', _initLSTCurrentEntity, null);
}
</script>
<% LSTCurrentEntity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblNewEntity 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 56px" width=56>
	<PARAM NAME="_ExtentX" VALUE="1482">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblNewEntity">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="New Entity">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblNewEntity()
{
	lblNewEntity.setCaption('New Entity');
}
function _lblNewEntity_ctor()
{
	CreateLabel('lblNewEntity', _initlblNewEntity, null);
}
</script>
<% lblNewEntity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=LSTNewEntity 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="LSTNewEntity">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecEntityNew">
	<PARAM NAME="BoundColumn" VALUE="EntityID">
	<PARAM NAME="ListField" VALUE="EntityName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLSTNewEntity()
{
	RecEntityNew.advise(RS_ONDATASETCOMPLETE, 'LSTNewEntity.setRowSource(RecEntityNew, \'EntityName\', \'EntityID\');');
}
function _LSTNewEntity_ctor()
{
	CreateListbox('LSTNewEntity', _initLSTNewEntity, null);
}
</script>
<% LSTNewEntity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
	<TR>
		<TD align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblTransferDate 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 71px" width=71>
	<PARAM NAME="_ExtentX" VALUE="1879">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblTransferDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Transfer Date">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTransferDate()
{
	lblTransferDate.setCaption('Transfer Date');
}
function _lblTransferDate_ctor()
{
	CreateLabel('lblTransferDate', _initlblTransferDate, null);
}
</script>
<% lblTransferDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtTransferDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtTransferDate">
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
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtTransferDate()
{
	txtTransferDate.setStyle(TXT_TEXTBOX);
	txtTransferDate.setMaxLength(10);
	txtTransferDate.setColumnCount(10);
}
function _txtTransferDate_ctor()
{
	CreateTextbox('txtTransferDate', _inittxtTransferDate, null);
}
</script>
<% txtTransferDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=right>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnAddNew 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 68px" width=68>
	<PARAM NAME="_ExtentX" VALUE="1799">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAddNew">
	<PARAM NAME="Caption" VALUE="   Add   ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnAddNew()
{
	btnAddNew.value = '   Add   ';
	btnAddNew.setStyle(0);
}
function _btnAddNew_ctor()
{
	CreateButton('btnAddNew', _initbtnAddNew, null);
}
</script>
<% btnAddNew.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD align=left>&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" id=btnUpdate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="1773">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnUpdate()
{
	btnUpdate.value = 'Update';
	btnUpdate.setStyle(0);
}
function _btnUpdate_ctor()
{
	CreateButton('btnUpdate', _initbtnUpdate, null);
}
</script>
<% btnUpdate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>



<BR>

<TABLE ALIGN=center border=1 cellPadding=1 cellSpacing=1>
	<TR>
		<TD>
       
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnComplete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 152px" width=152>
	<PARAM NAME="_ExtentX" VALUE="4022">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnComplete">
	<PARAM NAME="Caption" VALUE=" Complete Transfer ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnComplete()
{
	btnComplete.value = ' Complete Transfer ';
	btnComplete.setStyle(0);
}
function _btnComplete_ctor()
{
	CreateButton('btnComplete', _initbtnComplete, null);
}
</script>
<% btnComplete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnDelete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 164px" width=164>
	<PARAM NAME="_ExtentX" VALUE="4339">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE="     Delete Transfer     ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnDelete()
{
	btnDelete.value = '     Delete Transfer     ';
	btnDelete.setStyle(0);
}
function _btnDelete_ctor()
{
	CreateButton('btnDelete', _initbtnDelete, null);
}
</script>
<% btnDelete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturn 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturn">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturn()
{
	btnReturn.value = 'Return';
	btnReturn.setStyle(0);
}
function _btnReturn_ctor()
{
	CreateButton('btnReturn', _initbtnReturn, null);
}
</script>
<% btnReturn.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>
<BR>

<TABLE align=center border=1 cellPadding=1 cellSpacing=1>
    
    <TR>
       
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnInclude 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 152px" width=152>
	<PARAM NAME="_ExtentX" VALUE="4022">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnInclude">
	<PARAM NAME="Caption" VALUE="Include Current NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnInclude()
{
	btnInclude.value = 'Include Current NXX';
	btnInclude.setStyle(0);
}
function _btnInclude_ctor()
{
	CreateButton('btnInclude', _initbtnInclude, null);
}
</script>
<% btnInclude.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnExclude 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 157px" width=157>
	<PARAM NAME="_ExtentX" VALUE="4154">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnExclude">
	<PARAM NAME="Caption" VALUE="Exclude Current NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnExclude()
{
	btnExclude.value = 'Exclude Current NXX';
	btnExclude.setStyle(0);
}
function _btnExclude_ctor()
{
	CreateButton('btnExclude', _initbtnExclude, null);
}
</script>
<% btnExclude.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR></TABLE>

<BR><TABLE ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 797px" 
	width=797>
	<PARAM NAME="_ExtentX" VALUE="21087">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="Rec1">
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
	<PARAM NAME="EnableRowNav" VALUE="-1">
	<PARAM NAME="HiliteColor" VALUE="LimeGreen">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NXX","Status","TransferID"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="123,178,200">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NXX","Status","Include"'>
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
	<PARAM NAME="PageSize" VALUE="5">
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
	<PARAM NAME="GridWidth" VALUE="797">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 5;
Grid1.setDataSource(Rec1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=797';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=123';
Grid1.headerWidth[1] = ' WIDTH=178';
Grid1.headerWidth[2] = ' WIDTH=200';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NXX\'';
Grid1.colHeader[1] = '\'Status\'';
Grid1.colHeader[2] = '\'Include\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=123';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'NXX\')';
Grid1.colAttributes[1] = '  WIDTH=178';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'Status\')';
Grid1.colAttributes[2] = '  WIDTH=200';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'TransferID\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid1.hasPageNumber = true;
Grid1.hiliteAttributes = ' bgcolor=LimeGreen';
var objRecNavbar = Grid1.showRecordNavbar(40,1);
objRecNavbar.getButton(1).value = '   <   ';
objRecNavbar.getButton(2).value = '   >   ';
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
	
</TABLE>

<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

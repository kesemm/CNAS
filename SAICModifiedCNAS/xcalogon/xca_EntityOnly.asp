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
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

dim RecIndex
				

function AssociationValid(EntityID)

	dim AssociationValidTemp
	
	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objRec1=server.CreateObject("ADODB.Recordset")
	Set objCmd1=server.CreateObject("ADODB.Command")
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	on error resume next
	objCmd1.CommandText=	"Select UserID from xca_User where UserType='a' and EntityID =  " & EntityID
	set objRec1=objCmd1.Execute
	if objRec1.EOF then
		AssociationValidTemp=true
	else
		AssociationValidTemp=false
	end if
	objConn1.close
	Set objConn1=Nothing
	Set objRec1=Nothing
	Set objCmd1=Nothing
	
	AssociationValid=AssociationValidTemp
	
end function

function DeletionValid(EntityID)

	dim DeletionValidTemp
	
	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objRec1=server.CreateObject("ADODB.Recordset")
	Set objCmd1=server.CreateObject("ADODB.Command")
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	on error resume next
	objCmd1.CommandText=	"Select UserID from xca_User where EntityID =  " & EntityID
	set objRec1=objCmd1.Execute
	if objRec1.EOF then
		DeletionValidTemp=true
	else
		DeletionValidTemp=false
	end if
	objConn1.close
	Set objConn1=Nothing
	Set objRec1=Nothing
	Set objCmd1=Nothing
	
	DeletionValid=DeletionValidTemp
	
end function

function AdminName()

	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objRec1=server.CreateObject("ADODB.Recordset")
	Set objCmd1=server.CreateObject("ADODB.Command")
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	on error resume next
	objCmd1.CommandText=	"Select Value from xca_Parms where Name='ADMIN'"
	set objRec1=objCmd1.Execute
	if not objRec1.EOF then
		AdminNameTemp=objRec1("Value")
	end if
	objConn1.close
	Set objConn1=Nothing
	Set objRec1=Nothing
	Set objCmd1=Nothing
	
	AdminName=AdminNameTemp
	
end function

sub resetSessionValue()
	
	session("EntityName")=""
	session("EntityContact")=""
	session("EntityAddress")=""
	session("EntityCity")=""
	
	session("EntityProvince")=""
	
	session("EntityPostalCode")=""
	
	session("EntityType")=""
		
	session("EntityStatus")=""
		
	session("EntityEMail")=""
	session("EntityFax")=""
	session("EntityTelephone")=""
	session("EntityExtension")=""
	session("EntityID")=""
	
end sub

sub setSessionValue()

	EntityName.value=session("EntityName")
	EntityContact.value=session("EntityContact")
	EntityAddress.value=session("EntityAddress")
	EntityCity.value=session("EntityCity")
	
	EntityProvince.SelectByValue(session("EntityProvince"))
	
	EntityPostalCode.value=session("EntityPostalCode")
	
	EntityType.SelectByValue(session("EntityType"))
			
	EntityStatus.SelectByValue(session("EntityStatus"))
			
	EntityEMail.value=session("EntityEMail")
	EntityFax.value=session("EntityFax")
	EntityTelephone.value=session("EntityTelephone")
	EntityExtension.value=session("EntityExtension")
	EntityID.value=session("EntityID")

end sub

sub getSessionValue()

	session("EntityName")=EntityName.value
	session("EntityContact")=EntityContact.value
	session("EntityAddress")=EntityAddress.value
	session("EntityCity")=EntityCity.value
	
	txt=EntityProvince.selectedIndex
	session("EntityProvince")=EntityProvince.GetValue(txt)
	
	session("EntityPostalCode")=EntityPostalCode.value
	
	txt=EntityType.selectedIndex
	session("EntityType")=EntityType.GetValue(txt)
		
	txt=EntityStatus.selectedIndex
	session("EntityStatus")=EntityStatus.GetValue(txt)
		
	session("EntityEMail")=EntityEMail.value
	session("EntityFax")=EntityFax.value
	session("EntityTelephone")=EntityTelephone.value
	session("EntityExtension")=EntityExtension.value
	session("EntityID")=EntityID.value
	
end sub

Sub btnUpdate_onclick()
	
	session("EntityAct")="Update"
	getSessionValue
	
End Sub

Sub btnAdd_onclick()
	
	session("EntityAct")="Add"
	getSessionValue
	
End Sub

Sub btnDelete_onclick()
	
	session("EntityAct")="Delete"
	getSessionValue
							
End Sub

Sub btnGetCurrent_onclick()

	session("originalName")=trim(RecEntityAll.fields.getvalue("EntityName"))
	EntityName.value=trim(session("originalName"))
	EntityContact.value=trim(RecEntityAll.fields.getvalue("EntityContact"))
	EntityAddress.value=trim(RecEntityAll.fields.getvalue("EntityAddress"))
	EntityCity.value=trim(RecEntityAll.fields.getvalue("EntityCity"))
	EntityProvince.selectByvalue(RecEntityAll.fields.getvalue("EntityProvince"))
	EntityPostalCode.value=trim(RecEntityAll.fields.getvalue("EntityPostalCode"))
	EntityType.selectByText(RecEntityAll.fields.getvalue("TypeName"))
	EntityStatus.selectByText(RecEntityAll.fields.getvalue("StatusName"))
	EntityEMail.value=trim(RecEntityAll.fields.getvalue("EntityEmail"))
	EntityFax.value=trim(RecEntityAll.fields.getvalue("EntityFax"))
	EntityTelephone.value=trim(RecEntityAll.fields.getvalue("EntityTelephone"))
	EntityExtension.value=trim(RecEntityAll.fields.getvalue("EntityExtension"))
	EntityID.value=trim(RecEntityAll.fields.getvalue("EntityID"))
	
End Sub

Sub btnReturnToMain_onclick()

	Response.Redirect "xca_MenuSecurityAdmin.asp"
	
End Sub

Sub btnClearSceen_onclick()
	resetSessionValue
	setSessionValue
	EntityProvince.SelectByValue(EntityProvince.getValue(0))
	EntityType.SelectByValue(EntityType.getValue(0))
	EntityStatus.SelectByValue(EntityStatus.getValue(0))
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor="#d7c7a4">
<%

	'setSessionValue
	Select Case session("EntityAct")		
		Case "Add"
			session("EntityAct")=""
			Set objConn=server.CreateObject("ADODB.Connection")
			Set objRec=server.CreateObject("ADODB.Recordset")
			Set objCmd=server.CreateObject("ADODB.Command")
	
			objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
			objCmd.ActiveConnection = objConn
	
			on error resume next
			objCmd.CommandText="Select EntityName from xca_Entity where EntityName = '" & Replace(trim(session("EntityName")),"'","''") & "'"
			set objRec=objCmd.Execute%>

<%if not objRec.EOF then  %>
				<SCRIPT Language="JavaScript">
				alert("The entity is already in the data base. It can not be added again.")
				</SCRIPT>

<%else
				if session("EntityName")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("Entity name must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityContact")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("Entity contact must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityAddress")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("Entity address must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityCity")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("Entity city be entered to add a record.")
					</SCRIPT>

<%
				
				elseif session("EntityProvince")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("Entity province must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityPostalCode")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("Entity postal code must be entered to add a record.")
					</SCRIPT>

<%				elseif session("EntityType")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("Entity type must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityStatus")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("Entity status must be entered to add a record.")
					</SCRIPT>

<%				elseif session("EntityTelephone")="" then
					%>
					<SCRIPT Language="JavaScript">
					alert("Entity telephone must be entered to add a record.")
					</SCRIPT>

<%
				elseif session("EntityEMail")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User e-mail must be entered to add a record.")
					</SCRIPT>

<%
				else
'Added to solve hot fix problem Nov 3 2003
objConn.close
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn

					objCmd.CommandText=	"AddEntity '" & Replace(session("EntityType"),"'","''") _
													& "', '" & Replace(session("EntityStatus"),"'","''") _
													& "', '"& Replace(session("EntityName"),"'","''") _ 
													& "', '"& Replace(session("EntityContact"),"'","''") _
													& "', '"& Replace(session("EntityAddress"),"'","''") _
													& "', '"& Replace(session("EntityCity"),"'","''") _
													& "', '"& Replace(session("EntityProvince"),"'","''") _
													& "', '"& Replace(session("EntityPostalCode"),"'","''") _
													& "', '"& Replace(session("EntityEMail"),"'","''") _
													& "', '"& Replace(session("EntityFax"),"'","''") _
													& "', '"& Replace(session("EntityTelephone"),"'","''") _
													& "', '"& Replace(session("EntityExtension"),"'","''") _
													& "'"
					objCmd.Execute %>

<%if objConn.Errors.Count <> 0 then  %>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while adding the user.")
						</SCRIPT>

<%else
						log "C","","",session("UserUserID"),Now,0,"Add","","Entity" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","Entity Added", "Entity  < " & trim(session("EntityName")) & " > added on " & date 
						ResetSessionValue
						setSessionValue
					%>
						<SCRIPT Language="JavaScript">
						alert("The entity record has been added successfully.")
						</SCRIPT>
						
<%	
					end if%>

<%end if%>

<% end if%>

<%objConn.close
			
			Set objConn=Nothing
			Set objRec=Nothing
			Set objCmd=Nothing
			
			RecEntityAll.requery
			setSessionValue
		case "Update"	
		
			session("EntityAct")=""
			
			if session("EntityID")="" then
			%>
					<SCRIPT Language="JavaScript">
					alert("To update a record, you must select an entity.")
					</SCRIPT>

<%
			else	
				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")
		
				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
				on error resume next
				
					if session("EntityName")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("Entity name must be entered to update a record.")
						</SCRIPT>

<%
					elseif Trim(session("originalName"))=trim(AdminName()) and (trim(session("EntityName"))<>Trim(session("originalName")) or trim(session("EntityType"))<>"a" or trim(session("EntityStatus"))<>"a") then
					%>	<SCRIPT Language="JavaScript">
						alert("Administrator's name, type or status can not be changed.")
						</SCRIPT>

<%
					elseif session("EntityContact")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("Entity contact must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityAddress")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("Entity address must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityCity")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity city must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityProvince")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity province must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityPostalCode")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity postal code must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityType")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity type must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityType")<>"a" and (not AssociationValid(session("EntityID"))) then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity type can not be changed to User because there are administrative users associated with it.")
						</SCRIPT>

<%
					elseif session("EntityStatus")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("User status must be entered to update a record.")
						</SCRIPT>

<%
					elseif session("EntityTelephone")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity telephone must be entered to update a record.")
						</SCRIPT>

<%				
					elseif session("EntityEMail")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("Entity e-mail must be entered to update a record.")
						</SCRIPT>

<%
					else
'Added to solve hot fix problem Nov 3 2003
objConn.close
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn

						objCmd.CommandText=	"UpdateEntity '" & Replace(session("EntityType"),"'","''") _
														& "', '" & Replace(session("EntityStatus"),"'","''") _
														& "', '"& Replace(session("EntityName"),"'","''") _ 
														& "', '"& Replace(session("EntityContact"),"'","''") _
														& "', '"& Replace(session("EntityAddress"),"'","''") _
														& "', '"& Replace(session("EntityCity"),"'","''") _
														& "', '"& Replace(session("EntityProvince"),"'","''") _
														& "', '"& Replace(session("EntityPostalCode"),"'","''") _
														& "', '"& Replace(session("EntityEMail"),"'","''") _
														& "', '"& Replace(session("EntityFax"),"'","''") _
														& "', '"& Replace(session("EntityTelephone"),"'","''") _
														& "', '"& Replace(session("EntityExtension"),"'","''") _
														& "', '"& session("EntityID") _
														& "'"
						objCmd.Execute %>

<%if objConn.Errors.Count <> 0 then  %>
							<SCRIPT Language="JavaScript">
							alert("An error has occured while updating the entity.")
							</SCRIPT>

<%else
							log "C","","",session("UserUserID"),Now,0,"Update","","Entity" 
							'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
							'email session("AdminEntityEMail"),EMailTo,"","Entity Updated", "Entity  < " & trim(session("EntityName")) & " > updated on " & date 
						%>
							<SCRIPT Language="JavaScript">
							alert("The entity record has been updated successfully.")
							</SCRIPT>

<%	
						end if%>

<%end if%>

<%objConn.close
				
				Set objConn=Nothing
				Set objRec=Nothing
				Set objCmd=Nothing
				
				RecEntityAll.requery
				
			end if	
					
		case "Delete"
		
			session("EntityAct")=""
			if session("EntityID")="" then 
			%>
				<SCRIPT Language="JavaScript">
				alert("Please specify an entity by selecting a current record.")
				</SCRIPT>

<%	
			elseif trim(session("originalName"))=trim(ADMINName()) then
			%>
				<SCRIPT Language="JavaScript">
				alert("Administrator entity can not be deleted.")
				</SCRIPT>

<%	
			elseif not DeletionValid(session("EntityID")) then
			%>
				<SCRIPT Language="JavaScript">
				alert("Entity can not be deleted because there are users associated with it.")
				</SCRIPT>

<%	
			else
				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")
				
				on error resume next
				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
				
				objCmd.CommandText="delete from xca_Entity where EntityID = " & trim(session("EntityID")) 
				set objRec=objCmd.Execute
				'resetSessionValue
				RecEntityAll.requery
				resetSessionValue
				setSessionValue	
				log "C","","",session("UserUserID"),Now,0,"Delete","","Entity" 
				'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
				'email session("AdminEntityEMail"),EMailTo,"","Entity Deleted", "Entity  < " & trim(session("EntityName")) & " > deleted on " & date 
				%>
				<SCRIPT Language="JavaScript">
				alert("The entity has been deleted successfully.")
				</SCRIPT>

<%		
								
				Set objConn=Nothing
				Set objRec=Nothing
				Set objCmd=Nothing
			
			end if	
				
		case else	
		
	end select 	
	session("EntityAct")=""
%>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecEntityAll 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Entity.EntityID,\sxca_Entity.EntityName,\sxca_Types.TypeName,\sxca_status_Security.StatusName,\sxca_Entity.EntityContact,\sxca_Entity.EntityType,\sxca_Entity.EntityStatus,\sxca_Entity.EntityAddress,\sxca_Entity.EntityCity,\sxca_Entity.EntityProvince,\sxca_Entity.EntityPostalCode,\sxca_Entity.EntityEmail,\sxca_Entity.EntityFax,\sxca_Entity.EntityTelephone,\sxca_Entity.EntityExtension\sFROM\sxca_Entity\sINNER\sJOIN\sxca_Types\sON\sxca_Entity.EntityType\s=\sxca_Types.TypeCode\sINNER\sJOIN\sxca_status_Security\sON\sxca_Entity.EntityStatus\s=\sxca_status_Security.StatusCode\sORDER\sBY\sxca_Entity.EntityName\q,TCControlID_Unmatched=\qRecEntityAll\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_PreDefined\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Entity.EntityID,\sxca_Entity.EntityName,\sxca_Types.TypeName,\sxca_status_Security.StatusName,\sxca_Entity.EntityContact,\sxca_Entity.EntityType,\sxca_Entity.EntityStatus,\sxca_Entity.EntityAddress,\sxca_Entity.EntityCity,\sxca_Entity.EntityProvince,\sxca_Entity.EntityPostalCode,\sxca_Entity.EntityEmail,\sxca_Entity.EntityFax,\sxca_Entity.EntityTelephone,\sxca_Entity.EntityExtension\sFROM\sxca_Entity\sINNER\sJOIN\sxca_Types\sON\sxca_Entity.EntityType\s=\sxca_Types.TypeCode\sINNER\sJOIN\sxca_status_Security\sON\sxca_Entity.EntityStatus\s=\sxca_status_Security.StatusCode\sORDER\sBY\sxca_Entity.EntityName\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
					   </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecEntityAll()
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
	cmdTmp.CommandText = 'SELECT xca_Entity.EntityID, xca_Entity.EntityName, xca_Types.TypeName, xca_status_Security.StatusName, xca_Entity.EntityContact, xca_Entity.EntityType, xca_Entity.EntityStatus, xca_Entity.EntityAddress, xca_Entity.EntityCity, xca_Entity.EntityProvince, xca_Entity.EntityPostalCode, xca_Entity.EntityEmail, xca_Entity.EntityFax, xca_Entity.EntityTelephone, xca_Entity.EntityExtension FROM xca_Entity INNER JOIN xca_Types ON xca_Entity.EntityType = xca_Types.TypeCode INNER JOIN xca_status_Security ON xca_Entity.EntityStatus = xca_status_Security.StatusCode ORDER BY xca_Entity.EntityName';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecEntityAll.setRecordSource(rsTmp);
	RecEntityAll.open();
	if (thisPage.getState('pb_RecEntityAll') != null)
		RecEntityAll.setBookmark(thisPage.getState('pb_RecEntityAll'));
}
function _RecEntityAll_ctor()
{
	CreateRecordset('RecEntityAll', _initRecEntityAll, null);
}
function _RecEntityAll_dtor()
{
	RecEntityAll._preserveState();
	thisPage.setState('pb_RecEntityAll', RecEntityAll.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecStatus 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sStatusCode,\sStatusName\sfrom\sxca_status_security\q,TCControlID_Unmatched=\qRecStatus\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_status_codes\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sStatusCode,\sStatusName\sfrom\sxca_status_security\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = 'select StatusCode, StatusName from xca_status_security';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecType 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sTypeCode,\sTypeName\sFROM\sxca_Types\sORDER\sBY\sTypeCode\sDESC\q,TCControlID_Unmatched=\qRecType\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sTypeCode,\sTypeName\sFROM\sxca_Types\sORDER\sBY\sTypeCode\sDESC\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
				  </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecType()
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
	cmdTmp.CommandText = 'SELECT TypeCode, TypeName FROM xca_Types ORDER BY TypeCode DESC';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecType.setRecordSource(rsTmp);
	RecType.open();
	if (thisPage.getState('pb_RecType') != null)
		RecType.setBookmark(thisPage.getState('pb_RecType'));
}
function _RecType_ctor()
{
	CreateRecordset('RecType', _initRecType, null);
}
function _RecType_dtor()
{
	RecType._preserveState();
	thisPage.setState('pb_RecType', RecType.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecProvince 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sProvinceAbbreviation,\sProvinceNane\s\sFROM\sxca_Province\q,TCControlID_Unmatched=\qRecProvince\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sProvinceAbbreviation,\sProvinceNane\s\sFROM\sxca_Province\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=1,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
				  </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecProvince()
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
	cmdTmp.Prepared = true;
	cmdTmp.CommandText = 'SELECT ProvinceAbbreviation, ProvinceNane  FROM xca_Province';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecProvince.setRecordSource(rsTmp);
	RecProvince.open();
	if (thisPage.getState('pb_RecProvince') != null)
		RecProvince.setBookmark(thisPage.getState('pb_RecProvince'));
}
function _RecProvince_ctor()
{
	CreateRecordset('RecProvince', _initRecProvince, null);
}
function _RecProvince_dtor()
{
	RecProvince._preserveState();
	thisPage.setState('pb_RecProvince', RecProvince.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<TABLE WIDTH="75%" ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 
      id=lblTitle style="HEIGHT: 37px; LEFT: 0px; TOP: 0px; WIDTH: 486px" 
      width=486>
	<PARAM NAME="_ExtentX" VALUE="12859">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="CNAS Entity Security Administration">
	<PARAM NAME="FontFace" VALUE="Arial Black">
	<PARAM NAME="FontSize" VALUE="5">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      														  
      </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial Black" SIZE="5" COLOR="Maroon">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTitle()
{
	lblTitle.setCaption('CNAS Entity Security Administration');
}
function _lblTitle_ctor()
{
	CreateLabel('lblTitle', _initlblTitle, null);
}
</script>
<% lblTitle.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>
<table align=center border="0" cellPadding="1" cellSpacing="0"  > .
 <tbody>
     <tr>
        <td align="right" noWrap><font face="Arial" size="2"><strong>Entity 
            Name</strong></font>
       </td><td>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityName style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" 
      width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityName">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityName()
{
	EntityName.setStyle(TXT_TEXTBOX);
	EntityName.setMaxLength(50);
	EntityName.setColumnCount(50);
}
function _EntityName_ctor()
{
	CreateTextbox('EntityName', _initEntityName, null);
}
</script>
<% EntityName.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td align="right" noWrap><strong><font face="Arial" size="2">Entity Type 
            
</font></strong></td>
        <td align="left" noWrap vAlign="center">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=EntityType style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="EntityType">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecType">
	<PARAM NAME="BoundColumn" VALUE="TypeCode">
	<PARAM NAME="ListField" VALUE="TypeName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      															  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityType()
{
	RecType.advise(RS_ONDATASETCOMPLETE, 'EntityType.setRowSource(RecType, \'TypeName\', \'TypeCode\');');
}
function _EntityType_ctor()
{
	CreateListbox('EntityType', _initEntityType, null);
}
</script>
<% EntityType.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="right" noWrap><font face="Arial" size="2"><strong>Entity 
            Contact</strong></font></td>
        <td noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityContact style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" 
      width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityContact">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityContact()
{
	EntityContact.setStyle(TXT_TEXTBOX);
	EntityContact.setMaxLength(35);
	EntityContact.setColumnCount(35);
}
function _EntityContact_ctor()
{
	CreateTextbox('EntityContact', _initEntityContact, null);
}
</script>
<% EntityContact.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td noWrap>
            <div align="right"><strong><font face="Arial" size="2">Entity Status 
            
</font></strong></div></td>
        <td noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=EntityStatus style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="EntityStatus">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecStatus">
	<PARAM NAME="BoundColumn" VALUE="StatusCode">
	<PARAM NAME="ListField" VALUE="StatusName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      															  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityStatus()
{
	RecStatus.advise(RS_ONDATASETCOMPLETE, 'EntityStatus.setRowSource(RecStatus, \'StatusName\', \'StatusCode\');');
}
function _EntityStatus_ctor()
{
	CreateListbox('EntityStatus', _initEntityStatus, null);
}
</script>
<% EntityStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="right" noWrap><strong><font face="Arial" size="2">Entity 
            Address</font></strong></td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityAddress style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" 
      width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityAddress">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="60">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityAddress()
{
	EntityAddress.setStyle(TXT_TEXTBOX);
	EntityAddress.setMaxLength(60);
	EntityAddress.setColumnCount(35);
}
function _EntityAddress_ctor()
{
	CreateTextbox('EntityAddress', _initEntityAddress, null);
}
</script>
<% EntityAddress.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td noWrap align="left">
            <div align="right">
            <div align="right"><strong><font face="Arial" size="2">Entity 
            Email</font></strong></div></div>
</td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityEmail style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
      width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityEmail">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityEmail()
{
	EntityEmail.setStyle(TXT_TEXTBOX);
	EntityEmail.setMaxLength(50);
	EntityEmail.setColumnCount(30);
}
function _EntityEmail_ctor()
{
	CreateTextbox('EntityEmail', _initEntityEmail, null);
}
</script>
<% EntityEmail.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
    <tr>
        <td align="right" noWrap><font face="Arial" size="2"><strong>Entity 
            City</strong></font></td>
        <td noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=EntityCity 
      style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityCity">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityCity()
{
	EntityCity.setStyle(TXT_TEXTBOX);
	EntityCity.setMaxLength(20);
	EntityCity.setColumnCount(20);
}
function _EntityCity_ctor()
{
	CreateTextbox('EntityCity', _initEntityCity, null);
}
</script>
<% EntityCity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td noWrap align="left">
            <div align="right">
            <div align="right"><strong><font face="Arial" size="2">Entity 
            Fax</font></strong></div></div>
</td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityFax style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" 
      width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityFax">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="12">
	<PARAM NAME="DisplayWidth" VALUE="12">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityFax()
{
	EntityFax.setStyle(TXT_TEXTBOX);
	EntityFax.setMaxLength(12);
	EntityFax.setColumnCount(12);
}
function _EntityFax_ctor()
{
	CreateTextbox('EntityFax', _initEntityFax, null);
}
</script>
<% EntityFax.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
    <tr>
        <td align="right" noWrap><strong><font face="Arial" size="2">Entity 
            Province</font></strong></td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=EntityProvince style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="EntityProvince">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecProvince">
	<PARAM NAME="BoundColumn" VALUE="ProvinceAbbreviation">
	<PARAM NAME="ListField" VALUE="ProvinceNane">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      															  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityProvince()
{
	RecProvince.advise(RS_ONDATASETCOMPLETE, 'EntityProvince.setRowSource(RecProvince, \'ProvinceNane\', \'ProvinceAbbreviation\');');
}
function _EntityProvince_ctor()
{
	CreateListbox('EntityProvince', _initEntityProvince, null);
}
</script>
<% EntityProvince.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td noWrap align="left">
            <div align="right"><strong><font face="Arial" size="2">Entity 
            Telephone</font></strong></div>
</td>
        <td noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityTelephone style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" 
      width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityTelephone">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="12">
	<PARAM NAME="DisplayWidth" VALUE="12">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityTelephone()
{
	EntityTelephone.setStyle(TXT_TEXTBOX);
	EntityTelephone.setMaxLength(12);
	EntityTelephone.setColumnCount(12);
}
function _EntityTelephone_ctor()
{
	CreateTextbox('EntityTelephone', _initEntityTelephone, null);
}
</script>
<% EntityTelephone.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td></tr>
    <tr>
        <td align="right" noWrap>
            <div align="right"><strong><font face="Arial" size="2">Entity Postal 
            Code</font></strong></div>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityPostalCode style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 36px" 
      width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityPostalCode">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="6">
	<PARAM NAME="DisplayWidth" VALUE="6">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityPostalCode()
{
	EntityPostalCode.setStyle(TXT_TEXTBOX);
	EntityPostalCode.setMaxLength(6);
	EntityPostalCode.setColumnCount(6);
}
function _EntityPostalCode_ctor()
{
	CreateTextbox('EntityPostalCode', _initEntityPostalCode, null);
}
</script>
<% EntityPostalCode.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        <td noWrap align="right"><strong><font face="Arial" size="2">Entity 
            Extension</font></strong>
        <td noWrap align="left">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityExtension style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 30px" 
      width=30>
	<PARAM NAME="_ExtentX" VALUE="794">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityExtension">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="5">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityExtension()
{
	EntityExtension.setStyle(TXT_TEXTBOX);
	EntityExtension.setMaxLength(4);
	EntityExtension.setColumnCount(5);
}
function _EntityExtension_ctor()
{
	CreateTextbox('EntityExtension', _initEntityExtension, null);
}
</script>
<% EntityExtension.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    <tr>
        <td align="right" noWrap>
        <td align="left" noWrap>
        <td noWrap align="right" style="FONT-FAMILY: sans-serif; FONT-SIZE: x-small; FONT-WEIGHT: bold" 
       >Entity ID
        <td noWrap align="left">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=EntityID style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 48px" 
      width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="8">
	<PARAM NAME="DisplayWidth" VALUE="8">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityID()
{
	EntityID.setStyle(TXT_TEXTBOX);
	EntityID.disabled = true;
	EntityID.setMaxLength(8);
	EntityID.setColumnCount(8);
}
function _EntityID_ctor()
{
	CreateTextbox('EntityID', _initEntityID, null);
}
</script>
<% EntityID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
    </td>

</tr> 

</tbody>
</table>

<BR>
<TABLE ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
	<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnGetCurrent style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 93px" 
      width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGetCurrent">
	<PARAM NAME="Caption" VALUE="Get Current">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      									  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGetCurrent()
{
	btnGetCurrent.value = 'Get Current';
	btnGetCurrent.setStyle(0);
}
function _btnGetCurrent_ctor()
{
	CreateButton('btnGetCurrent', _initbtnGetCurrent, null);
}
</script>
<% btnGetCurrent.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnUpdate style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 120px" 
      width=120>
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update Current">
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
	btnUpdate.value = 'Update Current';
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
		<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnAdd style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 78px" width=78>
	<PARAM NAME="_ExtentX" VALUE="2064">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAdd">
	<PARAM NAME="Caption" VALUE="Add New">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      									  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnAdd()
{
	btnAdd.value = 'Add New';
	btnAdd.setStyle(0);
}
function _btnAdd_ctor()
{
	CreateButton('btnAdd', _initbtnAdd, null);
}
</script>
<% btnAdd.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnDelete style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 115px" 
      width=115>
	<PARAM NAME="_ExtentX" VALUE="3043">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE="Delete Current">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      									  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnDelete()
{
	btnDelete.value = 'Delete Current';
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnClearSceen 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 748px; WIDTH: 101px" width=101>
	<PARAM NAME="_ExtentX" VALUE="2672">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnClearSceen">
	<PARAM NAME="Caption" VALUE="Clear Sceen">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnClearSceen()
{
	btnClearSceen.value = 'Clear Sceen';
	btnClearSceen.setStyle(0);
}
function _btnClearSceen_ctor()
{
	CreateButton('btnClearSceen', _initbtnClearSceen, null);
}
</script>
<% btnClearSceen.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMain 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 775px; WIDTH: 61px" width=61>
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

<BR><TABLE width="75%" ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 
      id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 704px" 
      width=704>
	<PARAM NAME="_ExtentX" VALUE="18627">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="RecEntityAll">
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
	<PARAM NAME="ColumnsNames" VALUE='"EntityName","TypeName","StatusName","EntityContact"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3">
	<PARAM NAME="displayWidth" VALUE="197,102,136,285">
	<PARAM NAME="Coltype" VALUE="1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Entity Name","Type","Status","Contact"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,">
	<PARAM NAME="HeaderFont" VALUE=",,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,">
	<PARAM NAME="DetailFont" VALUE=",,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,">
	<PARAM NAME="ColumnCount" VALUE="4">
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
	<PARAM NAME="GridWidth" VALUE="704">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../">
	
      																																																																							 
       </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 5;
Grid1.setDataSource(RecEntityAll);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=4 rules=ALL WIDTH=704';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=197';
Grid1.headerWidth[1] = ' WIDTH=102';
Grid1.headerWidth[2] = ' WIDTH=136';
Grid1.headerWidth[3] = ' WIDTH=285';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'Entity Name\'';
Grid1.colHeader[1] = '\'Type\'';
Grid1.colHeader[2] = '\'Status\'';
Grid1.colHeader[3] = '\'Contact\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=197';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'RecEntityAll.fields.getValue(\'EntityName\')';
Grid1.colAttributes[1] = '  WIDTH=102';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'RecEntityAll.fields.getValue(\'TypeName\')';
Grid1.colAttributes[2] = '  WIDTH=136';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'RecEntityAll.fields.getValue(\'StatusName\')';
Grid1.colAttributes[3] = '  WIDTH=285';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'RecEntityAll.fields.getValue(\'EntityContact\')';
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
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
	'on error resume next
	objCmd1.CommandText=	"Select EntityType from xca_Entity where EntityID =  " & EntityID
	set objRec1=objCmd1.Execute
	if objRec1.EOF then
		AssociationValidTemp=false
		
	elseif	objRec1("EntityType")<>"a" then
		
		AssociationValidTemp=false
	else
		
		AssociationValidTemp=true
	end if
	objConn1.close
	Set objConn1=Nothing
	Set objRec1=Nothing
	Set objCmd1=Nothing
	
	AssociationValid=AssociationValidTemp
	
end function

sub resetSessionValue()
	
	session("UserName")=""
	session("UserLogon")=""
	session("UserPassword")=""
	session("EntityID")=""
	session("UserType")=""
	session("UserStatus")=""
	session("UserTelephone")=""
	session("UserExtension")=""
	session("UserFax")=""
	session("UserEmail")=""
	session("UserAddress")=""
	session("UserCity")=""
	session("UserProvince")=""
	session("UserPostalCode")=""
	
end sub

sub setSessionValue()

	UserName.value=session("UserName")
	UserLogon.value=session("UserLogon")
	UserPassword.value=session("UserPassword")
	
	EntityName.SelectByValue(session("EntityID"))
	UserType.SelectByValue(session("UserType"))
	UserStatus.SelectByValue(session("UserStatus"))
	
	'Response.Write session("EntityID") & "||" & session("UserType") & "||" & session("UserStatus")
	UserTelephone.value=session("UserTelephone")
	UserExtension.value=session("UserExtension")
	UserFax.value=session("UserFax")
	UserEmail.value=session("UserEmail")
	UserAddress.value=session("UserAddress")
	UserAddress.value=session("UserCity")
	UserProvince.SelectByValue(session("UserProvince"))
	UserPostalCode.value=session("UserPostalCode")

end sub

sub getSessionValue()
	session("UserName")=UserName.value
	session("UserLogon")=UserLogon.value
	session("UserPassword")=UserPassword.value
	
	txt=EntityName.selectedIndex
	session("EntityID")=EntityName.getValue(txt)
	
	txt=UserType.selectedIndex
	session("UserType")=UserType.getValue(txt)
	
	txt=UserStatus.selectedIndex
	session("UserStatus")=UserStatus.getValue(txt)
	
	'Response.Write session("EntityID") & "||" & session("UserType") & "||" & session("UserStatus")
	session("UserTelephone")=UserTelephone.value
	session("UserExtension")=UserExtension.value
	session("UserFax")=UserFax.value
	session("UserEmail")=UserEmail.value
	session("UserAddress")=UserAddress.value
	session("UserCity")=UserCity.value
	txt=UserProvince.selectedIndex
	session("UserProvince")=UserProvince.getValue(txt)
	session("UserPostalCode")=UserPostalCode.value
end sub

Sub btnUpdate_onclick()
	
	session("UserAct")="Update"
	getSessionValue
	'Response.Redirect "xca_UserOnly.asp"	

End Sub

Sub btnAdd_onclick()
	
	session("UserAct")="Add"
	getSessionValue
	'Response.Redirect "xca_UserOnly.asp"

End Sub

Sub btnDelete_onclick()
	
	session("UserAct")="Delete"
	getSessionValue
	Response.Redirect "xca_UserOnly.asp"
	
	'rec1.deleteRecord
						
End Sub

Sub btnGetCurrent_onclick()

	UserName.value=Trim(RecUser.fields.getvalue("UserName"))
	UserLogon.value=Trim(RecUser.fields.getvalue("UserLogon"))
	UserPassword.value=Trim(RecUser.fields.getvalue("UserPassword"))
	
	EntityName.SelectByText(RecUser.fields.getvalue("EntityName"))
	UserType.SelectByText(trim(RecUser.fields.getvalue("TypeName")))
	UserStatus.SelectByText(trim(RecUser.fields.getvalue("StatusName")))
	
	UserTelephone.value=Trim(RecUser.fields.getvalue("UserTelephone"))
	UserExtension.value=Trim(RecUser.fields.getvalue("UserExtension"))
	UserFax.value=Trim(RecUser.fields.getvalue("UserFax"))
	UserEmail.value=Trim(RecUser.fields.getvalue("UserEmail"))
	UserAddress.value=Trim(RecUser.fields.getvalue("UserAddress"))
	UserCity.value=Trim(RecUser.fields.getvalue("UserCity"))
	UserProvince.SelectByText(trim(RecUser.fields.getvalue("UserProvince")))
	UserPostalCode.value=Trim(RecUser.fields.getvalue("UserPostalCode"))
	
End Sub

Sub btnReturnToMain_onclick()
	Response.Redirect "xca_MenuSecurityAdmin.asp"
End Sub

Sub btnClearScreen_onclick()
	ResetSessionValue
	setSessionValue
	EntityName.SelectByText(EntityName.getText(0))
	UserType.SelectByText(UserType.getText(0))
	UserStatus.SelectByText(UserStatus.getText(0))
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor="#d7c7a4">
<%
	'setSessionValue
	Select Case session("UserAct")		
		Case "Add"
			session("UserAct")=""
			Set objConn=server.CreateObject("ADODB.Connection")
			Set objRec=server.CreateObject("ADODB.Recordset")
			Set objCmd=server.CreateObject("ADODB.Command")
	
			objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
			objCmd.ActiveConnection = objConn
	
			on error resume next
			objCmd.CommandText="Select UserLogon from xca_User where UserLogon = '" & Replace(trim(session("UserLogon")) ,"'","''")& "'"
			set objRec=objCmd.Execute%>
	
			<%if not objRec.EOF then  %>
				<SCRIPT Language="JavaScript">
				alert("The user is already in the data base. It can not be added again.")
				</SCRIPT>
			<%else
				if session("UserName")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("User name must be entered to add a record.")
					</SCRIPT>
				<%
				elseif session("UserLogon")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User logon must be entered to add a record.")
					</SCRIPT>
				<%
				elseif session("UserPassword")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("User password must be entered to add a record.")
					</SCRIPT>
				<%
				elseif session("EntityID")="" then
				%>	<SCRIPT Language="JavaScript">
					alert("Entity must be entered to add a record.")
					</SCRIPT>
				<%
				
				elseif session("UserType")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User type must be entered to add a record.")
					</SCRIPT>	
				<%
				elseif session("UserType")="a" and (not AssociationValid(session("EntityID"))) then
				
				%>
					<SCRIPT Language="JavaScript">
					alert("An administrative user can not be associated with a non-administrative entity.")
					</SCRIPT>	
				<%
				elseif session("UserStatus")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User status must be entered to add a record.")
					</SCRIPT>	
				<%
				elseif session("UserTelephone")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User telephone must be entered to add a record.")
					</SCRIPT>	
				<%
				elseif session("UserEmail")="" then
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
					objCmd.CommandText=	"AddUser " & Replace(trim(session("EntityID")),"'","''") _
													& ", '" & Replace(trim(session("UserType")),"'","''") _
													& "', '"& Replace(session("UserStatus"),"'","''") _ 
													& "', '"& Replace(session("UserName"),"'","''") _
													& "', '"& Replace(session("UserTelephone"),"'","''") _
													& "', '"& Replace(session("UserExtension"),"'","''") _
													& "', '"& Replace(session("UserFax"),"'","''") _
													& "', '"& Replace(session("UserEmail"),"'","''") _
													& "', '"& Replace(session("UserLogon"),"'","''") _
													& "', '"& Replace(session("UserPassword"),"'","''") _
													& "', '"& Replace(session("UserAddress"),"'","''") _
	& "', '"& Replace(session("UserCity"),"'","''") _
	& "', '"& Replace(session("UserProvince"),"'","''") _
	& "', '"& Replace(session("UserPostalCode"),"'","''") _
	& "'"
					objCmd.Execute %>
				
					<%if objConn.Errors.Count <> 0 then  %>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while adding the user.")
						</SCRIPT>
					<%else
						log "C","","",session("UserUserID"),Now,0,"Add","","User" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","User Added", "User  < " & trim(session("UserLogon")) & " > added on " & date 
						resetSessionValue
						setSessionValue
					%>
						<SCRIPT Language="JavaScript">
						alert("The user record has been added successfully.")
						</SCRIPT>
					<%	
					end if%>
					
				<%end if%>
				
			<% end if%>
	
			<%objConn.close
			
			Set objConn=Nothing
			Set objRec=Nothing
			Set objCmd=Nothing
			
			'resetSessionValue
			
			RecUser.requery
		
		case "Update"	
			session("UserAct")=""
			if session("UserLogon")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("User logon name is missing.")
					</SCRIPT>
				<%
			else	
				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")
		
				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
		
				on error resume next
				objCmd.CommandText="Select UserLogon from xca_User where UserLogon = '" & Replace(trim(session("UserLogon")),"'","''") & "'"
				set objRec=objCmd.Execute%>
		
				<%if objRec.EOF then  %>
					<SCRIPT Language="JavaScript">
					alert("The user does not exist in the data base. No user is updated.")
					</SCRIPT>
				<%else
					if session("UserName")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("User name must be entered to update a record.")
						</SCRIPT>
					<%
					elseif session("UserPassword")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("User password must be entered to update a record.")
						</SCRIPT>
					<%
					elseif session("EntityID")="" then
					%>	<SCRIPT Language="JavaScript">
						alert("Entity must be entered to update a record.")
						</SCRIPT>
					<%
					elseif session("UserType")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("User type must be entered to update a record.")
						</SCRIPT>	
					<%
					elseif session("UserType")="a" and (not AssociationValid(session("EntityID"))) then
					%>
					<SCRIPT Language="JavaScript">
					alert("An administrative user can not be associated with a non-administrative entity.")
					</SCRIPT>	
					<%
					elseif session("UserStatus")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("User status must be entered to update a record.")
						</SCRIPT>	
					<%
					elseif session("UserTelephone")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("User telephone must be entered to update a record.")
						</SCRIPT>	
					<%
					
					elseif session("UserEmail")="" then
					%>
						<SCRIPT Language="JavaScript">
						alert("User e-mail must be entered to update a record.")
						</SCRIPT>	
					<%
					else
'Added to solve hot fix problem Nov 3 2003
objConn.close
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn

						objCmd.CommandText=	"UpdateUser " & Replace(trim(session("EntityID")),"'","''") _
														& ", '" & Replace(trim(session("UserType")),"'","''") _
														& "', '"& Replace(session("UserStatus"),"'","''") _ 
														& "', '"& Replace(session("UserName"),"'","''") _
														& "', '"& Replace(session("UserTelephone"),"'","''") _
														& "', '"& Replace(session("UserExtension"),"'","''") _
														& "', '"& Replace(session("UserFax"),"'","''") _
														& "', '"& Replace(session("UserEmail"),"'","''") _
														& "', '"& Replace(session("UserLogon"),"'","''") _
														& "', '"& Replace(session("UserPassword"),"'","''") _
														& "', '"& Replace(session("UserAddress"),"'","''") _
		& "', '"& Replace(session("UserCity"),"'","''") _
		& "', '"& Replace(session("UserProvince"),"'","''") _
		& "', '"& Replace(session("UserPostalCode"),"'","''") _
		& "'"
						objCmd.Execute %>
					
						<%if objConn.Errors.Count <> 0 then  %>
							<SCRIPT Language="JavaScript">
							alert("An error has occured while updating the user.")
							</SCRIPT>
						<%else
							log "C","","",session("UserUserID"),Now,0,"Update","","User" 
							'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
							'email session("AdminEntityEMail"),EMailTo,"","User Updated", "User  < " & trim(session("UserLogon")) & " > updated on " & date 
					
						%>
							<SCRIPT Language="JavaScript">
							alert("The user record has been updated successfully.")
							</SCRIPT>
						<%	
						end if%>
						
					<%end if%>
					
				<% end if%>
		
				<%objConn.close
				
				Set objConn=Nothing
				Set objRec=Nothing
				Set objCmd=Nothing
				
				'resetSessionValue
				
				RecUser.requery
			end if			
		case "Delete"
			session("UserAct")=""
			if session("UserLogon")="" then 
			%>
				<SCRIPT Language="JavaScript">
				alert("Please specify a user logon ID by selecting a current record.")
				</SCRIPT>
			<%	
			else
				
				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")
	
				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
	
				on error resume next
				objCmd.CommandText="Select UserLogon from xca_User where UserLogon = '" & trim(session("UserLogon")) & "'"
				set objRec=objCmd.Execute%>
	
				<%if objRec.EOF then  %>
					<SCRIPT Language="JavaScript">
					alert("The user does not exist in the data base.")
					</SCRIPT>
				<%else
'Added to solve hot fix problem Nov 3 2003
objConn.close
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn
				
					objCmd.CommandText= "SELECT xca_Part1.UserID " _
										& "FROM xca_User INNER JOIN " _
										& "xca_Part1 ON xca_User.UserID = xca_Part1.UserID " _
										& "WHERE (xca_User.UserLogon = '" & trim(session("UserLogon")) & "' AND " _ 
										& "(xca_Part1.RequestStatus IN ('NW', 'UP', 'AS ', 'RS ')))"
					set objRec=objCmd.Execute
					%>
				
					<%if not objRec.EOF then %>
						<SCRIPT Language="JavaScript">
						alert("The user can not be deleted because there are associated open tickets.")
						</SCRIPT>
					<%else
						
						'on error resume next
'Added to solve hot fix problem Nov 3 2003
objConn.close
objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
objCmd.ActiveConnection = objConn

						objCmd.CommandText="delete from xca_User where UserLogon = '" & trim(session("UserLogon")) & "'"
						set objRec=objCmd.Execute
						'resetSessionValue
						RecUser.requery
					
						log "C","","",session("UserUserID"),Now,0,"Delete","","User" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","User Deleted", "User  < " & trim(session("UserLogon")) & " > deleted on " & date 
						%>
						<SCRIPT Language="JavaScript">
						alert("The user has been deleted successfully.")
						</SCRIPT>
						<%		
						
					end if	
					
				end if
				
				Set objConn=Nothing
				Set objRec=Nothing
				Set objCmd=Nothing
			
			end if	
				
		case else	
		
	end select 	
	session("UserAct")=""
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\r\nFROM\sxca_PreDefined\r\nORDER\sBY\sPreNXX\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_PreDefined\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\r\nFROM\sxca_PreDefined\r\nORDER\sBY\sPreNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
				  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRec1()
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
	cmdTmp.CommandText = 'SELECT * FROM xca_PreDefined ORDER BY PreNXX';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Rec1.setRecordSource(rsTmp);
	Rec1.open();
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecStatus 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sCOStatus\sfrom\sxca_status_codes\sorder\sby\sCOStatus\q,TCControlID_Unmatched=\qRecStatus\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_status_codes\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sCOStatus\sfrom\sxca_status_codes\sorder\sby\sCOStatus\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = 'select COStatus from xca_status_codes order by COStatus';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecUser 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sa.UserName,\sa.UserLogon,\sa.UserPassword,\sb.EntityName,\sc.TypeName,\sd.StatusName,\sa.UserTelephone,\sa.UserExtension,\sa.UserFax,\sa.UserEmail,\sa.UserType,\sa.UserStatus,\sa.UserAddress,\sa.UserCity,\sa.UserProvince,\sa.UserPostalCode\sFROM\sxca_User\sa\sINNER\sJOIN\sxca_Entity\sb\sON\sa.EntityID\s=\sb.EntityID\sINNER\sJOIN\sxca_Types_User\sc\sON\sa.UserType\s=\sc.TypeCode\sINNER\sJOIN\sxca_status_Security\sd\sON\sa.UserStatus\s=\sd.StatusCode\sORDER\sBY\sa.UserLogon\q,TCControlID_Unmatched=\qRecUser\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sa.UserName,\sa.UserLogon,\sa.UserPassword,\sb.EntityName,\sc.TypeName,\sd.StatusName,\sa.UserTelephone,\sa.UserExtension,\sa.UserFax,\sa.UserEmail,\sa.UserType,\sa.UserStatus,\sa.UserAddress,\sa.UserCity,\sa.UserProvince,\sa.userPostalCode\sFROM\sxca_User\sa\sINNER\sJOIN\sxca_Entity\sb\sON\sa.EntityID\s=\sb.EntityID\sINNER\sJOIN\sxca_Types_User\sc\sON\sa.UserType\s=\sc.TypeCode\sINNER\sJOIN\sxca_status_Security\sd\sON\sa.UserStatus\s=\sd.StatusCode\sORDER\sBY\sa.UserLogon\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
				  </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecUser()
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
	cmdTmp.CommandText = 'SELECT a.UserName, a.UserLogon, a.UserPassword, b.EntityName, c.TypeName, d.StatusName, a.UserTelephone, a.UserExtension, a.UserFax, a.UserEmail, a.UserType, a.UserStatus, a.UserAddress, a.UserCity, a.UserProvince, a.UserPostalCode FROM xca_User a INNER JOIN xca_Entity b ON a.EntityID = b.EntityID INNER JOIN xca_Types_User c ON a.UserType = c.TypeCode INNER JOIN xca_status_Security d ON a.UserStatus = d.StatusCode ORDER BY a.UserLogon';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	RecUser.setRecordSource(rsTmp);
	RecUser.open();
	if (thisPage.getState('pb_RecUser') != null)
		RecUser.setBookmark(thisPage.getState('pb_RecUser'));
}
function _RecUser_ctor()
{
	CreateRecordset('RecUser', _initRecUser, null);
}
function _RecUser_dtor()
{
	RecUser._preserveState();
	thisPage.setState('pb_RecUser', RecUser.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecEntity 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sEntityID,\sEntityName\sfrom\sxca_Entity\sorder\sby\sEntityName\q,TCControlID_Unmatched=\qRecEntity\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sEntityID,\sEntityName\sfrom\sxca_Entity\sorder\sby\sEntityName\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=1,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.Prepared = true;
	cmdTmp.CommandText = 'Select EntityID, EntityName from xca_Entity order by EntityName';
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
      id=lblTitle style="HEIGHT: 37px; LEFT: 10px; TOP: 350px; WIDTH: 470px" 
      width=470>
	<PARAM NAME="_ExtentX" VALUE="12435">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="CNAS User Security Administration">
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
	lblTitle.setCaption('CNAS User Security Administration');
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
<table align=center border="0" cellPadding="1" cellSpacing="0"  > 
 <tbody>
  <TR>
 <td align="right" noWrap borderColor="#7ba89a"><strong><font face="Arial" size="2">User Name</font></strong></td>
        <td noWrap align="left" >
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserName style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
      width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserName">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUserName()
{
	UserName.setStyle(TXT_TEXTBOX);
	UserName.setMaxLength(35);
	UserName.setColumnCount(30);
}
function _UserName_ctor()
{
	CreateTextbox('UserName', _initUserName, null);
}
</script>
<% UserName.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
       </td><td noWrap align="right" borderColor="#7ba89a">
        <strong>
        <font face="Arial" size="2">User Status</font></strong></td>
<td align="left" noWrap borderColor="#d7c7a4">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=UserStatus style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 80px" 
      width=80>
	<PARAM NAME="_ExtentX" VALUE="2117">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="UserStatus">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="serStatus">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="2">
	<PARAM NAME="CLED1" VALUE="Active">
	<PARAM NAME="CLEV1" VALUE="a">
	<PARAM NAME="CLED2" VALUE="Inactive">
	<PARAM NAME="CLEV2" VALUE="i">
	<PARAM NAME="LocalPath" VALUE="../">
	
      																  </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUserStatus()
{
	UserStatus.addItem('Active', 'a');
	UserStatus.addItem('Inactive', 'i');
	UserStatus.setDataField('serStatus');
}
function _UserStatus_ctor()
{
	CreateListbox('UserStatus', _initUserStatus, null);
}
</script>
<% UserStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan--> 
</td></TR><tr>
<td align="right" noWrap borderColor="#d7c7a4"><strong>
<font face="Arial" size="2">Logon ID</font></strong></td>
<td align="left" noWrap borderColor="#d7c7a4">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserLogon style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
      width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserLogon">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="12">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUserLogon()
{
	UserLogon.setStyle(TXT_TEXTBOX);
	UserLogon.setMaxLength(10);
	UserLogon.setColumnCount(10);
}
function _UserLogon_ctor()
{
	CreateTextbox('UserLogon', _initUserLogon, null);
}
</script>
<% UserLogon.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                   
</td>
<td align="right" noWrap>
<strong>
<font face="Arial" size="2">User 
            Telephone</font></strong></td>
<td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserTelephone style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" 
      width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserTelephone">
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
function _initUserTelephone()
{
	UserTelephone.setStyle(TXT_TEXTBOX);
	UserTelephone.setMaxLength(12);
	UserTelephone.setColumnCount(12);
}
function _UserTelephone_ctor()
{
	CreateTextbox('UserTelephone', _initUserTelephone, null);
}
</script>
<% UserTelephone.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td> </tr> <tr>
<td align="right" noWrap><strong>
<font face="Arial" size="2">Password</font></strong></td>

<td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserPassword style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" 
      width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserPassword">
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
function _initUserPassword()
{
	UserPassword.setStyle(TXT_TEXTBOX);
	UserPassword.setMaxLength(12);
	UserPassword.setColumnCount(12);
}
function _UserPassword_ctor()
{
	CreateTextbox('UserPassword', _initUserPassword, null);
}
</script>
<% UserPassword.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
<td align="right" noWrap><strong>
<font face="Arial" size="2">User 
            Extension</font></strong></td>
 <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserExtension style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 30px" 
      width=30>
	<PARAM NAME="_ExtentX" VALUE="794">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserExtension">
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
function _initUserExtension()
{
	UserExtension.setStyle(TXT_TEXTBOX);
	UserExtension.setMaxLength(4);
	UserExtension.setColumnCount(5);
}
function _UserExtension_ctor()
{
	CreateTextbox('UserExtension', _initUserExtension, null);
}
</script>
<% UserExtension.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
                    
</td>
 </tr> <tr>      
 <td align="right" noWrap>
<strong> 
 <font face="Arial" size="2">Entity 
            Name</font></strong></td>
 <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=EntityName style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="EntityName">
	<PARAM NAME="DataSource" VALUE="RecEntity">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="RecEntity">
	<PARAM NAME="BoundColumn" VALUE="EntityID">
	<PARAM NAME="ListField" VALUE="EntityName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      															  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityName()
{
	RecEntity.advise(RS_ONDATASETCOMPLETE, 'EntityName.setRowSource(RecEntity, \'EntityName\', \'EntityID\');');
	EntityName.setDataSource(RecEntity);
	EntityName.setDataField('EntityName');
}
function _EntityName_ctor()
{
	CreateListbox('EntityName', _initEntityName, null);
}
</script>
<% EntityName.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
<td align="right" noWrap><strong>
<font face="Arial" size="2">User 
            Fax</font></strong></td>
 <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserFax style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" 
width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserFax">
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
function _initUserFax()
{
	UserFax.setStyle(TXT_TEXTBOX);
	UserFax.setMaxLength(12);
	UserFax.setColumnCount(12);
}
function _UserFax_ctor()
{
	CreateTextbox('UserFax', _initUserFax, null);
}
</script>
<% UserFax.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</td> </tr> <tr>     
<td align="right" noWrap>
 <strong>
<font face="Arial" size="2">User 
            Type</font></strong></td>
 <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=UserType style="HEIGHT: 21px; LEFT: 10px; TOP: 543px; WIDTH: 84px" 
      width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="UserType">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="3">
	<PARAM NAME="CLED1" VALUE="User">
	<PARAM NAME="CLEV1" VALUE="u">
	<PARAM NAME="CLED2" VALUE="Admin">
	<PARAM NAME="CLEV2" VALUE="a">
	<PARAM NAME="CLED3" VALUE="Manager">
	<PARAM NAME="CLEV3" VALUE="m">
	<PARAM NAME="LocalPath" VALUE="../">
	
      																		  </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUserType()
{
	UserType.addItem('User', 'u');
	UserType.addItem('Admin', 'a');
	UserType.addItem('Manager', 'm');
}
function _UserType_ctor()
{
	CreateListbox('UserType', _initUserType, null);
}
</script>
<% UserType.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
  <td align="right" noWrap><strong>
  <font face="Arial" size="2">User 
            Email</font></strong></td>
  <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserEmail style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 156px" 
      width=156>
	<PARAM NAME="_ExtentX" VALUE="4128">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserEmail">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="26">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      													  
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initUserEmail()
{
	UserEmail.setStyle(TXT_TEXTBOX);
	UserEmail.setMaxLength(50);
	UserEmail.setColumnCount(26);
}
function _UserEmail_ctor()
{
	CreateTextbox('UserEmail', _initUserEmail, null);
}
</script>
<% UserEmail.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
</tr>
<tr> 
  <td align="right" noWrap><strong>
  <font face="Arial" size="2">Address</font></strong></td>
  <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserAddress style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 156px" 
      width=156>
	<PARAM NAME="_ExtentX" VALUE="4128">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserAddress">
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
function _initUserAddress()
{
	UserAddress.setStyle(TXT_TEXTBOX);
	UserAddress.setMaxLength(60);
	UserAddress.setColumnCount(35);
}
function _UserAddress_ctor()
{
	CreateTextbox('UserAddress', _initUserAddress, null);
}
</script>
<% UserAddress.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
  <td align="right" noWrap><strong>
  <font face="Arial" size="2">City</font></strong></td>
  <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserCity style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 156px" 
      width=156>
	<PARAM NAME="_ExtentX" VALUE="4128">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserCity">
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
function _initUserCity()
{
	UserCity.setStyle(TXT_TEXTBOX);
	UserCity.setMaxLength(20);
	UserCity.setColumnCount(20);
}
function _UserCity_ctor()
{
	CreateTextbox('UserCity', _initUserCity, null);
}
</script>
<% UserCity.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
    <tr>
        <td align="right" noWrap><strong><font face="Arial" size="2">Province</font></strong></td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
      id=UserProvince style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
      width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="UserProvince">
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
function _initUserProvince()
{
	RecProvince.advise(RS_ONDATASETCOMPLETE, 'UserProvince.setRowSource(RecProvince, \'ProvinceNane\', \'ProvinceAbbreviation\');');
}
function _EntityProvince_ctor()
{
	CreateListbox('UserProvince', _initUserProvince, null);
}
</script>
<% UserProvince.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
  <td align="right" noWrap><strong>
  <font face="Arial" size="2">Postal Code</font></strong></td>
  <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=UserPostalCode style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 156px" 
      width=156>
	<PARAM NAME="_ExtentX" VALUE="4128">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="UserPostalCode">
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
function _initUserPostalCode()
{
	UserPostalCode.setStyle(TXT_TEXTBOX);
	UserPostalCode.setMaxLength(6);
	UserPostalCode.setColumnCount(6);
}
function _UserPostalCode_ctor()
{
	CreateTextbox('UserPostalCode', _initUserPostalCode, null);
}
</script>
<% UserPostalCode.display %>

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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnClearScreen 
      style="HEIGHT: 27px; LEFT: 10px; TOP: 691px; WIDTH: 106px" width=106>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>

<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnReturnToMain 
      style="HEIGHT: 27px; LEFT: 10px; TOP: 718px; WIDTH: 61px" width=61>
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
	<PARAM NAME="Recordset" VALUE="RecUser">
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
	<PARAM NAME="ColumnsNames" VALUE='"UserName","UserLogon","EntityName","TypeName","StatusName"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4">
	<PARAM NAME="displayWidth" VALUE="179,146,253,68,68">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"User Name","User Logon","Entity Name","Type","Status"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,">
	<PARAM NAME="ColumnCount" VALUE="5">
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
Grid1.setDataSource(RecUser);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=5 rules=ALL WIDTH=704';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=179';
Grid1.headerWidth[1] = ' WIDTH=146';
Grid1.headerWidth[2] = ' WIDTH=253';
Grid1.headerWidth[3] = ' WIDTH=68';
Grid1.headerWidth[4] = ' WIDTH=68';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'User Name\'';
Grid1.colHeader[1] = '\'User Logon\'';
Grid1.colHeader[2] = '\'Company\'';
Grid1.colHeader[3] = '\'Type\'';
Grid1.colHeader[4] = '\'Status\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=179';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'RecUser.fields.getValue(\'UserName\')';
Grid1.colAttributes[1] = '  WIDTH=146';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'RecUser.fields.getValue(\'UserLogon\')';
Grid1.colAttributes[2] = '  WIDTH=253';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'RecUser.fields.getValue(\'EntityName\')';
Grid1.colAttributes[3] = '  WIDTH=68';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'RecUser.fields.getValue(\'TypeName\')';
Grid1.colAttributes[4] = '  WIDTH=68';
Grid1.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[4] = 'RecUser.fields.getValue(\'StatusName\')';
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
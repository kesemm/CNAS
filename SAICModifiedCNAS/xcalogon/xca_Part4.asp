<%@ Language=VBScript %>
<%
session("undo")=""	
Response.Buffer=true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="xca_CNASLib.inc"-->
<HTML>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%

UserEntityID =int(session("UserEntityID"))
session("P4UserEntityID")=UserEntityID

Dim Part4NPA,Part4NXX,P4SwitchID,P4DateOfApplication2,txtAuthorizedID,P4DateOfReceipt

function checkNull(c) 
	dim temp
	if isnull(c) then 
		temp=""
	else  temp=c 	
	end if
	checkNull=temp
end function

function GetNewPart4Data(pNPA,pNXX)
'function GetPart4Data(pNPA,pNXX)

	dim objConn
	dim objCmd
	dim objRec

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

	objCmd.CommandText="Get_NewPart4_Data " & pNPA & ", " & pNXX 
	'objCmd.CommandText="Get_Part4_Data " & pNPA & ", " & pNXX 
			
	set objRec=objCmd.Execute
	if not objRec.EOF then
		
		session("pSignature")=session("UserUserLogon")
		P4Part4Date=date
		
		session("P4EntityID")=checkNull(objRec("EntityID"))
		session("pPart4NPA")=checkNull(objRec("AssignedNPA"))
		session("pPart4NXX")=checkNull(objRec("AssignedNXX"))
		session("pSwitchID")=checkNull(objRec("SwitchID"))
		session("pApplicationDate")=checkNull(objRec("ApplicationDate"))
		session("P3DateofReceipt")=checkNull(objRec("P3DateofReceipt"))
		session("ReceiptDate")=checkNull(objRec("DateofReceipt"))
		session("P3EffDate")=checkNull(objRec("EffectiveDate"))
		session("Tix")=checkNull(objRec("Tix"))
		session("P1EntityEmail")=checkNull(objRec("EntityEmail"))
		session("P1UserEmail")=checkNull(objRec("UserEmail"))

		GetNewPart4DataTemp=true
		'GetPart4DataTemp=true
	else
		GetNewPart4DataTemp=false						
		'GetPart4DataTemp=false						
	end if	
				
	objRec.close
	objConn.close
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing

	GetNewPart4Data=GetNewPart4DataTemp
	'GetPart4Data=GetPart4DataTemp
end function

function GetAdminEmail()

	dim GetAdminEmailTemp
	dim objConn
	dim objCmd
	dim objRec

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	objCmd.CommandText="GetAdminEMAil"
	set objRec=objCmd.Execute
	if not objRec.EOF then
		GetAdminEmailTemp=objRec("EntityEmail")
	else
		GetAdminEmailTemp=""					
	end if	
				
	objRec.close
	objConn.close
	
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	
	GetAdminEmail=GetAdminEmailTemp
	
end function

function GetRsTicket()

	dim GetRsTicketTemp
	dim objConn
	dim objCmd
	dim objRec

	Set objConn=server.CreateObject("ADODB.Connection")
	Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	objCmd.CommandText="Select Tix,NPA, NXX1preferred From xca_Part1 where RequestStatus='RS' and NPA= '"&session("pPart4NPA")&"' and NXX1preferred='"&session("pPart4NXX")&"'"
	'objCmd.CommandText="Select Tix, NPA, NXX1preferred From xca_Part1 where RequestStatus='RS' and NPA= '"&session("pPart4NPA")&"' and NXX1preferred='"&session("pPart4NXX")&"' and Tix= '"&session("Tix")&"'"
	
	set objRec=objCmd.Execute
	if not objRec.EOF then
		GetRsTicketTemp=objRec("Tix")
		GetRsTicketTemp1=objRec("NPA")
		GetRsTicketTemp2=objRec("NXX1Preferred")
		objCmd.CommandText="Update xca_Part1 set RequestStatus='CA' where NPA= '"&GetRsTicketTemp1&"' and NXX1Preferred= '"&GetRsTicketTemp2&"' and RequestStatus= 'RS'"
		objCmd.Execute
	else
		GetRsTicketTemp=""					
	end if	
				
	objRec.close
	objConn.close
	
	Set objConn=Nothing
	Set objRec=Nothing
	Set objCmd=Nothing
	
	GetRsTicket=GetRsTicketTemp
	
end function

''''''''''''''''''''''''''''''////Adding data to Part4 Table\\\\'''''''''''''''''''''''''''
'function AddPart4Data()
function AddPart4bweData()

	dim objConn
	dim objCmd
	'dim objRec
	
	Set objConn=server.CreateObject("ADODB.Connection")
	'Set objRec=server.CreateObject("ADODB.Recordset")
	Set objCmd=server.CreateObject("ADODB.Command")
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn

	objCmd.CommandText="AddPart4bwe_InService_Trans '" & cint(session("Tix")) _
													& "','"	& session("pPart4Name") _ 
													& "','" & session("pPart4Title") _ 
													& "','" & session("pPart4Date") _ 
													& "','" & session("pPart4NPA") _
													& "','"	& session("pPart4NXX") _
													& "','" & session("pSwitchID") _ 
													& "','" & session("pApplicationDate") _ 
													& "','" & session("pInServiceDate") _
													& "','" & session("pSignature") _ 
													& "','" & session("P4UserEntityID") _
													& "','" & Date() & "'"

	objCmd.Execute
	if objConn.Errors.Count <> 0 then 
		AddPart4bweDataTemp=false
	else
		AddPart4bweDataTemp=true
	end if	
	
	objConn.close
	Set objConn=Nothing
	Set objCmd=Nothing
	AddPart4bweData=AddPart4bweDataTemp
	
end function
	
sqlp = "Select Value From xca_Parms Where Name = 'P4DATE'"
	GetParmsData.setSQLText(sqlp)
	GetParmsData.Open
	session("P4DATEValue")=GetParmsData.fields.getValue("Value")
''''''''''P4DATEValue is being read'''''''''''''
	'Response.Write session("P4DATEValue") & "TEST"
			
Sub btnSubmit_onclick()
session("undo")=""
'Response.Write "This is the ppp P4DATEValue->" & session("P4DATEValue")
ttt=Trim(session("P4DATEValue"))
Select Case ttt
Case "NO"
'If P4DATEValue="NO" Then
'Response.Write "This is the NO P4DATEValue->" & session("P4DATEValue")
	if session("Part4Act")="done" then 
		Response.Redirect "xca_Part4Deny.asp"
	else		
	
		session("pPart4Name")=Replace(txtAuthorizedRep.value,"'","''")
		session("pPart4Title")=Replace(txtAuthorizedRepTitle.value,"'","''")
		session("pPart4Date")=Date()
		session("pInServiceDate")=Replace(txtInServiceDate.value,"'","''")
		'session("pSignature")=txtAuthorizedID
		session("AuthorizedRep")=Replace(txtAuthorizedRep.value,"'","''")
		session("AuthorizedRepTitle")=Replace(txtAuthorizedRepTitle.value,"'","''")
		session("Part4Act")="Submit"
		
			if session("pInServiceDate")="" then
 				session("undo")="MissingDate"
 				 			
 			elseif not IsDateReal(session("pInServiceDate")) then
 				session("undo")="MissingFormat"
 				
  			elseif datediff("d", session("P3EffDate"), session("pInServiceDate")) < 0 then
 				session("undo")="InCorrectDate"
 			
 			'elseif datediff("d", session("ReceiptDate"),session("pInServiceDate")) < cint(GetP4DAYS()) then
 			 	'undo=true	
 			'session("undo")="MissingEarly"		
 			
			elseif datediff("d", date(),session("pInServiceDate")) > 0 then
 			 	'undo=true
 			 	session("undo")="MissingLate"			
 		
			elseif trim(session("AuthorizedRep"))="" then
 			 	'undo=true			
 			 	session("undo")="MissingRep"
 			
			elseif trim(session("AuthorizedRepTitle"))="" then
 				'undo=true			
 				session("undo")="MissingRepTitle"	
 			
 			else
 				'undo=false 		
 				session("undo")=""
 				if AddPart4bweData() then
 					pNXX=trim(session("pPart4NXX"))
 					
 					log "R",session("pPart4NPA"),session("pPart4NXX"),session("UserUserID"),Now,session("Tix"),"In-Service","In-Service","Part4" 
					
					Tix=GetRsTicket()
					if Tix<>"" then
						log "R",session("ppart4NPA"),session("pPart4NXX"),session("UserUserID"),Now,Tix,"Associated",session("Tix"),"Part4" 
					end if
					
					if session("UserUserEMail")=session("P1UserEmail")	then 'if P1 and P4 entered by the same user
 						EMailTo= session("UserUserEMail") & "," & session("P1EntityEmail") & "," & GetAdminEmail()
 					else	'if P1 and P4 entered by diffent users 
 						EMailTo= session("UserUserEMail") & "," & session("P1EntityEmail") & "," & session("P1UserEmail") & "," & GetAdminEmail()
 					end if
 						
'
' This section was added by G. Brown Feb 8, 2000
'
UserEntityType=session("UserEntityType")
UserUserType=session("UserUserType")
If UserEntityType <> "a" and UserUserType <> "a" then
email session("AdminEntityEMail"),EMailTo,"","CNAS In-Service Confirmation", "Ticket# " & session("Tix")& ", NPA : " & "" & session("ppart4NPA")& ", NXX : " & session("pPart4NXX") & " In-Service confirmed on " & date 
end if
session("EMailTo")=EMailTo
session("Part4Act")="done"
					
					Response.Redirect "xca_Part4Confirm.asp"
					
 				else
					session("Part4Act")=""
					'if not undo then 'retain the previously entered information
						txtInServiceDate.value=session("pInServiceDate")
						txtAuthorizedRep.value=session("AuthorizedRep")
						txtAuthorizedRepTitle.value=session("AuthorizedRepTitle")
						session("undo")="SubmitFail"
				end if
			end if
		end if

Case "YES"
'Elseif P4DATEValue="YES" Then
'Response.Write "This is the YES P4DATEValue->" & P4DATEValue
		if session("Part4Act")="done" then 
		Response.Redirect "xca_Part4Deny.asp"
		else		
	
			session("pPart4Name")=Replace(txtAuthorizedRep.value,"'","''")
			session("pPart4Title")=Replace(txtAuthorizedRepTitle.value,"'","''")
			session("pPart4Date")=Date()
			session("pInServiceDate")=Replace(txtInServiceDate.value,"'","''")
			'session("pSignature")=txtAuthorizedID
			session("AuthorizedRep")=Replace(txtAuthorizedRep.value,"'","''")
			session("AuthorizedRepTitle")=Replace(txtAuthorizedRepTitle.value,"'","''")
			session("Part4Act")="Submit"
			
				if session("pInServiceDate")="" then
	 				session("undo")="MissingDate"
	 				 			
	 			elseif not IsDateReal(session("pInServiceDate")) then
	 				session("undo")="MissingFormat"
	 				
	  			'elseif datediff("d", session("P3EffDate"), session("pInServiceDate")) < 0 then
	 			'	session("undo")="InCorrectDate"
	 			
	 			'elseif datediff("d", session("ReceiptDate"),session("pInServiceDate")) < cint(GetP4DAYS()) then
	 			 	'undo=true	
	 			'session("undo")="MissingEarly"		
	 			
				'elseif datediff("d", date(),session("pInServiceDate")) > 0 then
	 			 	'undo=true
	 			' 	session("undo")="MissingLate"			
	 		
				elseif trim(session("AuthorizedRep"))="" then
	 			 	'undo=true			
	 			 	session("undo")="MissingRep"
	 			
				elseif trim(session("AuthorizedRepTitle"))="" then
	 				'undo=true			
	 				session("undo")="MissingRepTitle"	
	 			
	 			else
	 				'undo=false 		
	 				session("undo")=""
	 				if AddPart4bweData() then
	 					pNXX=trim(session("pPart4NXX"))
	 					
	 					log "R",session("pPart4NPA"),session("pPart4NXX"),session("UserUserID"),Now,session("Tix"),"In-Service","In-Service","Part4" 
						
						Tix=GetRsTicket()
						if Tix<>"" then
							log "R",session("ppart4NPA"),session("pPart4NXX"),session("UserUserID"),Now,Tix,"Associated",session("Tix"),"Part4" 
						end if
						
						if session("UserUserEMail")=session("P1UserEmail")	then 'if P1 and P4 entered by the same user
	 						EMailTo= session("UserUserEMail") & "," & session("P1EntityEmail") & "," & GetAdminEmail()
	 					else	'if P1 and P4 entered by diffent users 
	 						EMailTo= session("UserUserEMail") & "," & session("P1EntityEmail") & "," & session("P1UserEmail") & "," & GetAdminEmail()
	 					end if
	 						
						email session("AdminEntityEMail"),EMailTo,"","CNAS In-Service Confirmation", "Ticket# " & session("Tix")& ", NPA : " & "" & session("ppart4NPA")& ", NXX : " & session("pPart4NXX") & " In-Service confirmed on " & date 
						
						session("EMailTo")=EMailTo
						session("Part4Act")="done"
						
						Response.Redirect "xca_Part4Confirm.asp"
						
	 				else
						session("Part4Act")=""
						'if not undo then 'retain the previously entered information
							txtInServiceDate.value=session("pInServiceDate")
							txtAuthorizedRep.value=session("AuthorizedRep")
							txtAuthorizedRepTitle.value=session("AuthorizedRepTitle")
							session("undo")="SubmitFail"
					end if
				end if
			end if
Case else
'	Response.Write "This is the P4DATEValue->"&ttt
'	Else P4DATEValue=""
'End If
End Select
'''''''''''''''''''''''''''''''''''''''''end of Brian's addition'''''''''''''''''''''''''''''''''''''''''''''''
End Sub


Sub btnReturnToMain_onclick()
	Response.Redirect  "xca_MenuSubPost.asp"
End Sub

'</SCRIPT>

select case session("undo")

case "MissingDate"
	%>						
 		<SCRIPT Language="JavaScript">
		alert("Please enter an In-Service date.")
		</SCRIPT>
    <%
case "MissingFormat"
	%>
 		<SCRIPT Language="JavaScript">
		alert("In-Service date is not in the right format.")
		</SCRIPT>
	<%

case "InCorrectDate"
	%>
		<SCRIPT Language="JavaScript">
		alert("In-Service date must be equal or prior to the current date, and greater than or equal to the NXX Effective Date. If current date is less than the NXX Effective Date, then the Part 4 Confirmation of Code Activation cannot be completed.")
		</SCRIPT>
	<%

case "MissingEarly"
	%>
 		<SCRIPT Language="JavaScript">
		alert("In-Service date too early.")
		</SCRIPT>
	<%

case "MissingLate"
	%>
 		<SCRIPT Language="JavaScript">
		alert("In-Service date must be equal or prior to the current date.")
		</SCRIPT>
	<%
		
case "MissingRep"
	%>
 		<SCRIPT Language="JavaScript">
		alert("Please enter an authorized representative.")
		</SCRIPT>
	<%
			

case "MissingRepTitle"
	%>
 		<SCRIPT Language="JavaScript">
		alert("Please enter the authorized representative title.")
		</SCRIPT>
	<%

case "SubmitFail"
	%>
 		<SCRIPT Language="JavaScript">
		alert("An error has occured while adding the record. Please contact your CNAS administrator.")
		</SCRIPT>
	<%
case ""


		select case session("Part4Act")

		 		case "" 
		 			
					if session("ppart4NPA")<>"" and isnumeric(session("ppart4NPA")) and session("pPart4NXX")<>"" and isnumeric(session("pPart4NXX")) then
						Result = GetNewPart4Data(session("ppart4NPA"),session("pPart4NXX"))
					end if
								 	
		 		case "Submit"

				case "done"	
					
					
				case else
					session("Part4Act")=""
		end select
end select 	
			
			'Response.Redirect "xca_MainMenuPost.asp"
		''''''''''''mike stuff
		AdminData=session("ADMIN")
'		UserEntityID =cint(session("P4EntityID"))
'		UserEntityID =int(session("UserEntityID"))
		AdminUserID=session("UserUserID")
'Response.Write UserEntityID
'Response.Write AdminUserID

'If I use the session variable for UserID(session("UserUserID")), I get the current user.  If I try using just the 
'column name, I get the error message: "Error converting data type varchar to numeric. Using int(UserID) 
'eliminates the error, but gives blank data.
'sql3 = "SELECT * FROM xca_User, xca_Entity, xca_COCode WHERE xca_COCode.Tix = '"&session("Tix")&"' AND xca_COCode.NPA = '"&session("pPart4NPA")&"' AND xca_COCode.NXX = '"&session("pPart4NXX")&"' AND xca_User.UserID = '"&int(UserID)&"' AND xca_Entity.EntityID = '"&UserEntityID&"'"
'sql3 = "Select * From xca_COCode, xca_User Where xca_COCode.EntityID = '"&UserEntityID&"' and xca_COCode.Tix = '"&session("Tix")&"' and xca_COCode.NPA = '"&session("pPart4NPA")&"' and xca_COCode.NXX = '"&session("pPart4NXX")&"'"


'Simplified sql statements.
sql3 = "SELECT * FROM xca_Part3, xca_COCode WHERE xca_Part3.Tix = '"&session("Tix")&"' AND xca_COCode.NPA = '"&session("pPart4NPA")&"' AND xca_COCode.NXX = '"&session("pPart4NXX")&"'"
'sql3 = "Select * From xca_Part1 Where EntityID = '"&UserEntityID&"' and UserID = '"&session("UserUserID")&"'"
	GetPart3UserData.setSQLText(sql3)
	GetPart3UserData.Open
	p3EntityID=GetPart3UserData.fields.getValue("EntityID")
	'session("P1UserID")=GetPart1UserData.fields.getValue("UserID")
	'pTix=GetPart1UserData.fields.getValue("Tix")
	'Response.Write "Tix is=>" & pTix
	'Response.Write "|--|UserID is=>" & session("P1UserID")
	'Response.Write "This is p3EntityID" & p3EntityID
	'P1UID=session("P1UserID")


'get Admin info for top of form
sqlADMIN="Select * from xca_Entity where EntityName ='"&AdminData&"'"
	GetAdminEntityName.setSQLText(sqlADMIN)
	GetAdminEntityName.Open

sqlEnt="Select EntityName From xca_Entity Where xca_Entity.EntityID = '"&p3EntityID&"'"
	GetEntityName.setSQLText(sqlEnt)
	GetEntityName.open
	EntName=GetEntityName.fields.getValue("EntityName")
	'Response.Write "|--|EntityName is=>" & EntName


		'get app info for top of form
'sql = "Select * from xca_Entity where EntityID = '"&UserEntityID&"'"
sql = "Select * from xca_Entity, xca_User, xca_Part3 Where xca_User.UserID = '"&session("UserUserID")&"' and xca_Part3.EntityID = '"&p3EntityID&"' and xca_Entity.EntityName = '"&EntName&"' and xca_Part3.AssignedNPA = '"&session("pPart4NPA")&"' and xca_Part3.AssignedNXX = '"&session("pPart4NXX")&"'"
	GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open
	pEntityName=GetUserEntityName.fields.getValue("EntityName")
	'Response.Write "|--|EntityID is=>" & p3EntityID
	'Response.Write "|--|NPA is=>" & session("pPart4NPA")
	'Response.Write "|--|NXX is=>" & session("pPart4NXX")
	'Response.Write "|--|EntityName is=>" & pEntityName
	
''''''''''''''''''
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sEntityName\sFrom\sxca_Entity\sWhere\sEntityID\s=\s?\q,TCControlID_Unmatched=\qGetEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sEntityName\sFrom\sxca_Entity\sWhere\sEntityID\s=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetEntityName()
{
}
function _initGetEntityName()
{
	GetEntityName.advise(RS_ONBEFOREOPEN, _setParametersGetEntityName);
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
	cmdTmp.CommandText = 'Select EntityName From xca_Entity Where EntityID = ?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetEntityName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetEntityName') != null)
		GetEntityName.setBookmark(thisPage.getState('pb_GetEntityName'));
}
function _GetEntityName_ctor()
{
	CreateRecordset('GetEntityName', _initGetEntityName, null);
}
function _GetEntityName_dtor()
{
	GetEntityName._preserveState();
	thisPage.setState('pb_GetEntityName', GetEntityName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 id=GetUserEntityName 
	style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part3\sWHERE\s(xca_Entity.EntityID\s=\s?)\sAND\s(xca_User.UserID\s=\s?)\sAND\s(xca_Part3.AssignedNPA\s=\s?)\sAND\s(xca_Part3.AssignedNXX\s=\s?)\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Entity,\sxca_User,\sxca_Part3\sWHERE\s(xca_Entity.EntityID\s=\s?)\sAND\s(xca_User.UserID\s=\s?)\sAND\s(xca_Part3.AssignedNPA\s=\s?)\sAND\s(xca_Part3.AssignedNXX\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=4,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1),Row3=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam3\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=0),Row4=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam4\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=0)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Entity, xca_User, xca_Part3 WHERE (xca_Entity.EntityID = ?) AND (xca_User.UserID = ?) AND (xca_Part3.AssignedNPA = ?) AND (xca_Part3.AssignedNXX = ?)';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetParmsData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sValue\sFrom\sxca_Parms\swhere\sName='P4DATE\q,TCControlID_Unmatched=\qGetParmsData\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sValue\sFrom\sxca_Parms\swhere\sName='P4DATE\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetParmsData()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select Value From xca_Parms where Name=\'P4DATE';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetParmsData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetParmsData') != null)
		GetParmsData.setBookmark(thisPage.getState('pb_GetParmsData'));
}
function _GetParmsData_ctor()
{
	CreateRecordset('GetParmsData', _initGetParmsData, null);
}
function _GetParmsData_dtor()
{
	GetParmsData._preserveState();
	thisPage.setState('pb_GetParmsData', GetParmsData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart3UserData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sxca_Part3,\sxca_COCode\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_COCode.NPA\s=\s?)\sAND\s(xca_COCode.NXX\s=\s?)\q,TCControlID_Unmatched=\qGetPart3UserData\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sxca_Part3,\sxca_COCode\sWHERE\s(xca_Part3.Tix\s=\s?)\sAND\s(xca_COCode.NPA\s=\s?)\sAND\s(xca_COCode.NXX\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=3,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qNumeric\q,CSize_Unmatched=\q19\q,CReq=1),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1),Row3=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam3\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersGetPart3UserData()
{
}
function _initGetPart3UserData()
{
	GetPart3UserData.advise(RS_ONBEFOREOPEN, _setParametersGetPart3UserData);
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
	cmdTmp.CommandText = 'SELECT * FROM xca_Part3, xca_COCode WHERE (xca_Part3.Tix = ?) AND (xca_COCode.NPA = ?) AND (xca_COCode.NXX = ?)';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart3UserData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPart3UserData') != null)
		GetPart3UserData.setBookmark(thisPage.getState('pb_GetPart3UserData'));
}
function _GetPart3UserData_ctor()
{
	CreateRecordset('GetPart3UserData', _initGetPart3UserData, null);
}
function _GetPart3UserData_dtor()
{
	GetPart3UserData._preserveState();
	thisPage.setState('pb_GetPart3UserData', GetPart3UserData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminEntityName()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_Entity';
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
		
 			
<BODY leftmargin=20 rightmargin=20 bgColor=#d7c7a4>	

<table align="left" border="0" cellPadding="0" cellSpacing="0" width="804" background ="" height="108" style  ="HEIGHT: 108px; WIDTH: 804px">
    <tr>
        <td align="left"><strong><font face="Arial" color="maroon" size="4">
	CNA 
            Ticket #:&nbsp;&nbsp;<strong><font face="Arial" color="maroon" size="5">
            <% Response.Write session("Tix") %></font></strong></font></strong></td></TD>
	</tr>
	<tr>
	</tr>
	<tr>
		<td wrap align=middle><font color=maroon face="Arial Black" size="5"><strong>
		Part 4: Confirmation of Code Activation 
    </strong></font>
		</td>
	</tr>
</table>

<P>&nbsp;<P>
<P>&nbsp;<P>
<P>&nbsp;<P>
<P>&nbsp;<P>

<table align="left" border="0" cellPadding="0" cellSpacing="0">
	<tr>
		<td wrap><strong><font size="2" face=arial><strong>Code Applicants are required to retain a copy of all 
            application forms, appendices and supporting data in the event of an 
            audit.</strong></font></strong> 
		</td>
	</tr> 
    <TR>
        <TD style="FONT-WEIGHT: bold">&nbsp;
	<tr>
        <td wrap style="FONT-WEIGHT: bold"><label><strong><font size="3" face=Arial color="#993300">
        Contact Information:</font></strong></label> 
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
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code 
            Applicant Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" wrap><font face="Arial"> </font>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Company</STRONG></STRONG></font></font></font><STRONG><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
 </font></font></STRONG> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityname 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityname">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Entity Name</STRONG> 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityName 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityName">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityContact">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Contact Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityAddress()
{
	AppEntityAddress.setDataSource(GetUserEntityName);
	AppEntityAddress.setDataField('UserAddress');
}
function _AppEntityAddress_ctor()
{
	CreateLabel('AppEntityAddress', _initAppEntityAddress, null);
}
</script>
<% AppEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Street 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityCity">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityCity()
{
	AppEntityCity.setDataSource(GetUserEntityName);
	AppEntityCity.setDataField('UserCity');
}
function _AppEntityCity_ctor()
{
	CreateLabel('AppEntityCity', _initAppEntityCity, null);
}
</script>
<% AppEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>City</STRONG> 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityProvince()
{
	AppEntityProvince.setDataSource(GetUserEntityName);
	AppEntityProvince.setDataField('UserProvince');
}
function _AppEntityProvince_ctor()
{
	CreateLabel('AppEntityProvince', _initAppEntityProvince, null);
}
</script>
<% AppEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Province</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityPostalCode()
{
	AppEntityPostalCode.setDataSource(GetUserEntityName);
	AppEntityPostalCode.setDataField('UserPostalCode');
}
function _AppEntityPostalCode_ctor()
{
	CreateLabel('AppEntityPostalCode', _initAppEntityPostalCode, null);
}
</script>
<% AppEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Postal Code</STRONG> 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 59px" width=59>
	<PARAM NAME="_ExtentX" VALUE="1561">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>E-Mail 
            Address</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 48px" width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityFax">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Facsimile</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 52px" width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Telephone</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 84px" width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4"><STRONG>Extension</STRONG></STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" COLOR="Maroon"><B>
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

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 align=left>
    <TBODY>
    
    <TR>
        <TD colSpan=3><font size="3" face=arial color="#993300"><strong>
            1.1 CO Code Information:</strong></font>
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;
	<TR>
		<TD><STRONG><FONT face=Arial size=3>1)</FONT>&nbsp;&nbsp;</STRONG></TD>
		<TD><STRONG><FONT face=Arial size=2>CO Code:</FONT></STRONG></TD>
		<TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write session("ppart4NPA")%>
             -
            <%response.write session("pPart4NXX")%></font></STRONG>
		</TD>
	</TR>
	<TR>
		<TD valign=top><STRONG><FONT face=Arial size=3>2)</FONT></STRONG></TD>
		<TD><STRONG><FONT face=Arial size=2>Switch Identification <BR>(Switching Entity/POI):</FONT></STRONG></TD>
		<TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write session("pSwitchID")%></font></STRONG>
		</TD>
	</TR>
	<TR>
		<TD><STRONG><FONT face=Arial size=3>3)</FONT></STRONG></TD>
		<TD><STRONG><FONT face=Arial size=2>Part 1 Application Date:</FONT></STRONG></TD>
		<TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write session("pApplicationDate")%></font></STRONG>
		</TD>
	</TR>
    <TR>
        <TD>
        <TD><STRONG><FONT face=Arial size=2>Part 1 Date of Receipt:</FONT></STRONG>
          <TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%=session("ReceiptDate")%></font></STRONG>
		</td>
	</tr>
    <TR>
        <TD>
        <TD><STRONG><FONT face=Arial size=2>Part 3 Date of Receipt:</FONT></STRONG>
        <TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write session("P3DateofReceipt")%></font></STRONG>
		</td>
	</tr>
    <TR>
        <TD>
        <TD><STRONG><FONT face=Arial size=2>Part 4 Date of Receipt:</FONT></STRONG>
        <TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write Date()%></font></STRONG>
		</td>
	</tr>
    <TR>
        <TD>
        <TD><STRONG><FONT face=Arial size=2>NXX Effective Date:</FONT></STRONG>
        <TD align=left><font face=Arial size=3 color=maroon><STRONG>&nbsp;
            <%response.write session("P3EffDate")%></font></STRONG>
		</td>
	</tr>
    <TR>
        <TD>&nbsp;
        <TD>&nbsp;&nbsp;&nbsp;&nbsp;
        <TD align=left>
        </TD>
	</TR>
</TBODY>
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

<TABLE ALIGN=left BORDER=0 CELLSPACING=0 CELLPADDING=0>
    <TR>
        <TD colSpan=2><font size="3" face=arial color="#993300"><STRONG>1.2 Confirmation 
            Information:</STRONG></font>
    <TR>
        <TD colSpan=2>&nbsp; 
    <TR>
        <TD colSpan=2><P><FONT face=arial size=2><strong>By submiting a Part 4, I certify 
            that the CO Code(NXX) specified below is in service and that the CO 
            Code (NXX) is being used for purpose specified in the original 
            application (See Section 6.3.3).</FONT></P></STRONG>
    <TR>
        <TD colSpan=2>&nbsp;
    <TR>
        <TD><STRONG><FONT face=Arial size=2>In-Service Date:
            <%Response.write "(Date must be between the NXX Effective Date and the current date)" %>
             </FONT></STRONG>
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtInServiceDate 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtInServiceDate">
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
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtInServiceDate()
{
	txtInServiceDate.setStyle(TXT_TEXTBOX);
	txtInServiceDate.setMaxLength(10);
	txtInServiceDate.setColumnCount(10);
}
function _txtInServiceDate_ctor()
{
	CreateTextbox('txtInServiceDate', _inittxtInServiceDate, null);
}
</script>
<% txtInServiceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		<font face=Arial size=1><strong>(dd/mm/ccyy)</strong></font>
		</td>
    <TR>
        <TD>&nbsp; 
        <TD>
	<TR>
		<TD><STRONG><FONT face=Arial size=2>Authorized Representative of Code 
            Applicant:</FONT></STRONG>&nbsp;&nbsp; 
		</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtAuthorizedRep 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtAuthorizedRep">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtAuthorizedRep()
{
	txtAuthorizedRep.setStyle(TXT_TEXTBOX);
	txtAuthorizedRep.setMaxLength(35);
	txtAuthorizedRep.setColumnCount(35);
}
function _txtAuthorizedRep_ctor()
{
	CreateTextbox('txtAuthorizedRep', _inittxtAuthorizedRep, null);
}
</script>
<% txtAuthorizedRep.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
	</TR>
	<TR>
		<TD><FONT face=Arial size=2><STRONG>Title:</STRONG></FONT> 
		</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtAuthorizedRepTitle 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 590px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtAuthorizedRepTitle">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtAuthorizedRepTitle()
{
	txtAuthorizedRepTitle.setStyle(TXT_TEXTBOX);
	txtAuthorizedRepTitle.setMaxLength(35);
	txtAuthorizedRepTitle.setColumnCount(35);
}
function _txtAuthorizedRepTitle_ctor()
{
	CreateTextbox('txtAuthorizedRepTitle', _inittxtAuthorizedRepTitle, null);
}
</script>
<% txtAuthorizedRepTitle.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
	</TR>
    <TR>
        <TD>&nbsp; 
        <TD>
	<TR>
		<TD><STRONG><FONT face=Arial size=2>Logon that created Part 4:</FONT></STRONG>
		</TD>
		<TD><font size=3 face=Arial color=maroon><STRONG>
            <%response.write session("pSignature")%></font> 
		</TD>
	</TR>
	<TR>
		<TD><STRONG><FONT face=Arial size=2>Date:</FONT></STRONG> 
		</TD>
		<TD><font size=3 face=Arial color=maroon><STRONG>
            <%response.write Date()%></font> 
		</TD>
	</TR>
    <TR>
        <TD>&nbsp; 
        <TD>
    <TR>
        <TD>&nbsp; 
        <TD>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSubmit style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 63px" 
            width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSubmit">
	<PARAM NAME="Caption" VALUE="Submit">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnSubmit()
{
	btnSubmit.value = 'Submit';
	btnSubmit.setStyle(0);
}
function _btnSubmit_ctor()
{
	CreateButton('btnSubmit', _initbtnSubmit, null);
}
</script>
<% btnSubmit.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnReturnToMain 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 636px; WIDTH: 61px" width=61>
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
		<TD align=right>
</TD>
	</TR>
</TABLE></FONT> 


</BODY>

<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>

</HTML>

<%@ Language=VBScript %>

<%
Response.Buffer=true
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

<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

dim RecIndex


function NXXIncludable(NPA,NXX)
	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objCmd1=server.CreateObject("ADODB.Command")
	Set objRec1=server.CreateObject("ADODB.Command")
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	
	'on error resume next
	objCmd1.CommandText="select Status from xca_COCode where NPA='"  & NPA & "' and NXX = '" & NXX & "'"
	set objRec1=objCmd1.Execute
	if not objRec1.EOF then
		if objRec1("Status")="S" then
			NXXIncludableTemp=true
		else
			NXXIncludableTemp=false
		end if
	end if
	objConn1.close
	Set objConn1=Nothing
	Set objCmd1=Nothing
	NXXIncludable=NXXIncludableTemp
end function

Sub btnInclude_onclick()
	if not Rec1.EOF then
		if NXXIncludable(txtNewNPA.value,Rec1.fields.getValue("NXX")) then
			Rec1.fields.setValue "NPASplitID",txtSplitID.value
			log "C",session("pCurrentNPA"),"",session("UserUserID"),Now,0,"Edit",txtSplitID.value,"NPA Split" 
			'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
			'email session("AdminUserEMail"),EMailTo,"","", ""
		else
			session("SplitAct")="NOTIncludable"
			'log "C",session("pCurrentNPA"),"",session("UserUserID"),Now,0,"Edit",txtSplitID.value,"NPA Split" 
			'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
			'email session("AdminUserEMail"),EMailTo,"","", "" 
		end if
		if not Rec1.EOF then
			Rec1.movenext
		end if
	end if	
End Sub

Sub btnExclude_onclick()
	Rec1.fields.setValue "NPASplitID","Excluded"
	log "C",session("pCurrentNPA"),"",session("UserUserID"),Now,0,"Edit",txtSplitID.value,"NPA Split" 
	if not Rec1.EOF then
		Rec1.movenext
	end if
End Sub

sub getSessionParamSplit()
	session("pSplitID")=txtSplitID.value	
	session("pCurrentNPA")=txtCurrentNPA.value	
	session("pNewNPA")=txtNewNPA.value
	session("pPDPStartDate")=txtStartDate.value	
	session("pPDPEndDate")=txtEndDate.value
end sub

Sub btnAddNew_onclick()
	getSessionParamSplit
	session("LastDate")="31/12/9999"
	session("SplitAct")="Add"
	Response.Redirect "xca_Split.asp"
End Sub

Sub btnUpdate_onclick()
	getSessionParamSplit
	session("LastDate")="31/12/9999"
	session("SplitAct")="Update"
	Response.Redirect "xca_Split.asp"
End Sub

Sub btnComplete_onclick()

	getSessionParamSplit
	session("LastDate")="31/12/9999"
	''''''''''''''''''''''''''''''''''
	if (not IsDateReal(txtEndDate.value)) or (not IsDateReal(txtStartDate.value)) then
		session("Error")="Wrong Format in date field"
			
	elseif cdate(txtStartDate.value) > cdate(txtEndDate.value) then
		session("Error")="Wrong order in start date and end date"
			 
	elseif cdate(txtEndDate.value) > date() then
		session("Error")="Wrong end date"
		
	else
		''''''''''''''''''''''''''''''
		Set objConn=server.CreateObject("ADODB.Connection")
		Set objRec=server.CreateObject("ADODB.Recordset")
		Set objCmd=server.CreateObject("ADODB.Command")
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
		
		SQLString=	"SELECT NPA, NXX, TransferID, NPASplitID FROM xca_COCode WHERE (NPA = '" _
					& trim(txtCurrentNPA.value) & "') AND (TransferID <> 'Excluded') AND " _
					& "(NPASplitID = '" & trim(txtSplitID.value) & "')"
		
	
		objCmd.CommandText=SQLString
				
		set objRec=objCmd.Execute
		
		if not objRec.EOF then
	
			session("NPA_S_M")= trim(txtCurrentNPA.value)
			session("SplitID_S_M")=trim(txtSplitID.value)
			Response.Redirect "xca_Split_message.asp"	
					
		end if			
		objRec.close
		objConn.close
		Set objConn=Nothing
		Set objRec=Nothing
		Set objCmd=Nothing
		''''''''''''''''''''''''''''''
	end if	
	''''''''''''''''''''''''''''''''''
	session("SplitAct")="Split"
	Response.Redirect "xca_Split.asp"
End Sub

Sub btnDelete_onclick()
	session("pSplitID")=txtSplitID.value	
	session("pCurrentNPA")=""	
	session("pCurrentNPA_log")=txtCurrentNPA.value	
	session("pNewNPA")=""
	session("pPDPStartDate")=""	
	session("pPDPEndDate")=""
	session("SplitAct")="Delete"
	Response.Redirect "xca_Split.asp"
End Sub

Sub btnReturn_onclick()
	Response.Redirect "xca_SplitAdmin.asp"
End Sub

Sub btnGoto_onclick()
	Found=false
	if Rec1.isOpen() then
		Rec1.moveFirst
		do while (not Rec1.EOF) and (not Found)
			if Rec1.fields.getvalue("NXX")=txtNXX.value then 
				found=true
			else
				Rec1.moveNext
			end if	
		loop
	end if	
End Sub

</SCRIPT>

</HEAD>

<BODY bgColor="#d7c7a4">
<% 

dim NPA
txtSplitID.value=session("pSplitID")
'txtCurrentNPA1.selectByText(session("pCurrentNPA"))
txtCurrentNPA.value=session("pCurrentNPA")
txtNewNPA.value=session("pNewNPA")
txtStartDate.value= session("pPDPStartDate")
txtEndDate.value=session("pPDPEndDate")

dim objConn
dim objCmd
dim objRec

Select Case session("SplitAct")
Case "Add"

	session("Error")=""
	if	txtSplitID.value="" then
		session("Error")="Missing Split ID."
		%>
		<SCRIPT Language="JavaScript">
		alert("Missing Split ID.")
		</SCRIPT>
		<%
	elseif	(txtCurrentNPA.value="") or (txtNewNPA.value="") then
		session("Error")="Missing Current NPA or New NPA value."
		%>
		<SCRIPT Language="JavaScript">
		alert("Missing Current NPA or New NPA value.")
		</SCRIPT>
		<%
		
	elseif (not IsDateReal(txtEndDate.value)) or (not IsDateReal(txtStartDate.value)) then
		session("Error")="Wrong Format in date field(s)."
		%>
		<SCRIPT Language="JavaScript">
		alert("Incorrect format in date field(s).")
		</SCRIPT>
		<%
		
	else 
		if cdate(txtStartDate.value) < date() then
		session("Error")="Wrong order in start date and end date"
		%>
		<SCRIPT Language="JavaScript">
		alert("PDP start date must be after the current date.")
		</SCRIPT>
		<%	
		end if	
		if cdate(txtStartDate.value) > cdate(txtEndDate.value) then
		session("Error")="Wrong order in start date and end date"
		%>
		<SCRIPT Language="JavaScript">
		alert("Start date must be prior to end date.")
		</SCRIPT>
		<%	
		end if
		if txtCurrentNPA.value=txtNewNPA.value then
		session("Error")="Same NPA"
		%>
		<SCRIPT Language="JavaScript">
		alert("Current NPA and New NPA can not be the same.")
		</SCRIPT>
		<%	
		end if
		
	end if
	if session("Error")="" then	

		Set objConn=server.CreateObject("ADODB.Connection")
		Set objRec=server.CreateObject("ADODB.Recordset")
		Set objCmd=server.CreateObject("ADODB.Command")
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn	
		on error resume next
		
		objCmd.CommandText="CheckExistingSplit '" & Replace(trim(txtSplitID.value),"'","''") & "'"
		set objRec=objCmd.Execute
			
		if not objRec.EOF then  %>
			<SCRIPT Language="JavaScript">
			alert("The Split ID is already in the data base. This record can not be added.")
			</SCRIPT>
		<%else
		
			objCmd.CommandText="CheckExistingNPA '" & Replace(trim(txtCurrentNPA.value),"'","''") & "'"
			set objRec=objCmd.Execute
	
			if objRec.EOF then  %>
			
				<SCRIPT Language="JavaScript">
				alert("The specified current NPA does not exist in the database. Split can not be added.")
				</SCRIPT>
			
			<%else			
				objCmd.CommandText="CheckExistingNPA '" & Replace(trim(txtNewNPA.value),"'","''") & "'"
				set objRec=objCmd.Execute
	
				if objRec.EOF then  %>
				<%	objCmd.CommandText=	"CreateNPA '" & trim(txtNewNPA.value) & "', '" & txtStartDate.value  & "'" '& trim(session("pEarliestInServiceDate")) & "'"
					objCmd.Execute %>
							
					<%if objConn.Errors.Count <> 0 then  
						proceed=false
					%>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while creating the NPA.")
						</SCRIPT>
					<%else
						proceed=true
						log "C",trim(txtNewNPA.value),"",session("UserUserID"),Now,0,"New","","NPA Split" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","NPA Created", "NPA  < " & trim(txtNewNPA.value) & " > created on " & date 
					end if%>
				<% else %>	
					<%	proceed=true			%>
				<% end if
				
				if  proceed then
				
					objCmd.CommandText=	"AddSplit '" & Replace(trim(txtSplitID.value),"'","''") & "', '" & session("LastDate") & "','" & trim(txtStartDate.value) _
													& "', '" & trim(txtEndDate.value)& "', '" & Replace(trim(txtCurrentNPA.value),"'","''") _
													& "', '" & Replace(trim(txtNewNPA.value),"'","''")& "', '" & "P" & "'"
					objCmd.Execute %>
							
					<%if objConn.Errors.Count <> 0 then  %>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while adding the split.")
						</SCRIPT>
					<%else
						log "C",trim(txtCurrentNPA.value),"",session("UserUserID"),Now,0,"New",trim(txtSplitID.value),"NPA Split" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","NPA Split Added", "NPA split < " & trim(txtSplitID.value) & " > added on " & date 
					end if%>
				
				<%end if%>
			<%end if%>
				
		<% end if%>
	
		<%objConn.close
		Set objConn=Nothing
		Set objRec=Nothing
		Set objCmd=Nothing
		session("SplitAct")
	else
		session("Error")=""
	end if
	session("SplitAct")=""
	
case "Update"
	
	session("Error")=""
	if	txtSplitID.value="" then
		session("Error")="Missing Split ID."
		%>
		<SCRIPT Language="JavaScript">
		alert("Missing Split ID.")
		</SCRIPT>
		<%
	elseif	(txtCurrentNPA.value="") or (txtNewNPA.value="") then
		session("Error")="Missing Current NPA or New NPA value."
		%>
		<SCRIPT Language="JavaScript">
		alert("Missing Current NPA or New NPA value.")
		</SCRIPT>
		<%
		
	elseif (not IsDateReal(txtEndDate.value)) or (not IsDateReal(txtStartDate.value)) then
		session("Error")="Wrong Format in date field(s)."
		%>
		<SCRIPT Language="JavaScript">
		alert("Incorrect format in date field(s).")
		</SCRIPT>
		<%
		
	else 
		if cdate(txtStartDate.value) < date() then
		session("Error")="Wrong order in start date and end date"
		%>
		<SCRIPT Language="JavaScript">
		alert("PDP start date must be after the current date.")
		</SCRIPT>
		<%	
		end if	
		if cdate(txtStartDate.value) > cdate(txtEndDate.value) then
		session("Error")="Wrong order in start date and end date"
		%>
		<SCRIPT Language="JavaScript">
		alert("Start date must be prior to end date.")
		</SCRIPT>
		<%	
		end if
		if txtCurrentNPA.value=txtNewNPA.value then
		session("Error")="Same NPA"
		%>
		<SCRIPT Language="JavaScript">
		alert("Current NPA and New NPA can not be the same.")
		</SCRIPT>
		<%	
		end if
	end if
	if session("Error")="" then	

		Set objConn=server.CreateObject("ADODB.Connection")
		Set objRec=server.CreateObject("ADODB.Recordset")
		Set objCmd=server.CreateObject("ADODB.Command")
	
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
	
		on error resume next
		objCmd.CommandText="CheckExistingSplit '" & Replace(trim(txtSplitID.value),"'","''") & "'"
		set objRec = objCmd.Execute%>
	
		<%if objRec.EOF then  %>
			<SCRIPT Language="JavaScript">
			alert("The Split ID does not exist in the data base. No record is updated.")
			</SCRIPT>
		<%else
			objCmd.CommandText="CheckExistingNPA '" & Replace(trim(txtCurrentNPA.value),"'","''") & "'"
			set objRec=objCmd.Execute
	
			if objRec.EOF then  %>
			
				<SCRIPT Language="JavaScript">
				alert("The specified current NPA does not exist in the database. Split can not be Updated.")
				</SCRIPT>
			
			<%else			
				objCmd.CommandText="CheckExistingNPA '" & Replace(trim(txtNewNPA.value),"'","''") & "'"
				set objRec=objCmd.Execute
	
				if objRec.EOF then  %>
				<%	objCmd.CommandText=	"CreateNPA '" & trim(txtNewNPA.value) & "', '" & txtStartDate.value  & "'" '& trim(session("pEarliestInServiceDate")) & "'"
					objCmd.Execute %>
							
					<%if objConn.Errors.Count <> 0 then 
					 
						Proceed=false
					%>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while creating the NPA.")
						</SCRIPT>
					<%else
					
						Proceed=true
						log "C",trim(txtNewNPA.value),"",session("UserUserID"),Now,0,"New","","NPA Split" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","NPA Created", "NPA  < " & trim(txtNewNPA.value) & " > created on " & date 
					end if%>
				
				<%else%>
				
					<%Proceed=true%>
									
				<% end if
			
				if Proceed then
				
					objCmd.CommandText=	"UpdateSplit '" & Replace(trim(txtSplitID.value),"'","''") & "', '" & trim(txtStartDate.value) _
												& "', '" & trim(txtEndDate.value)& "', '" & trim(txtCurrentNPA.value) _
												& "', '" & trim(txtNewNPA.value)& "', '" & "P" & "'"
					objCmd.Execute %>
							
					<%if objConn.Errors.Count <> 0 then  %>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while adding the split.")
						</SCRIPT>
					<%else
						log "C",trim(txtCurrentNPA.value),"",session("UserUserID"),Now,0,"Edit",trim(txtSplitID.value),"NPA Split" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","NPA Split Updateded", "NPA split < " & trim(txtSplitID.value) & " > updated on " & date 
					end if
					
				end if		
						
			end if%>
	
		<%end if%>
	
		<%objConn.close
		Set objConn=Nothing
		Set objRec=Nothing
		Set objCmd=Nothing
		'session("SplitAct")=""
	else
		session("Error")=""
	end if
	session("SplitAct")=""
	
case "Split"	
	session("SplitAct")=""
	if session("Error")="Wrong Format in date field" then
		session("Error")=""
		%>
		<SCRIPT Language="JavaScript">
		alert("Incorrect format in date field(s).")
		</SCRIPT>
		<%
		
	elseif session("Error")="Wrong order in start date and end date" then
		session("Error")=""
		%>
		<SCRIPT Language="JavaScript">
		alert("Start date must be prior to end date.")
		</SCRIPT>
		<%	 
	elseif session("Error")="Wrong end date" then
		session("Error")=""
		%>
		<SCRIPT Language="JavaScript">
		alert("PDP End Date must be prior or equal to the current date to complete a split.")
		</SCRIPT>
		<%	
	else

		Set objConn=server.CreateObject("ADODB.Connection")
		Set objCmd=server.CreateObject("ADODB.Command")
	
		on error resume next
		objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
		objCmd.ActiveConnection = objConn
		objCmd.CommandText=	"Split '" & trim(txtCurrentNPA.value) & "', '" & trim(txtNewNPA.value)& "', '" & Replace(trim(txtSplitID.value),"'","''") & "', '" & trim(Session("LastDate"))& "'"
		objCmd.Execute%>

		<%if objConn.Errors.Count <> 0 then  %>
			<SCRIPT Language="JavaScript">
			alert("An error has occured while executing the split.")
			</SCRIPT>
		<%else
			log "C",trim(txtCurrentNPA.value),"",session("UserUserID"),Now,0,"Complete",trim(txtSplitID.value),"NPA Split" 
			'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
			'email session("AdminEntityEMail"),EMailTo,"","NPA Split Completed", "NPA split < " & trim(txtSplitID.value) & " > completed on " & date 
		end if%>
	
		<%objConn.close
		Set objConn=Nothing
		Set objCmd=Nothing
				
	end if
	txtCurrentNPA.value=""
	txtNewNPA.value=""
	txtEndDate.value=""
	txtStartDate.value=""
	txtNXX.value=""
	txtSplitID.value=""
	session("pCurrentNPA")=""
	session("pSplitID")=""

case "Delete"	

'	txtSplitID.value=session("pSplitID")
'	txtCurrentNPA.value=session("pCurrentNPA")
'	txtNewNPA.value=session("pNewNPA")
'	txtStartDate.value= session("pPDPStartDate")
'	txtEndDate.value=session("pPDPEndDate")
	
	txtSplitID.value=""
	on error resume next
	Set objConn=server.CreateObject("ADODB.Connection")
	Set objCmd=server.CreateObject("ADODB.Command")
	
	objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd.ActiveConnection = objConn
	objCmd.CommandText=	"DeleteSplit '" & Replace(trim(session("pSplitID")),"'","''") & "'"
	objCmd.Execute%>
	
	<%if objConn.Errors.Count <> 0 then  %>
		<SCRIPT Language="JavaScript">
		alert("An error has occured while deleting the split.")
		</SCRIPT>
	<%else
		log "C",trim(session("pCurrentNPA_log")),"",session("UserUserID"),Now,0,"Delete",trim(session("pSplitID")),"NPA Split" 
		'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
		'email session("AdminEntityEMail"),EMailTo,"","NPA Split Deleted", "NPA split < " & trim(session("pSplitID")) & " > deleted on " & date 
	end if%>
	
	<%objConn.close
	Set objConn=Nothing
	Set objCmd=Nothing
	session("SplitAct")=""	
	session("pSplitID")=""
	'session("SplitAct")=""
	session("pCurrentNPA_log")=""
case "NOTIncludable"

	%>
		<SCRIPT Language="JavaScript">
		alert("This NXX can not be included in the split because the status of the split-to-NXX is not available.")
		</SCRIPT>
	<%
	session("SplitAct")=""
end select	

dim NPASplitID1
dim NPASplitID2
dim NPASplitID3

NPA=session("pCurrentNPA")
NPASplitID1=session("pSplitID")
'NPASplitID2=""
'NPASplitID3=null
'txtCurrentNPA
Rec1.open
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sv_xca_COCode_Split.NPA,\sv_xca_COCode_Split.NXX,\sxca_status_codes.COStatusDescription,\sv_xca_COCode_Split.Status,\sv_xca_COCode_Split.NPASplitID\sFROM\sv_xca_COCode_Split\sINNER\sJOIN\sxca_status_codes\sON\sv_xca_COCode_Split.Status\s=\sxca_status_codes.COStatus\sWHERE\s(v_xca_COCode_Split.NPA\s=\s?)\sAND\s(v_xca_COCode_Split.NPASplitID\s=\s?\sOR\sv_xca_COCode_Split.NPASplitID\s=\s'Excluded')\sORDER\sBY\sv_xca_COCode_Split.NXX\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sv_xca_COCode_Split.NPA,\sv_xca_COCode_Split.NXX,\sxca_status_codes.COStatusDescription,\sv_xca_COCode_Split.Status,\sv_xca_COCode_Split.NPASplitID\sFROM\sv_xca_COCode_Split\sINNER\sJOIN\sxca_status_codes\sON\sv_xca_COCode_Split.Status\s=\sxca_status_codes.COStatus\sWHERE\s(v_xca_COCode_Split.NPA\s=\s?)\sAND\s(v_xca_COCode_Split.NPASplitID\s=\s?\sOR\sv_xca_COCode_Split.NPASplitID\s=\s'Excluded')\sORDER\sBY\sv_xca_COCode_Split.NXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=2,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1,CValue_Unmatched=\qNPA\q),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q10\q,CReq=0,CValue_Unmatched=\qNPASplitID1\q)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersRec1()
{
	Rec1.setParameter(0,NPA);
	Rec1.setParameter(1,NPASplitID1);
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
	cmdTmp.CommandText = 'SELECT v_xca_COCode_Split.NPA, v_xca_COCode_Split.NXX, xca_status_codes.COStatusDescription, v_xca_COCode_Split.Status, v_xca_COCode_Split.NPASplitID FROM v_xca_COCode_Split INNER JOIN xca_status_codes ON v_xca_COCode_Split.Status = xca_status_codes.COStatus WHERE (v_xca_COCode_Split.NPA = ?) AND (v_xca_COCode_Split.NPASplitID = ? OR v_xca_COCode_Split.NPASplitID = \'Excluded\') ORDER BY v_xca_COCode_Split.NXX';
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


<TABLE WIDTH=75% ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 id=lblTitle 
	style="HEIGHT: 37px; LEFT: 0px; TOP: 0px; WIDTH: 137px" width=137>
	<PARAM NAME="_ExtentX" VALUE="3625">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA Split">
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
	lblTitle.setCaption('NPA Split');
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
</TABLE>

<TABLE WIDTH=84.7% BORDER=0 CELLSPACING=1 CELLPADDING=1 height=75 style="HEIGHT: 75px; WIDTH: 681px" align=center>
	<TR>
		<TD  align=right>
           
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblSplitID 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblSplitID">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA Split ID">
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
function _initlblSplitID()
{
	lblSplitID.setCaption('NPA Split ID');
}
function _lblSplitID_ctor()
{
	CreateLabel('lblSplitID', _initlblSplitID, null);
}
</script>
<% lblSplitID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
<TD>
&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtSplitID 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtSplitID">
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
function _inittxtSplitID()
{
	txtSplitID.setStyle(TXT_TEXTBOX);
	txtSplitID.setMaxLength(10);
	txtSplitID.setColumnCount(10);
}
function _txtSplitID_ctor()
{
	CreateTextbox('txtSplitID', _inittxtSplitID, null);
}
</script>
<% txtSplitID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD  align=right>
           
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnAddNew 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAddNew">
	<PARAM NAME="Caption" VALUE="  Add  ">
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
	btnAddNew.value = '  Add  ';
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
<TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnUpdate 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 67px" width=67>
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
	<TR>
		<TD  align=right>
            
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblCurrentNPA 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 64px" width=64>
	<PARAM NAME="_ExtentX" VALUE="1693">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblCurrentNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Current NPA">
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
function _initlblCurrentNPA()
{
	lblCurrentNPA.setCaption('Current NPA');
}
function _lblCurrentNPA_ctor()
{
	CreateLabel('lblCurrentNPA', _initlblCurrentNPA, null);
}
</script>
<% lblCurrentNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtCurrentNPA 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtCurrentNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtCurrentNPA()
{
	txtCurrentNPA.setStyle(TXT_TEXTBOX);
	txtCurrentNPA.setMaxLength(3);
	txtCurrentNPA.setColumnCount(3);
}
function _txtCurrentNPA_ctor()
{
	CreateTextbox('txtCurrentNPA', _inittxtCurrentNPA, null);
}
</script>
<% txtCurrentNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->           
</TD>
		<TD  align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblNewNPA 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblNewNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="New NPA">
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
function _initlblNewNPA()
{
	lblNewNPA.setCaption('New NPA');
}
function _lblNewNPA_ctor()
{
	CreateLabel('lblNewNPA', _initlblNewNPA, null);
}
</script>
<% lblNewNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>
&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtNewNPA 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNewNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNewNPA()
{
	txtNewNPA.setStyle(TXT_TEXTBOX);
	txtNewNPA.setMaxLength(3);
	txtNewNPA.setColumnCount(3);
}
function _txtNewNPA_ctor()
{
	CreateTextbox('txtNewNPA', _inittxtNewNPA, null);
}
</script>
<% txtNewNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</DIV>
</TD>
	</TR>
	<TR>
		<TD  align=right>
         
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblStartDate 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblStartDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="PDP Start Date">
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
function _initlblStartDate()
{
	lblStartDate.setCaption('PDP Start Date');
}
function _lblStartDate_ctor()
{
	CreateLabel('lblStartDate', _initlblStartDate, null);
}
</script>
<% lblStartDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>
&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtStartDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtStartDate">
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
function _inittxtStartDate()
{
	txtStartDate.setStyle(TXT_TEXTBOX);
	txtStartDate.setMaxLength(10);
	txtStartDate.setColumnCount(10);
}
function _txtStartDate_ctor()
{
	CreateTextbox('txtStartDate', _inittxtStartDate, null);
}
</script>
<% txtStartDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD  align=right>
         
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblEndDate 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 70px" width=70>
	<PARAM NAME="_ExtentX" VALUE="1852">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblEndDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="PDP End Date">
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
function _initlblEndDate()
{
	lblEndDate.setCaption('PDP End Date');
}
function _lblEndDate_ctor()
{
	CreateLabel('lblEndDate', _initlblEndDate, null);
}
</script>
<% lblEndDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>
&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtEndDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEndDate">
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
function _inittxtEndDate()
{
	txtEndDate.setStyle(TXT_TEXTBOX);
	txtEndDate.setMaxLength(10);
	txtEndDate.setColumnCount(10);
}
function _txtEndDate_ctor()
{
	CreateTextbox('txtEndDate', _inittxtEndDate, null);
}
</script>
<% txtEndDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><BR>

<TABLE ALIGN=center border=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
       
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnComplete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 117px" width=117>
	<PARAM NAME="_ExtentX" VALUE="3096">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnComplete">
	<PARAM NAME="Caption" VALUE="Complete Split">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnComplete()
{
	btnComplete.value = 'Complete Split';
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
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE="Delete Split">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnDelete()
{
	btnDelete.value = 'Delete Split';
	btnDelete.setStyle(0);
}
function _btnDelete_ctor()
{
	CreateButton('btnDelete', _initbtnDelete, null);
}
</script>
<% btnDelete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD> <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnInclude 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 66px" width=66>
	<PARAM NAME="_ExtentX" VALUE="1746">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnInclude">
	<PARAM NAME="Caption" VALUE="Include">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnInclude()
{
	btnInclude.value = 'Include';
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
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 71px" width=71>
	<PARAM NAME="_ExtentX" VALUE="1879">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnExclude">
	<PARAM NAME="Caption" VALUE="Exclude">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnExclude()
{
	btnExclude.value = 'Exclude';
	btnExclude.setStyle(0);
}
function _btnExclude_ctor()
{
	CreateButton('btnExclude', _initbtnExclude, null);
}
</script>
<% btnExclude.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoto style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 82px" 
	width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoto">
	<PARAM NAME="Caption" VALUE="Goto NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoto()
{
	btnGoto.value = 'Goto NXX';
	btnGoto.setStyle(0);
}
function _btnGoto_ctor()
{
	CreateButton('btnGoto', _initbtnGoto, null);
}
</script>
<% btnGoto.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtNXX style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" 
	width=18>
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
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD>
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


<BR><TABLE ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 623px" 
	width=623>
	<PARAM NAME="_ExtentX" VALUE="16484">
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
	<PARAM NAME="ColumnsNames" VALUE='"NXX","COStatusDescription","NPASplitID"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="130,153,435">
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
	<PARAM NAME="TitleAlignment" VALUE="0">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="8">
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
	<PARAM NAME="GridWidth" VALUE="623">
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
Grid1.pageSize = 8;
Grid1.setDataSource(Rec1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=623';
Grid1.headerAttributes = '   bgcolor=Maroon align=Left';
Grid1.headerWidth[0] = ' WIDTH=130';
Grid1.headerWidth[1] = ' WIDTH=153';
Grid1.headerWidth[2] = ' WIDTH=435';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NXX\'';
Grid1.colHeader[1] = '\'Status\'';
Grid1.colHeader[2] = '\'Include\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=130';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'NXX\')';
Grid1.colAttributes[1] = '  WIDTH=153';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'COStatusDescription\')';
Grid1.colAttributes[2] = '  WIDTH=435';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'NPASplitID\')';
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

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>

</HTML>

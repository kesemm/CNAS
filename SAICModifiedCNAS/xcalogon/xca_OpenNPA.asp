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
<TITLE>Connection</TITLE>

<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>



function StatusValid(pStatus)

	dim StatusValidTemp
	
	Set objConn1=server.CreateObject("ADODB.Connection")
	Set objRec1=server.CreateObject("ADODB.Recordset")
	Set objCmd1=server.CreateObject("ADODB.Command")
	
	objConn1.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
	objCmd1.ActiveConnection = objConn1
	on error resume next
	objCmd1.CommandText=	"GetStatusCode '" & pStatus &"'"
	set objRec1=objCmd1.Execute
	
	if objRec1.EOF then
		StatusValidTemp=false
	else
		StatusValidTemp=true
	end if
	
	objConn1.close
	Set objConn1=Nothing
	Set objRec1=Nothing
	Set objCmd1=Nothing
	StatusValid=StatusValidTemp
	
end function

Sub btnUpdate_onclick()
	if IsDateReal(txtEarliestInServiceDate.value) then
		
		''txt=lstStatus.selectedIndex
			
		''Rec1.fields.setValue "Status",lstStatus.getValue(txt)
		''Rec1.fields.setValue "Status",txtStatusHidden.value	
		Rec1.fields.setValue "EarliestInServiceDate",txtEarliestInServiceDate.value
		Rec1.updateRecord
			
			
		log "C",trim(session("pNPA_log")),trim(Rec1.fields.getvalue("NXX")),session("UserUserID"),Now,0,"Update","","NPA Maint" 
		'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
		'email session("AdminEntityEMail"),EMailTo,"","NPA-NXX updated", "NPA-NXX  < " & trim(session("pNPA_log"))  & "-" & trim(Rec1.fields.getvalue("NXX")) & " > updateded on " & date 
			
		session("OpenNPAAct")="Retrieve"
		session("pNPA")=txtCurrentNPA.value
			
		Response.Redirect "xca_OpenNPA.asp"
		
	else
		session("OpenNPAAct")="WrongDateFormat"
		
	end if	
End Sub

Sub btnGetCurrent_onclick()
	txtNXX.value=Rec1.fields.getvalue("NXX")
	'lstStatus.selectByText(Rec1.fields.getvalue("COStatusDescription"))
	txtStatus.value=Rec1.fields.getvalue("COStatusDescription")
	'txtStatusHidden.value=Rec1.fields.getvalue("Status")
	txtEarliestInServiceDate.value=Rec1.fields.getvalue("EarliestInServiceDate")
End Sub

Sub btnAdd_onclick()
	session("OpenNPAAct")="Add"
	session("pNPA")=txtNewNPA.value	
	session("pNPA_log")=txtNewNPA.value
	session("pEarliestInServiceDate")=txtEarliestInServiceDate2.value	
	Response.Redirect "xca_OpenNPA.asp"
End Sub

Sub btnDelete_onclick()
	session("OpenNPAAct")="Delete"
	session("pNPA")=txtCurrentNPA.value
	Response.Redirect "xca_OpenNPA.asp"
End Sub

Sub btnRetrieve_onclick()
	session("OpenNPAAct")="Retrieve"
	session("pNPA")=txtCurrentNPA.value
	session("pNPA_log")=txtCurrentNPA.value
	Response.Redirect "xca_OpenNPA.asp"
End Sub

Sub btnReturnToMain_onclick()
	session("pNPA")=""
	session("pNPA_log")=""
	session("OpenNPAAct")="Retrieve"
	'session("pEarliestInServiceDate")=""
	Response.Redirect "xca_MenuC0CAdmin.asp"
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
		if found then 
			txtNXX.value=Rec1.fields.getvalue("NXX")
			txtStatus.value=Rec1.fields.getvalue("COStatusDescription")
			
			'lstStatus.selectByText(Rec1.fields.getvalue("COStatusDescription"))
			txtEarliestInServiceDate.value=Rec1.fields.getvalue("EarliestInServiceDate")
		else
			txtNXX.value=""
			txtStatus.value=""			
			'lstStatus.selectByValue("A") 'reset to the first list item
			txtEarliestInServiceDate.value=""
			Rec1.moveFirst
			session("OpenNPAAct")="GOTO NOT Found"
		end if	
	end if	
End Sub

</SCRIPT>
</HEAD>

<BODY bgColor="#d7c7a4">
<%	
	if isnumeric(session("pNPA")) then
		txtCurrentNPA.value = session("pNPA")
	end if
	
	dim NPA
	dim objConn
	dim objCmd
	dim objRec
	
	Select Case session("OpenNPAAct")
		
		Case "Add"
			if session("pNPA")="" then
				session("OpenNPAAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("Please specify an NPA to add.")
				</SCRIPT>
			<%
			elseif NOT isnumeric(session("pNPA")) then	
				session("OpenNPAAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("NPA must be a numeric data type.")
				</SCRIPT>
			<%
			elseif cLng(session("pNPA"))<200 or cLng(session("pNPA"))>999 then	
				session("OpenNPAAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("NPA is out of the range.")
				</SCRIPT>
			<%
			else
				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")
	
				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
	
				on error resume next
				objCmd.CommandText="CheckExistingNPA '" & trim(session("pNPA")) & "'"
				set objRec=objCmd.Execute%>
	
				<%if not objRec.EOF then  %>
			
					<SCRIPT Language="JavaScript">
					alert("The NPA is already in the data base, it can not be created again.")
					</SCRIPT>
			
				<%else%>
			
						<%if (not IsDateReal(session("pEarliestInServiceDate"))) then %>	
						
							<SCRIPT Language="JavaScript">
							alert("Incorrect format in date field.")
							</SCRIPT>
							
						<%else
							objCmd.CommandText=	"CreateNPA '" & trim(session("pNPA")) _
															& "', '" & trim(session("pEarliestInServiceDate")) & "'"
							objCmd.Execute %>
							<%if objConn.Errors.Count <> 0 then  %>
								<SCRIPT Language="JavaScript">
								alert("An error has occured while creating the NPA. It could be caused by a wrong date format.")
								</SCRIPT>
							<%else
								log "C",trim(session("pNPA_log")),"",session("UserUserID"),Now,0,"New","","NPA Maint" 
								'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
								'email session("AdminEntityEMail"),EMailTo,"","NPA Created", "NPA  < " & trim(session("pNPA_log")) & " > created on " & date 
							end if%>
							
						<%end if%>	
						
				<% end if%>
	
				<%objConn.close
				Set objConn=Nothing
				Set objRec=Nothing
				Set objCmd=Nothing
				session("OpenNPAAct")=""
			end if
			session("OpenNPAAct")=""
			
		case "Delete"
		
			Set objConn=server.CreateObject("ADODB.Connection")
			Set objRec=server.CreateObject("ADODB.Recordset")
			Set objCmd=server.CreateObject("ADODB.Command")
	
			objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
			objCmd.ActiveConnection = objConn
	
			on error resume next
			objCmd.CommandText="CheckExistingNPA '" & trim(session("pNPA")) & "'"
			set objRec=objCmd.Execute%>
	
			<%if objRec.EOF then  %>
				<SCRIPT Language="JavaScript">
				alert("The NPA does NOT exist in the data base, no records deleted.") 
				</SCRIPT>
			<%else
					
					objCmd.CommandText="CheckExistingNPAInUse '" & trim(session("pNPA")) & "'"
					set objRec=objCmd.Execute%>
	
					<%if not objRec.EOF then  %>
						<SCRIPT Language="JavaScript">
						alert("The NPA has NXXs Reserved, Assigned or In-Service , it can not be deleted.") 
						</SCRIPT>
					<%else
						objCmd.CommandText=	"DeleteNPA '" & trim(session("pNPA")) & "'"
						objCmd.Execute %>
				
						<%if objConn.Errors.Count <> 0 then  %>
							<SCRIPT Language="JavaScript">
							alert("An error has occured while creating the NPA.")
							</SCRIPT>
						<%else%>
							<SCRIPT Language="JavaScript">
							alert("The NPA has been deleted.")
							</SCRIPT>
							<%log "C",trim(session("pNPA_log")),"",session("UserUserID"),Now,0,"Delete","","NPA Maint" 
							'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
							'email session("AdminEntityEMail"),EMailTo,"","NPA Deleted", "NPA  < " & trim(session("pNPA_log")) & " > deleted on " & date 
						end if%>
						
					<%end if%>
					
			<% end if%>
	
			<%objConn.close
			Set objConn=Nothing
			Set objRec=Nothing
			Set objCmd=Nothing
			session("OpenNPAAct")=""
			
		case "Retrieve"	
			session("OpenNPAAct")=""
			if isnumeric(session("pNPA")) then
				txtCurrentNPA.value=session("pNPA")
				NPA = session("pNPA")
				Rec1.open
			end if
			
		case "WrongDateFormat"
			session("OpenNPAAct")=""
		 %>
			<SCRIPT Language="JavaScript">
			alert("Incorrect format in date field.")
			</SCRIPT>
		<%
		
		case "InvalidStatus"
			session("OpenNPAAct")=""
		 %>
			<SCRIPT Language="JavaScript">
			alert("Status is invalid.")
			</SCRIPT>
		<%
		case "GOTO NOT Found"
			session("OpenNPAAct")=""
		%>
			<SCRIPT Language="JavaScript">
			alert("The specified goto record can not be found.")
			</SCRIPT>
		<%
		case else
			session("pNPA")=""
	end select
		
	if isnumeric(session("pNPA")) then
		txtCurrentNPA.value=session("pNPA")
		NPA = session("pNPA")
		Rec1.open
	end if
	
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecStatus style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sCOStatus,\sCOStatusDescription\sFROM\sxca_status_codes\sORDER\sBY\sCOStatusDescription\q,TCControlID_Unmatched=\qRecStatus\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sCOStatus,\sCOStatusDescription\sFROM\sxca_status_codes\sORDER\sBY\sCOStatusDescription\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
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
	cmdTmp.CommandText = 'SELECT COStatus, COStatusDescription FROM xca_status_codes ORDER BY COStatusDescription';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_COCode.NPA,\sxca_COCode.NXX,\sxca_status_codes.COStatusDescription,\sxca_COCode.EarliestInServiceDate,\sxca_COCode.Status\sFROM\sxca_COCode\sINNER\sJOIN\sxca_status_codes\sON\sxca_COCode.Status\s=\sxca_status_codes.COStatus\sWHERE\s(xca_COCode.NPA\s=\s?)\sORDER\sBY\sxca_COCode.NXX\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_PreDefined\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_COCode.NPA,\sxca_COCode.NXX,\sxca_status_codes.COStatusDescription,\sxca_COCode.EarliestInServiceDate,\sxca_COCode.Status\sFROM\sxca_COCode\sINNER\sJOIN\sxca_status_codes\sON\sxca_COCode.Status\s=\sxca_status_codes.COStatus\sWHERE\s(xca_COCode.NPA\s=\s?)\sORDER\sBY\sxca_COCode.NXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1,CValue_Unmatched=\qNPA\q)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersRec1()
{
	Rec1.setParameter(0,NPA);
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
	cmdTmp.CommandText = 'SELECT xca_COCode.NPA, xca_COCode.NXX, xca_status_codes.COStatusDescription, xca_COCode.EarliestInServiceDate, xca_COCode.Status FROM xca_COCode INNER JOIN xca_status_codes ON xca_COCode.Status = xca_status_codes.COStatus WHERE (xca_COCode.NPA = ?) ORDER BY xca_COCode.NXX';
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
	style="HEIGHT: 37px; LEFT: 10px; TOP: 192px; WIDTH: 446px" width=446>
	<PARAM NAME="_ExtentX" VALUE="11800">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA And CO Code Maintenance">
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
	lblTitle.setCaption('NPA And CO Code Maintenance');
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
<TABLE ALIGN=center BORDER=1 CELLSPACING=1 CELLPADDING=1 >

<TR><TD>&nbsp;
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
&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtCurrentNPA 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 18px" width=18>
	<PARAM NAME="_ExtentX" VALUE="476">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtCurrentNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="222">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="3">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtCurrentNPA()
{
	txtCurrentNPA.setStyle(TXT_TEXTBOX);
	txtCurrentNPA.setDataField('222');
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
&nbsp;
</TD><TD align=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnRetrieve 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 110px" width=110>
	<PARAM NAME="_ExtentX" VALUE="2910">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnRetrieve">
	<PARAM NAME="Caption" VALUE="Retrieve NPA">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnRetrieve()
{
	btnRetrieve.value = 'Retrieve NPA';
	btnRetrieve.setStyle(0);
}
function _btnRetrieve_ctor()
{
	CreateButton('btnRetrieve', _initbtnRetrieve, null);
}
</script>
<% btnRetrieve.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD align=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnDelete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE=" Delete NPA ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnDelete()
{
	btnDelete.value = ' Delete NPA ';
	btnDelete.setStyle(0);
}
function _btnDelete_ctor()
{
	CreateButton('btnDelete', _initbtnDelete, null);
}
</script>
<% btnDelete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD align=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMain 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 319px; WIDTH: 61px" width=61>
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


<TABLE ALIGN=center border=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
	<TD>&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblNewNPA 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 105px" width=105>
	<PARAM NAME="_ExtentX" VALUE="2778">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblNewNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Open New NPA Code">
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
	lblNewNPA.setCaption('Open New NPA Code');
}
function _lblNewNPA_ctor()
{
	CreateLabel('lblNewNPA', _initlblNewNPA, null);
}
</script>
<% lblNewNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
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
&nbsp;
</TD><TD>&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lblEarliestInServiceDate 
	style="HEIGHT: 17px; LEFT: 10px; TOP: 376px; WIDTH: 73px" width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblEarliestInServiceDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Effective Date">
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
function _initlblEarliestInServiceDate()
{
	lblEarliestInServiceDate.setCaption('Effective Date');
}
function _lblEarliestInServiceDate_ctor()
{
	CreateLabel('lblEarliestInServiceDate', _initlblEarliestInServiceDate, null);
}
</script>
<% lblEarliestInServiceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txtEarliestInServiceDate2 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEarliestInServiceDate2">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEarliestInServiceDate2()
{
	txtEarliestInServiceDate2.setStyle(TXT_TEXTBOX);
	txtEarliestInServiceDate2.setMaxLength(20);
	txtEarliestInServiceDate2.setColumnCount(20);
}
function _txtEarliestInServiceDate2_ctor()
{
	CreateTextbox('txtEarliestInServiceDate2', _inittxtEarliestInServiceDate2, null);
}
</script>
<% txtEarliestInServiceDate2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;
</TD>
		<TD>
            
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnAdd style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 113px" 
	width=113>
	<PARAM NAME="_ExtentX" VALUE="2990">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAdd">
	<PARAM NAME="Caption" VALUE="Add New NPA">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnAdd()
{
	btnAdd.value = 'Add New NPA';
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
</TR>
</TABLE>

<TABLE ALIGN=center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR><TD align=center>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoto style="HEIGHT: 27px; LEFT: 10px; TOP: 439px; WIDTH: 49px" 
	width=49>
	<PARAM NAME="_ExtentX" VALUE="1296">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoto">
	<PARAM NAME="Caption" VALUE="Goto">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoto()
{
	btnGoto.value = 'Goto';
	btnGoto.setStyle(0);
}
function _btnGoto_ctor()
{
	CreateButton('btnGoto', _initbtnGoto, null);
}
</script>
<% btnGoto.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 NXX 
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

</TD>
		<TD>&nbsp;Status
            

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txtStatus style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtStatus">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="30">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtStatus()
{
	txtStatus.setStyle(TXT_TEXTBOX);
	txtStatus.disabled = true;
	txtStatus.setMaxLength(30);
	txtStatus.setColumnCount(20);
}
function _txtStatus_ctor()
{
	CreateTextbox('txtStatus', _inittxtStatus, null);
}
</script>
<% txtStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

&nbsp;
</TD>
		<TD>&nbsp;Effective Date
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtEarliestInServiceDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEarliestInServiceDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="12">
	<PARAM NAME="DisplayWidth" VALUE="12">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEarliestInServiceDate()
{
	txtEarliestInServiceDate.setStyle(TXT_TEXTBOX);
	txtEarliestInServiceDate.setMaxLength(12);
	txtEarliestInServiceDate.setColumnCount(12);
}
function _txtEarliestInServiceDate_ctor()
{
	CreateTextbox('txtEarliestInServiceDate', _inittxtEarliestInServiceDate, null);
}
</script>
<% txtEarliestInServiceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;
</TD><TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGetCurrent 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 529px; WIDTH: 73px" width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGetCurrent">
	<PARAM NAME="Caption" VALUE="Get NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGetCurrent()
{
	btnGetCurrent.value = 'Get NXX';
	btnGetCurrent.setStyle(0);
}
function _btnGetCurrent_ctor()
{
	CreateButton('btnGetCurrent', _initbtnGetCurrent, null);
}
</script>
<% btnGetCurrent.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD><TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnUpdate 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 556px; WIDTH: 100px" width=100>
	<PARAM NAME="_ExtentX" VALUE="2646">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnUpdate()
{
	btnUpdate.value = 'Update NXX';
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

<TABLE ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 631px" 
	width=631>
	<PARAM NAME="_ExtentX" VALUE="16695">
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
	<PARAM NAME="ColumnsNames" VALUE='"NXX","COStatusDescription","EarliestInServiceDate"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="106,230,286">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NXX","Status","Effective Date"'>
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
	<PARAM NAME="GridWidth" VALUE="631">
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
Grid1.pageSize = 10;
Grid1.setDataSource(Rec1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=631';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=106';
Grid1.headerWidth[1] = ' WIDTH=230';
Grid1.headerWidth[2] = ' WIDTH=286';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NXX\'';
Grid1.colHeader[1] = '\'Status\'';
Grid1.colHeader[2] = '\'Effective Date\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=106';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'NXX\')';
Grid1.colAttributes[1] = '  WIDTH=230';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'COStatusDescription\')';
Grid1.colAttributes[2] = '  WIDTH=286';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'EarliestInServiceDate\')';
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

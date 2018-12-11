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
				


Sub btnUpdate_onclick()
	'Response.Write "000"
	if txtNXX.value="" then
		session("PreDefinedAct")="UpdateNoNXX"
		Response.Redirect "xca_PreDefined.asp"	
	elseif trim(Rec1.fields.getValue("PreNXX")) <> trim(txtNXX.value) then
		session("PreDefinedAct")="NXXNotMatch"
		Response.Redirect "xca_PreDefined.asp"				
	else	'if StatusValid(txtStatus.value) then
		'Rec1.fields.setValue "PreNXX",txtNXX.value
		'Response.Write "111"
		session("PreDefinedAct")="Update"
		session("pNXX")=txtNXX.value
	
		txt=lstStatus.selectedIndex
		txt=lstStatus.getValue(txt)
	
		session("pStatus")=txt
	
		session("pDescription")=txtDescription.value
		Response.Redirect "xca_PreDefined.asp"
		
		
		'txt=lstStatus.selectedIndex
		'txt=lstStatus.getValue(txt)
		
		'Rec1.fields.setValue "PreStatus",txt
		
		'Rec1.fields.setValue "PreDescription",txtDescription.value
		
		'log "C","",trim(txtNXX.value),session("UserUserID"),Now,0,"Update","","Predefined" 
		'Response.Redirect "xca_PreDefined.asp"	
	end if	

End Sub

Sub btnAdd_onclick()
	
	session("PreDefinedAct")="Add"
	session("pNXX")=txtNXX.value
	
	txt=lstStatus.selectedIndex
	txt=lstStatus.getValue(txt)
	
	session("pStatus")=txt
	
	session("pDescription")=txtDescription.value
	Response.Redirect "xca_PreDefined.asp"

End Sub

Sub btnDelete_onclick()
	'log inspection
	rec1.deleteRecord
	log "C","",trim(txtNXX.value),session("UserUserID"),Now,0,"Delete","","Predefined" 
	'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
	'email session("AdminEntityEMail"),EMailTo,"","Predefined NXX Deleted", "Predefined NXX  < " & trim(txtNXX.value) & " > deleted on " & date 
						
End Sub

Sub btnGetCurrent_onclick()
	txtNXX.value=Rec1.fields.getvalue("PreNXX")
	'txtStatus.value=Rec1.fields.getvalue("PreStatus")
	lstStatus.selectByText(Rec1.fields.getvalue("COStatusDescription"))
	txtDescription.value=Rec1.fields.getvalue("PreDescription")
End Sub

Sub btnReturnToMain_onclick()
	Response.Redirect "xca_MenuC0CAdmin.asp"
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor="#d7c7a4">
<%
	
	Select Case session("PreDefinedAct")	
		
		Case "Add"
			Set objConn=server.CreateObject("ADODB.Connection")
			Set objRec=server.CreateObject("ADODB.Recordset")
			Set objCmd=server.CreateObject("ADODB.Command")
	
			objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
			objCmd.ActiveConnection = objConn
	
			on error resume next
			objCmd.CommandText="CheckExistingPreNXX '" & Replace(trim(session("pNXX")),"'","''") & "'"
			set objRec=objCmd.Execute%>
	
			<%if not objRec.EOF then  %>
				<SCRIPT Language="JavaScript">
				alert("The Predefined NXX is already in the data base. It can not be added again.")
				</SCRIPT>
			<%else
				if session("pNXX")="" then
				%>
					<SCRIPT Language="JavaScript">
					alert("NXX field must be entered to create a record.")
					</SCRIPT>
				<%
				elseif Not isnumeric(trim(session("pNXX"))) then
				%>	<SCRIPT Language="JavaScript">
					alert("NXX field must be a number.")
					</SCRIPT>
				<%
				elseif cInt(trim(session("pNXX"))) < 200 or cInt(trim(session("pNXX"))) > 999 then
				%>	<SCRIPT Language="JavaScript">
					alert("NXX number is out of range.")
					</SCRIPT>
				<%else
					pStatus=trim(session("pStatus"))
					
					objCmd.CommandText=	"AddPreNXX '" & trim(session("pNXX")) _
													& "', '" & trim(session("pStatus")) _
													& "', '"& Replace(session("pDescription"),"'","''")& "'"
					objCmd.Execute %>
				
					<%if objConn.Errors.Count <> 0 then  %>
						<SCRIPT Language="JavaScript">
						alert("An error has occured while adding the Predefined NXX.")
						</SCRIPT>
					<%else
						log "C","",trim(session("pNXX")),session("UserUserID"),Now,0,"Add","","Predefined" 
						'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
						'email session("AdminEntityEMail"),EMailTo,"","Predefined NXX added", "Predefined NXX  < " & trim(session("pNXX")) & " > added on " & date 
					end if
					%>
				<%end if%>
			<% end if%>
	
			<%objConn.close
			Set objConn=Nothing
			Set objRec=Nothing
			Set objCmd=Nothing
			
			session("PreDefinedAct")=""
			session("pNXX")=""
			session("pStatus")=""
			session("pDescription")=""
			Rec1.requery
			
		Case "Update"
			
			Set objConn=server.CreateObject("ADODB.Connection")
			Set objRec=server.CreateObject("ADODB.Recordset")
			Set objCmd=server.CreateObject("ADODB.Command")
	
			objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
			objCmd.ActiveConnection = objConn
	
			on error resume next
			objCmd.CommandText="CheckExistingPreNXX '" & Replace(trim(session("pNXX")),"'","''") & "'"
			set objRec=objCmd.Execute%>
	
			<%if objRec.EOF then  %>
				<SCRIPT Language="JavaScript">
				alert("The Predefined NXX is Not in the data base. No record is updated.")
				</SCRIPT>
			<%else
				
				pStatus=trim(session("pStatus"))
				
				session("pDescription")=Replace(session("pDescription"))
				
				objCmd.CommandText=	"Update xca_PreDefined set PreStatus= '"  & trim(session("pStatus")) _
												& "', PreDescription = '"   & session("pDescription") _
												& "' where PreNXX = '" & trim(session("pNXX")) _
												& "'"
				
				objCmd.Execute 
				
				%>
				<%if objConn.Errors.Count <> 0 then  %>
					<SCRIPT Language="JavaScript">
					alert("An error has occured while adding the Predefined NXX.")
					</SCRIPT>
				<%else
					log "C","",trim(session("pNXX")),session("UserUserID"),Now,0,"Add","","Predefined" 
					'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
					'email session("AdminEntityEMail"),EMailTo,"","Predefined NXX added", "Predefined NXX  < " & trim(session("pNXX")) & " > added on " & date 
				end if
				%>
				
			<% end if%>
	
			<%objConn.close
			
			Set objConn=Nothing
			Set objRec=Nothing
			Set objCmd=Nothing
			
			session("PreDefinedAct")=""
			session("pNXX")=""
			session("pStatus")=""
			session("pDescription")=""
			Rec1.requery
				
		case "UpdateStatusInvalid"	
			session("PreDefinedAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("Status Code is invalid.")
				</SCRIPT>
			<%	
			
		case "UpdateNoNXX"
			session("PreDefinedAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("NXX field must be entered when updating a record.")
				</SCRIPT>
			<%		
		case "NXXNotMatch"	
			session("PreDefinedAct")=""
			%>
				<SCRIPT Language="JavaScript">
				alert("NXX field can not be changed when updating a record.")
				</SCRIPT>
			<%				
		case else	
			'Rec1.requery
	end select 	
	
%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sa.PreNXX,b.COStatusDescription,a.PreDescription,a.PreStatus\r\nFROM\sxca_PreDefined\sa,\sxca_Status_Codes\sb\r\nwhere\sa.Prestatus=b.COStatus\r\nORDER\sBY\sa.PreNXX\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_PreDefined\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sa.PreNXX,b.COStatusDescription,a.PreDescription,a.PreStatus\r\nFROM\sxca_PreDefined\sa,\sxca_Status_Codes\sb\r\nwhere\sa.Prestatus=b.COStatus\r\nORDER\sBY\sa.PreNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
	cmdTmp.CommandText = 'SELECT a.PreNXX,b.COStatusDescription,a.PreDescription,a.PreStatus FROM xca_PreDefined a, xca_Status_Codes b where a.Prestatus=b.COStatus ORDER BY a.PreNXX';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RecStatus style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sCOStatus,COStatusDescription\sfrom\sxca_status_codes\swhere\sCOStatus\sin\s('P',\s'T',\s'U')\sorder\sby\sCOStatus\q,TCControlID_Unmatched=\qRecStatus\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_status_codes\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sCOStatus,COStatusDescription\sfrom\sxca_status_codes\swhere\sCOStatus\sin\s('P',\s'T',\s'U')\sorder\sby\sCOStatus\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
	cmdTmp.CommandText = 'select COStatus,COStatusDescription from xca_status_codes where COStatus in (\'P\', \'T\', \'U\') order by COStatus';
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
<TABLE WIDTH=75% ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 id=lblTitle 
	style="HEIGHT: 37px; LEFT: 10px; TOP: 192px; WIDTH: 527px" width=527>
	<PARAM NAME="_ExtentX" VALUE="13944">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Predefined Values for New NPA Code">
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
	lblTitle.setCaption('Predefined Values for New NPA Code');
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

<BR>

<TABLE ALIGN=center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
		<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGetCurrent 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 126px" width=126>
	<PARAM NAME="_ExtentX" VALUE="3334">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGetCurrent">
	<PARAM NAME="Caption" VALUE="Get Current NXX">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGetCurrent()
{
	btnGetCurrent.value = 'Get Current NXX';
	btnGetCurrent.setStyle(0);
}
function _btnGetCurrent_ctor()
{
	CreateButton('btnGetCurrent', _initbtnGetCurrent, null);
}
</script>
<% btnGetCurrent.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btngoto style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 107px" 
	width=107>
	<PARAM NAME="_ExtentX" VALUE="2831">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btngoto">
	<PARAM NAME="Caption" VALUE="Get Record #">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtngoto()
{
	btngoto.value = 'Get Record #';
	btngoto.setStyle(0);
	btngoto.hide();
}
function _btngoto_ctor()
{
	CreateButton('btngoto', _initbtngoto, null);
}
</script>
<% btngoto.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtNum style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 24px" 
	width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNum">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNum()
{
	txtNum.setStyle(TXT_TEXTBOX);
	txtNum.hide();
	txtNum.setMaxLength(4);
	txtNum.setColumnCount(4);
}
function _txtNum_ctor()
{
	CreateTextbox('txtNum', _inittxtNum, null);
}
</script>
<% txtNum.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</TD>
		<TD>&nbsp;NXX&nbsp;&nbsp;
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
		<TD>&nbsp;&nbsp;Status&nbsp;&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=lstStatus 
	style="HEIGHT: 21px; LEFT: 10px; TOP: 321px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
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
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstStatus()
{
	RecStatus.advise(RS_ONDATASETCOMPLETE, 'lstStatus.setRowSource(RecStatus, \'COStatusDescription\', \'COStatus\');');
}
function _lstStatus_ctor()
{
	CreateListbox('lstStatus', _initlstStatus, null);
}
</script>
<% lstStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD>&nbsp;&nbsp; Description&nbsp;&nbsp;
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txtDescription 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtDescription">
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
function _inittxtDescription()
{
	txtDescription.setStyle(TXT_TEXTBOX);
	txtDescription.setMaxLength(12);
	txtDescription.setColumnCount(12);
}
function _txtDescription_ctor()
{
	CreateTextbox('txtDescription', _inittxtDescription, null);
}
</script>
<% txtDescription.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>

<BR>
<TABLE ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnAdd style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 78px" 
	width=78>
	<PARAM NAME="_ExtentX" VALUE="2064">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAdd">
	<PARAM NAME="Caption" VALUE="Add New">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnUpdate 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 120px" width=120>
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update Current">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnDelete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 115px" width=115>
	<PARAM NAME="_ExtentX" VALUE="3043">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE="Delete Current">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMain 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 442px; WIDTH: 61px" width=61>
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

<BR><BR><TABLE ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 695px" 
	width=695>
	<PARAM NAME="_ExtentX" VALUE="18389">
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
	<PARAM NAME="ColumnsNames" VALUE='"PreNXX","COStatusDescription","PreDescription"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="177,0,319">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NXX","Status","Description"'>
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
	<PARAM NAME="GridWidth" VALUE="695">
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
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=695';
Grid1.headerAttributes = '   bgcolor=Maroon align=Left';
Grid1.headerWidth[0] = ' WIDTH=177';
Grid1.headerWidth[1] = ' WIDTH=0';
Grid1.headerWidth[2] = ' WIDTH=319';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NXX\'';
Grid1.colHeader[1] = '\'Status\'';
Grid1.colHeader[2] = '\'Description\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=177';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'PreNXX\')';
Grid1.colAttributes[1] = '  WIDTH=0';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'COStatusDescription\')';
Grid1.colAttributes[2] = '  WIDTH=319';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'PreDescription\')';
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
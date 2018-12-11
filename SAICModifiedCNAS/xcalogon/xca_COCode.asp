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
	Session("LastDate")="31/12/9999"
	if txtNPA.value<>"" and txtNXX.value<>"" then
		if txtInServiceDate.value="" then
			txtInServiceDate.value=Session("LastDate")
		end if	

		if not IsDateReal(txtEarliestInServiceDate.value) then
			session("Error")="Wrong Date Format"
		elseif not IsDateReal(txtInServiceDate.value) then
			session("Error")="Wrong Date Format"	
		elseif cdate(txtEarliestInServiceDate.value) > cdate(txtInServiceDate.value) then
			session("Error")="Wrong Date order"
		elseif not StatusValid(txtStatus.value) then
			session("Error")="InvalidStatus"
		elseif session("sParamNPA")<>txtNPA.value or session("sParamNXX")<>txtNXX.value	then
			session("Error")="NPANXXChanged"
			
		end if	
		'session("sParamNPA")=txtNPA.value
		'session("sParamNXX")=txtNXX.value
		
		session("spStatus")=txtStatus.value
		
		if txtTix.value="" then
			session("spTix")="0"
		elseif not isnumeric(txtTix.value) then
			session("Error")="Number Wrong"
		else	
			session("spTix")=txtTix.value
		end if	
				
		if txtEntityID.value="" then
			session("spEntityID")="0"
		elseif not isnumeric(txtEntityID.value) then
			session("Error")="Number Wrong"	
		else
			session("spEntityID")=txtEntityID.value
		end if	
		
		session("spLATA")=txtLATA.value
		session("spOCN")=txtOCN.value
		session("spSwitchID")=txtSwitchID.value
		session("spWireCenter")=txtWireCenter.value
		session("spRateCenter")=txtRateCenter.value
		session("spEarliestInServiceDate")=txtEarliestInServiceDate.value
		if txtInServiceDate.value="" then
			session("spInServiceDate")=Session("LastDate")
			
		else	
			session("spInServiceDate")=txtInServiceDate.value
		end if	
		session("spNPASplitID")=txtNPASplitID.value
		session("spTransferID")=txtTransferID.value
		session("COCodeAct")="Update"
		session("COCodeAct1")="Update" 'update subdevided into "Update" and "Clear" as is reruired for log table
		'Response.Redirect "xca_COCode.asp"
	end if
End Sub

Sub btnGetRecord_onclick()
	if txtNPA.value<>"" and txtNXX.value<>"" then
		session("COCodeAct")="GetRec" 
		'session("pNPA")=txtNPA.value
		'session("pNXX")=txtNXX.value
		session("sParamNPA")=txtNPA.value
		session("sParamNXX")=txtNXX.value
		'Response.Redirect "xca_COCode.asp"
	end if	
End Sub

Sub btnClearRecord_onclick()
	Session("LastDate")="31/12/9999"
	if txtNPA.value<>"" and txtNXX.value<>"" then
		if session("sParamNPA")<>txtNPA.value or session("sParamNXX")<>txtNXX.value	then
			session("Error")="NPANXXChanged"
		end if
		session("spStatus")="S"
		session("spTix")=0
		session("spEntityID")=0
		session("spLATA")=""
		session("spOCN")=""
		session("spSwitchID")=""
		session("spWireCenter")=""
		session("spRateCenter")=""
		session("spEarliestInServiceDate")=txtEarliestInServiceDate.value
		session("spInServiceDate")=Session("LastDate")
		session("spNPASplitID")="Excluded"
		session("spTransferID")="Excluded"
		session("COCodeAct")="Update"
		session("COCodeAct1")="Clear"
		'Response.Redirect "xca_COCode.asp"
	end if
End Sub

Sub btnClearScreen_onclick()
	'session("COCodeAct")="GetRec" 
	session("sParamNPA")=""
	session("sParamNXX")=""
	txtNPA.value=""
	txtNXX.value=""
				
	txtStatus.value=""
	txtTix.value=""
	txtEntityID.value=""
	txtLATA.value=""
	txtOCN.value=""
	txtSwitchID.value=""
	txtWireCenter.value=""
	txtRateCenter.value=""
	txtEarliestInServiceDate.value=""
	txtInServiceDate.value=""
	txtNPASplitID.value=""
	txtTransferID.value=""
	Response.Redirect "xca_COCode.asp"
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

	objCmd.CommandText="SelectCOCode '" & trim(ParamNPA) & "', '" & trim(ParamNXX) & "'"
			
	set objRec=objCmd.Execute
	GetRecordTemp=false
	if not objRec.EOF then
				
		txtNPA.value=checkNull(objRec("NPA"))
		txtNXX.value=checkNull(objRec("NXX"))
				
		txtStatus.value=checkNull(objRec("Status"))
		txtTix.value=checkNull(objRec("Tix"))
		txtEntityID.value=checkNull(objRec("EntityID"))
		txtLATA.value=checkNull(objRec("LATA"))
		txtOCN.value=checkNull(objRec("OCN"))
		txtSwitchID.value=checkNull(objRec("SwitchID"))
		txtWireCenter.value=checkNull(objRec("WireCenter"))
		txtRateCenter.value=checkNull(objRec("RateCenter"))
		txtEarliestInServiceDate.value=checkNull(objRec("EarliestInServiceDate"))
		txtInServiceDate.value=checkNull(objRec("InServiceDate"))
		txtNPASplitID.value=checkNull(objRec("NPASplitID"))
		txtTransferID.value=checkNull(objRec("TransferID"))
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
	
	Response.Redirect "xca_MenuCNASadmin.asp"
End Sub

</SCRIPT>
</HEAD>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%	select case session("COCodeAct")

 		case "GetRec" 
 		
 			ParamNPA=session("sParamNPA")
			ParamNXX=session("sParamNXX")
			if ParamNPA="" or ParamNXX="" then
				'btnClearScreen_onclick
 				'do nothing
			elseif (not isnumeric(ParamNPA)) or (not isnumeric(ParamNXX)) then%>
				<SCRIPT Language="VBSCRIPT">
				alert("Both NPA and NXX must be numeric data type.")
				</SCRIPT>

<% 
 			elseif not GetRecord(ParamNPA,ParamNXX)then%>
				<SCRIPT Language="VBSCRIPT">
				alert("No record exists for the specified NPA and NXX.")
				</SCRIPT>

<% 
 			end if
			 	
 		case "Update" 
 		
 			ParamNPA=session("sParamNPA")
			ParamNXX=session("sParamNXX")
			
			%>

<%select case session("Error")%>

<%	case "Wrong Date order" %>
				<SCRIPT Language="VBSCRIPT">
				alert("Incorrect order of date field(s) entered.")
				</SCRIPT>

<%	case "Wrong Date Format" %>
				<SCRIPT Language="VBSCRIPT">
				alert("Incorrect format of date field(s) entered.")
				</SCRIPT>

<%	case "InvalidStatus"%>
				<SCRIPT Language="VBSCRIPT">
				alert("Status is invalid.")
				</SCRIPT>

<%	case "NPANXXChanged"%>	
				<SCRIPT Language="VBSCRIPT">
				alert("NPA-NXX value has been changed on the screen. Operation cancelled.")
				</SCRIPT>

<%	case "Number Wrong"%>	
				<SCRIPT Language="VBSCRIPT">
				alert("Ticket# and Entity ID must be integer data type.")
				</SCRIPT>

<%	case else	
			
 				pStatus=session("spStatus")
				pTix=session("spTix")
				pEntityID=session("spEntityID")
				pLATA=session("spLATA")
				pOCN=session("spOCN")
				pSwitchID=session("spSwitchID")
				pWireCenter=session("spWireCenter")
				pRateCenter=session("spRateCenter")
				pEarliestInServiceDate=session("spEarliestInServiceDate")
				pInServiceDate=session("spInServiceDate")
				pNPASplitID=session("spNPASplitID")
				pTransferID=session("spTransferID")

				Set objConn=server.CreateObject("ADODB.Connection")
				Set objRec=server.CreateObject("ADODB.Recordset")
				Set objCmd=server.CreateObject("ADODB.Command")

				objConn.Open Application("cnasadmin_ConnectionString"), Application("cnasadmin_RuntimeUserName"), Application("cnasadmin_RuntimePassword")
				objCmd.ActiveConnection = objConn
				objCmd.CommandText=	"UpdateCOCode '" & trim(ParamNPA) _
													& "', '" & trim(ParamNXX) _
													& "', '" & Replace(trim(pStatus),"'","''") _
													& "', '" & Replace(trim(pTix),"'","''") _
													& "', " & trim(pEntityID) _
													& ", '" & Replace(trim(pLATA),"'","''") _
													& "', '" & Replace(trim(pOCN),"'","''") _
													& "', '" & Replace(trim(pSwitchID),"'","''") _
													& "', '" & Replace(trim(pWireCenter),"'","''") _
													& "', '" & Replace(trim(pRateCenter),"'","''") _
													& "', '" & Replace(trim(pEarliestInServiceDate),"'","''") _
													& "', '" & Replace(trim(pInServiceDate),"'","''") _
													& "', '" & Replace(trim(pNPASplitID),"'","''") _
													& "', '" & Replace(trim(pTransferID),"'","''") _
													& "'"
				objCmd.Execute
				objConn.close
				Set objCmd=Nothing
				
				Set objConn=Nothing
				Set objRec=Nothing
				
				log "C",trim(ParamNPA),trim(ParamNXX),session("UserUserID"),Now,0,session("COCodeAct1"),"","CO Maint" 
				'EMailTo= session("UserUserEMail") & "," & session("EntityUserEMail")
				'email session("AdminEntityEMail"),EMailTo,"","CO Code Update", "CO Code  < " & trim(ParamNPA) & "-" & trim(ParamNXX)& " > updated on " & date 
				%>
				<SCRIPT Language="VBSCRIPT">
				alert("The record has been updated successfully.")
				</SCRIPT>

<%
				
			end select	
			
			session("Error")=""
			
			call GetRecord(ParamNPA,ParamNXX)
			
		case else 
			
	end select
	session("COCodeAct")=""
%>

<TABLE WIDTH=75% ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 
            id=lblTitle style="HEIGHT: 37px; LEFT: 37px; TOP: 0px; WIDTH: 485px" 
            width=485>
	<PARAM NAME="_ExtentX" VALUE="12832">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="CO Codes Data Base Maintenance">
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
	lblTitle.setCaption('CO Codes Data Base Maintenance');
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


<TABLE border=1 cellPadding=2 cellSpacing=2 cols=2 align=center>
<TR>
<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnGetRecord 
            style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 109px" width=109>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnUpdate style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 108px" 
            width=108>
	<PARAM NAME="_ExtentX" VALUE="2858">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="  Update DB  ">
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
	btnUpdate.value = '  Update DB  ';
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
            id=btnClearScreen 
            style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 106px" width=106>
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
            id=btnClearRecord 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 146px; WIDTH: 112px" 
width=112>
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnClearRecord">
	<PARAM NAME="Caption" VALUE="Reset Record">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnClearRecord()
{
	btnClearRecord.value = 'Reset Record';
	btnClearRecord.setStyle(0);
}
function _btnClearRecord_ctor()
{
	CreateButton('btnClearRecord', _initbtnClearRecord, null);
}
</script>
<% btnClearRecord.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnReturnToMain 
            style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 119px" width=119>
	<PARAM NAME="_ExtentX" VALUE="3149">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMain">
	<PARAM NAME="Caption" VALUE="Return to  Main">
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
	btnReturnToMain.value = 'Return to  Main';
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
<p>
			<!-- <TABLE border=1 cellPadding=1 cellSpacing=1  cellPadding=2 cellSpacing=2 width=400 style="WIDTH: 400px" 
           >-->
                <TABLE border=0 cellPadding=2 cellSpacing=2 cols=2 align=center>
                <TR>
                    <TD align=left>
            <P>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblNPA style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 66px" 
            width=66>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</P>
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtNPA style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNPA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNPA()
{
	txtNPA.setStyle(TXT_TEXTBOX);
	txtNPA.setMaxLength(3);
	txtNPA.setColumnCount(30);
}
function _txtNPA_ctor()
{
	CreateTextbox('txtNPA', _inittxtNPA, null);
}
</script>
<% txtNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblNXX style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 23px" 
            width=23>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtNXX style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNXX">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="3">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNXX()
{
	txtNXX.setStyle(TXT_TEXTBOX);
	txtNXX.setMaxLength(3);
	txtNXX.setColumnCount(30);
}
function _txtNXX_ctor()
{
	CreateTextbox('txtNXX', _inittxtNXX, null);
}
</script>
<% txtNXX.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
</TR>
    <TR><TD></TD><TD></TD><TR>
</TR>
    <TR><TD></TD><TD></TD><TR>
</TR>
    <TR><TD></TD><TD></TD><TR>
</TR>
    <TR><TD></TD><TD></TD><TR>
<TD>
            <!--METADATA TYPE="DesignerControl" startspan
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtStatus style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtStatus">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="1">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtStatus()
{
	txtStatus.setStyle(TXT_TEXTBOX);
	txtStatus.setMaxLength(1);
	txtStatus.setColumnCount(30);
}
function _txtStatus_ctor()
{
	CreateTextbox('txtStatus', _inittxtStatus, null);
}
</script>
<% txtStatus.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblTicket style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 40px" 
            width=40>
	<PARAM NAME="_ExtentX" VALUE="1058">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblTicket">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Ticket#">
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
function _initlblTicket()
{
	lblTicket.setCaption('Ticket#');
}
function _lblTicket_ctor()
{
	CreateLabel('lblTicket', _initlblTicket, null);
}
</script>
<% lblTicket.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtTix style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtTix">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="6">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtTix()
{
	txtTix.setStyle(TXT_TEXTBOX);
	txtTix.setMaxLength(6);
	txtTix.setColumnCount(30);
}
function _txtTix_ctor()
{
	CreateTextbox('txtTix', _inittxtTix, null);
}
</script>
<% txtTix.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblEntityID 
            style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 46px" width=46>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtEntityID 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEntityID">
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
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEntityID()
{
	txtEntityID.setStyle(TXT_TEXTBOX);
	txtEntityID.setMaxLength(35);
	txtEntityID.setColumnCount(30);
}
function _txtEntityID_ctor()
{
	CreateTextbox('txtEntityID', _inittxtEntityID, null);
}
</script>
<% txtEntityID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtLATA style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtLATA">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="5">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtLATA()
{
	txtLATA.setStyle(TXT_TEXTBOX);
	txtLATA.setMaxLength(5);
	txtLATA.setColumnCount(30);
}
function _txtLATA_ctor()
{
	CreateTextbox('txtLATA', _inittxtLATA, null);
}
</script>
<% txtLATA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblOCN style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 26px" 
            width=26>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtOCN style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtOCN">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtOCN()
{
	txtOCN.setStyle(TXT_TEXTBOX);
	txtOCN.setMaxLength(4);
	txtOCN.setColumnCount(30);
}
function _txtOCN_ctor()
{
	CreateTextbox('txtOCN', _inittxtOCN, null);
}
</script>
<% txtOCN.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtSwitchID 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtSwitchID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="11">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtSwitchID()
{
	txtSwitchID.setStyle(TXT_TEXTBOX);
	txtSwitchID.setMaxLength(11);
	txtSwitchID.setColumnCount(30);
}
function _txtSwitchID_ctor()
{
	CreateTextbox('txtSwitchID', _inittxtSwitchID, null);
}
</script>
<% txtSwitchID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblTerminating 
            style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblTerminating">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Terminating SE/POI">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTerminating()
{
	lblTerminating.hide();
	lblTerminating.setCaption('Terminating SE/POI');
}
function _lblTerminating_ctor()
{
	CreateLabel('lblTerminating', _initlblTerminating, null);
}
</script>
<% lblTerminating.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblWireCenter 
            style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 62px" width=62>
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtWireCenter 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtWireCenter">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtWireCenter()
{
	txtWireCenter.setStyle(TXT_TEXTBOX);
	txtWireCenter.setMaxLength(10);
	txtWireCenter.setColumnCount(30);
}
function _txtWireCenter_ctor()
{
	CreateTextbox('txtWireCenter', _inittxtWireCenter, null);
}
</script>
<% txtWireCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblRateCenter 
            style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 62px" width=62>
	<PARAM NAME="_ExtentX" VALUE="1640">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblRateCenter">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Rate Center">
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
function _initlblRateCenter()
{
	lblRateCenter.setCaption('Rate Center');
}
function _lblRateCenter_ctor()
{
	CreateLabel('lblRateCenter', _initlblRateCenter, null);
}
</script>
<% lblRateCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtRateCenter 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtRateCenter">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtRateCenter()
{
	txtRateCenter.setStyle(TXT_TEXTBOX);
	txtRateCenter.setMaxLength(10);
	txtRateCenter.setColumnCount(30);
}
function _txtRateCenter_ctor()
{
	CreateTextbox('txtRateCenter', _inittxtRateCenter, null);
}
</script>
<% txtRateCenter.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>

                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblEarliestInServiceDate 
            style="HEIGHT: 17px; LEFT: 10px; TOP: 668px; WIDTH: 73px" width=73>
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
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
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
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtEarliestInServiceDate 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEarliestInServiceDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEarliestInServiceDate()
{
	txtEarliestInServiceDate.setStyle(TXT_TEXTBOX);
	txtEarliestInServiceDate.setMaxLength(10);
	txtEarliestInServiceDate.setColumnCount(30);
}
function _txtEarliestInServiceDate_ctor()
{
	CreateTextbox('txtEarliestInServiceDate', _inittxtEarliestInServiceDate, null);
}
</script>
<% txtEarliestInServiceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblInserviceDate 
            style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 78px" width=78>
	<PARAM NAME="_ExtentX" VALUE="2064">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblInserviceDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="In Service Date">
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
function _initlblInserviceDate()
{
	lblInserviceDate.setCaption('In Service Date');
}
function _lblInserviceDate_ctor()
{
	CreateLabel('lblInserviceDate', _initlblInserviceDate, null);
}
</script>
<% lblInserviceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtInServiceDate 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtInServiceDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtInServiceDate()
{
	txtInServiceDate.setStyle(TXT_TEXTBOX);
	txtInServiceDate.setMaxLength(10);
	txtInServiceDate.setColumnCount(30);
}
function _txtInServiceDate_ctor()
{
	CreateTextbox('txtInServiceDate', _inittxtInServiceDate, null);
}
</script>
<% txtInServiceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblTransferID 
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
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
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
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtTransferID 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtTransferID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtTransferID()
{
	txtTransferID.setStyle(TXT_TEXTBOX);
	txtTransferID.setMaxLength(10);
	txtTransferID.setColumnCount(30);
}
function _txtTransferID_ctor()
{
	CreateTextbox('txtTransferID', _inittxtTransferID, null);
}
</script>
<% txtTransferID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR>
                <TR>
                    <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 
            id=lblSplitID style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
            width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lblSplitID">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Split ID">
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
function _initlblSplitID()
{
	lblSplitID.setCaption('Split ID');
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
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtNPASplitID 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtNPASplitID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtNPASplitID()
{
	txtNPASplitID.setStyle(TXT_TEXTBOX);
	txtNPASplitID.setMaxLength(10);
	txtNPASplitID.setColumnCount(30);
}
function _txtNPASplitID_ctor()
{
	CreateTextbox('txtNPASplitID', _inittxtNPASplitID, null);
}
</script>
<% txtNPASplitID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</TD></TR></TABLE></p><BR><BR>

<p>

<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>


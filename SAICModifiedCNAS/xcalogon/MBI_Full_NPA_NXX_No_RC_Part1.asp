<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<!--#include file="xca_CNASLib.inc"-->
<form action="MBI_Full_NPA_NXX_No_RC_Post_Part1.asp" method="post" id="MBI_Full_NPA_NXX_Input" name="MBI_Full_NPA_NXX_Input" onSubmit="return validateForm()">
<html>
<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<title>Part 1 - MBI (Full Based on NPA and NXX with No Rate Centre)</title>
<script LANGUAGE="JavaScript"> <!--

       
    function checkdate(a) {

				var err=0,result
				if (a.length != 10) err=1
					d = a.substring(0, 2)//day  was-> b = a.substring(0, 2)// day
					c = a.substring(2, 3)// '/'
					b = a.substring(3, 5)//month was->d = a.substring(3, 5)// month
					e = a.substring(5, 6)// '/'
					f = a.substring(6, 10)// year
				if (b<1 || b>12) err = 1
				if (c != '/') err = 1
				if (d<1 || d>31) err = 1
				if (e != '/') err = 1
				if (f<1999) err = 1
				if (b==4 || b==6 || b==9 || b==11){
				if (d==31) err=1
				}
				if (b==2){
				var g=parseInt(f/4)
				if (isNaN(g)) {
				err=1
				}
				if (d>29) err=1
				if (d==29 && ((f/4)!=parseInt(f/4))) err=1
				}
				if (err==1) {
				return false;
				}
				else {
					return true;
			   }
		}  

 function validateForm()
        {    

           if (document.MBI_Full_NPA_NXX_Input.AuthorizedRep.value == "") {
                alert("You have not filled in the Authorized Rep field. Please type in an Authorized Name and submit again");
                document.MBI_Full_NPA_NXX_Input.AuthorizedRep.focus();               
                return false;
            }
            if (document.MBI_Full_NPA_NXX_Input.AuthorizedRepTitle.value == "") {
                alert("You have not filled in the Authorized Rep Title field. Please type in an Authorized Name Title and submit again");
                document.MBI_Full_NPA_NXX_Input.AuthorizedRepTitle.focus();
                return false;
            }
			if (document.MBI_Full_NPA_NXX_Input.ApplicationDate.value == "") {
                alert("You have not filled in the Application Date field. Please type in a valid date and submit again");
                document.MBI_Full_NPA_NXX_Input.ApplicationDate.focus();
                return false;
            }
            var result=checkdate(document.MBI_Full_NPA_NXX_Input.ApplicationDate.value) //this one             
            if (result==false)	{
				alert("The Application Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
                document.MBI_Full_NPA_NXX_Input.ApplicationDate.focus();
                return false;
			}
            if (document.MBI_Full_NPA_NXX_Input.OCN.value == "") {
                alert("You have not filled in the OCN field. Please type in a 4 characters and submit again");
                document.MBI_Full_NPA_NXX_Input.OCN.focus();
                return false;
            }
           if (document.MBI_Full_NPA_NXX_Input.OCN.value.length <4) {
                alert("The OCN field must be 4 characters. Please type in a 4 characters and submit again");
                document.MBI_Full_NPA_NXX_Input.OCN.focus();               
                return false; 
            }
			
		}
        // end hiding -->
// app-b    


</script>
 <meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%
SelectedNPA=session("aNPA")
SelectedNXX=session("aNXX")
UserEntityID=session("UserEntityID")
uname = session("UserUserName")
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuMBI.asp"
End Sub

sqluser = "Select * from xca_Entity,xca_User where xca_Entity.EntityID = '"&UserEntityID&"' and xca_User.UserName= '"&uname&"' "
	GetUserEntityName.setSQLText(sqluser)
	GetUserEntityName.Open
	
SET objConnectionCompany = server.createobject("ADODB.connection")
SET rstCompany =server.createobject("ADODB.recordset")
objConnectionCompany.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlCompany = "SELECT * From xca_Entity Inner Join xca_COCode On xca_Entity.EntityID=xca_COCode.EntityID WHERE xca_COCode.NPA='" & SelectedNPA &"' AND xca_COCode.NXX='" & SelectedNXX & "';"
SET rstCompany = objConnectionCompany.execute(sqlCompany)
	%>

</head>
<FORM>
<body leftmargin="20" rightmargin="20" bgColor="#d7c7a4" text="black" LANGUAGE=javascript>
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityName()
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
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetSelectedEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetSelectedEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetSelectedEntityName()
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
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetSelectedEntityName.setRecordSource(rsTmp);
}
function _GetSelectedEntityName_ctor()
{
	CreateRecordset('GetSelectedEntityName', _initGetSelectedEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<table border="0" cellpadding="0"><tr>
	<td wrap><font color="maroon" face="Arial Black" size="4"><strong>
Part 1 - Request for Full Block Assignment</strong></font>
            </td></tr>
</table>

<table align="center" border="0" cellPadding="0" cellSpacing="0">

<tr>
<td align="right"><font face="arial" size="2"><strong>Authorized Representative Name:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><INPUT id=AuthorizedRep name=AuthorizedRep size=40 maxlength=35></td>
</tr>

<tr>
<td align="right"><font face="arial" size="2"><strong>Title:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><INPUT id=AuthorizedRepTitle name=AuthorizedRepTitle size=40 maxlength=35></td>
</tr>

<tr>
<td align="right"><font face="arial" size="2"><strong>Application Date:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><INPUT id=ApplicationDate name=ApplicationDate size=12 maxlength=10></td>
</tr>
<tr><td></td><td align="left" wrap><font face="arial" size="1">dd/mm/ccyy</font></td></tr>

     
<tr>
<td align="right"><font face="arial" size="2"><strong>OCN:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><INPUT id=OCN name=OCN size=6 maxlength=4></td>
</tr>
<tr>
<td align="right"><font face="arial" size="2"><strong>Selected NPA:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><%= SelectedNPA %></td>
</tr>
<tr>
<td align="right"><font face="arial" size="2"><strong>Selected NXX:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><%= SelectedNXX %></td>
</tr>
<tr>
<td align="right"><font face="arial" size="2"><strong>CO Code Assigned To:&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><%= rstCompany("EntityName") %> <%= rstCompany("OCN") %></td>
</tr>
<tr>
<td align="right"><font face="arial" size="2"><strong>Supporting explanation (if desired):&nbsp;&nbsp;</strong></font></td>
<td align="left" wrap><INPUT id=SupportingExplanation name=SupportingExplanation size=50 maxlength=75></td>
</tr>
</table>
<table align="left" border="0" cellPadding="0" cellSpacing="1">
	<tr>
		<td><strong><font size="3" face="arial" color="#993300">1.1 MBI Applicant or Assignee Information:</font></strong></td>
	</tr>
</table>
<br><br>

<table align="left" border="0" cellPadding="1" cellSpacing="1" Col="2">
    <tr>
        <td align="center"><strong><u>Code Applicant Info:</u></strong></td>
    </tr>
    <tr> 
        <td>Entity Name:</td>
        <td align="left"><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
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
			<PARAM NAME="FontColor" VALUE="">
			<PARAM NAME="FontBold" VALUE="-1">
			<PARAM NAME="FontItalic" VALUE="0">
			<PARAM NAME="Visible" VALUE="-1">
			<PARAM NAME="FormatAsHTML" VALUE="0">
			<PARAM NAME="Platform" VALUE="0">
			<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
			-->
			<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial"><B>
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
</td>
</tr>

    <tr>
        <td>Contact Name:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>Street Address:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>City:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>Province:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>Postal Code:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>E-Mail Address: </td> 
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>Facsimile:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
        
    <tr>
        <td>Telephone:</td>
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
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
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
        </td>
		</tr>
</table>
<BR><BR>
<BR><BR>
<BR><BR>
<BR><BR>
<BR><BR>
<BR><BR>
<BR><BR>


    <tr>

<td align = "left" colSpan = "3"> <input type="submit" value="Submit" name="submit">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoToMainFrm 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoToMainFrm">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoToMainFrm()
{
	btnGoToMainFrm.value = 'Return';
	btnGoToMainFrm.setStyle(0);
}
function _btnGoToMainFrm_ctor()
{
	CreateButton('btnGoToMainFrm', _initbtnGoToMainFrm, null);
}
</script>
<% btnGoToMainFrm.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<tr>
<td align = "left" colSpan = "3" wrap>
	</td></tr></TABLE></FORM>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

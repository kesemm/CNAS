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
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>CNAS Administration Menu</title>
<script LANGUAGE="JavaScript"><!--


 function go(url) {
     location.href = url;
 }
 

	
--></script>
<script ID="serverEventHandlersVBS"
LANGUAGE="vbscript" RUNAT="Server">

Sub btnReturntoMain_onclick()
Response.Redirect "xca_MenuMainPost.asp"
End Sub


</script>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
</form>
<!--#include file="xca_CNASLib.inc"-->

<form action="xca_MenuInt.asp" method="post" id="formP4" name="formAdminMenu"
ONLOAD="history.go(0)&quot;">
</form>
<%
session("NoTixSent")=""


session("HereP3")="xca_MenuCNASadmin.asp"
	%>

<p align="center"><font face="Arial Black" color="maroon" size="5"><strong>CO Code
Administration Menu</strong></font> </p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="35%">
  <tr>
    <td width="25"><img height="36" src="ball25.gif" width="35"> </td>
    <td><font face="Arial"><a href="javascript:go('xca_OpenNPA.asp')">NPA and CO Code
    Maintenance</a></font></td>
  </tr>
  <tr>
    <td><img height="36" src="ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_COCode.asp')">CO Codes Database
    Maintenance</a></font></td>
  </tr>
  <tr>
    <td><img height="36" src="ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_TransferAdmin.asp')">CO CodeTransfers</a></font></td>
  </tr>
  <tr>
    <td><img height="36" src="ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_PreDefined.asp')">Predefined CO Code</a></font>
    </td>
  </tr>
  <tr>
    <td><img height="36" src="ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_SplitAdmin.asp')">NPA Splits</a></font></td>
  </tr>
  <tr>
    <td><img height="36" src="ball25.gif" width="35"></td>
    <td><a HREF="javascript:go('xca_WebNPAFileMenu.htm')"><font face="Arial">NPA Files
    Generation Menu</font></a></td>
  </tr>
  <tr>
    <td></td>
    <td>&nbsp; <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnReturntoMain 
            style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 115px" width=115>
	<PARAM NAME="_ExtentX" VALUE="3043">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturntoMain">
	<PARAM NAME="Caption" VALUE="Return to Main">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturntoMain()
{
	btnReturntoMain.value = 'Return to Main';
	btnReturntoMain.setStyle(0);
}
function _btnReturntoMain_ctor()
{
	CreateButton('btnReturntoMain', _initbtnReturntoMain, null);
}
</script>
<% btnReturntoMain.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
  </tr>
  <tr>
    <td></td>
  </tr>
</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

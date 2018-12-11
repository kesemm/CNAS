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

<body leftmargin="25" rightmargin="25" text="black" bgProperties="fixed" bgColor="#d7c7a4">

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

<p align="center"><font face="Arial Black" color="maroon" size="4"><strong>CO Code Request
Administration Menu</strong></font></p>

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="35"> </td>
    <td><font face="Arial"><a href="javascript:go('xca_Pending.asp')">Pending Requests</a> </font></td>
  </tr>
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="36"></td>
    <td><font face="Arial"><font face="Arial"><a href="javascript:go('xca_Part1adminPre.asp')"
    target>Enter New Requests ( Part 1)</a> Select this to request an NPA-NXX.&nbsp; You
    should receive a request ticket number on completion of request.&nbsp;</font> </font></td>
  </tr>
  
 <!-- <tr>
<!--    <td><img height="30" src="../images/ball25.gif" width="35"></td> -->
<!--  <td><font face="Arial"><a href="javascript:go('xca_Part1appEditPre.asp')" target>Edit -->
<!--     Existing Requests (Part 1)</a> by Ticket.&nbsp; Select this to edit a ticket your Entity -->
<!--     created.&nbsp;</font></td> -->
<!--   </tr> -->
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_Part1appViewPre.asp')">View Existing
    Requests (Part 1)</a> by Ticket.&nbsp; Select this to view a ticket your Entity
    created.&nbsp;</font></td>
  </tr>
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_Part1CancelPre.asp')">Cancel Existing
    Requests (Part 1)</a> by Ticket.&nbsp; Select this to cancel a ticket your Entity created
    and not processed by the CNAS Administrator</font> </td>
  </tr>
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_Part3Pre.asp')" target="page">Confirm/Deny&nbsp;
    Requests (Part 3)</a> by Ticket.&nbsp; </font></td>
  </tr>
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="35"></td>
    <td><font face="Arial"><a href="javascript:go('xca_Part3ViewPre.asp')" target="page">View
    Requests Confirmation (Part 3)</a> by Ticket.&nbsp; </font></td>
  </tr>
  <tr>
    <td><img height="30" src="../images/ball25.gif" width="36"></td>
    <td><font face="Arial"><a href="javascript:go('xca_Part4Pre.asp')">Enter In-Service Date
    for Existing Requests (Part 4)</a> by NPA-NXX </font></td>
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

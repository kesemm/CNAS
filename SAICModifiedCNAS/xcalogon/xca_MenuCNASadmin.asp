<%@ Language=VBScript %>

<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</form>
<!--#include file="xca_CNASLib.inc"-->

<form action="xca_MenuInt.asp" method="post" id="formP4" name="formAdminMenu" ONLOAD= history.go(0)">
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>CNAS Administration Menu</title>

<SCRIPT LANGUAGE="JavaScript"><!--


 function go(url) {
     location.href = url;
 }
 

	
--></script>
</head>
<body text="black" bgProperties=fixed bgColor="#d7c7a4">
<%
session("NoTixSent")=""


session("HereP3")="xca_MenuCNASadmin.asp"
	%>
<center>
<font face="Arial Black" color=maroon size=5><strong>ADMINISTRATION 
MENU</font></STRONG> <BR><BR><BR> 
</center>
<table align=center border=0 cellPadding=1 cellSpacing=1 width=35%>
 
    <tr>
        <td width=25><img height=36 src="../images/ball25.gif" 
            width=35> </td>
        <TD colSpan=3><A 
            href="javascript:go('xca_MenuC0CReqAdmin.asp')"><FONT face=Arial>CO Code Request 
            Administration</FONT></A></TD>
    <tr>
        <td><FONT><img height=36 src="../images/ball25.gif" 
            width=35></FONT></td>
        <TD colSpan=3><FONT face=Arial><FONT><A 
            href="javascript:go('xca_MenuC0CAdmin.asp')">CO 
            Code Administrati</FONT>on</A></FONT></TD>
    <tr>
        <td><img height=36 src="../images/ball25.gif" 
            width=35></td>
        <TD colSpan=3><FONT face=Arial><A 
            href="javascript:go('xca_RptAdminMenu.asp')">Administration Reports Menu</A></FONT></TD>
    <tr>
        <td><img height=36 src="../images/ball25.gif" 
            width=35></td>
        <TD colSpan=3><A 
            href="javascript:go('xca_MenuSecurityAdmin.asp')"><FONT face=Arial>Security and System 
            Administration</FONT></A></TD></tr>
          
    </table>
        </body><% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
        </HTML>

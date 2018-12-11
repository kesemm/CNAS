<%@ Language=VBScript %>

<%
Response.Buffer = true
Response.Expires=0
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<form action="xca_MenuInt.asp" method="post" id="formP4" name="formP4">
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Code Applicant Menu</title>

<SCRIPT LANGUAGE="JavaScript"><!--


 function go(url) {
     location.href = url;
 }
 

	
--></script>
<%


UserEntityType=session("UserEntityType")


 %>
 
  </head>
<body text="black" bgProperties="fixed" bgColor="#d7c7a4">
<center>
<font face="Arial Black" color=maroon size="5">CODE APPLICANT MANAGERS MENU</font></STRONG> 
</center>
<table align="center" border="0" cellPadding="1" cellSpacing="1" width="100%">
      <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"> 
</td>
        <td><font face="Arial"><a href="xca_Part1appPrePre.asp" target="">Input Part 
            1</a> Select this to request an NPA-NXX.&nbsp; You should receive a 
            request ticket number on completion of request.&nbsp;</font></td></tr>
    <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><a href="xca_Part1appEditPre.asp" target="">Edit Part 1</a> 
            by Ticket.&nbsp; Select this to edit a ticket your Entity 
            created.&nbsp;</font></td></tr>
    <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><a href="xca_Part1appViewPre.asp">View Part 1</a> 
            by Ticket.&nbsp; Select this to view a ticket your Entity 
            created.&nbsp;</font></td></tr>
    <tr>
     <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><a href="xca_Part1CancelPre.asp">Cancel Part 1</a> 
            by Ticket.&nbsp; Select this to cancel a ticket your Entity created 
            and not processed by the CNAS Administrator</font></td><font face="Arial"></font>
    <TR>
    <TD >
        <IMG height =36 src="../images/ball25.gif" width=35></TD>
        <TD colSpan=3><FONT face=Arial><A href="javascript:go('xca_Part3Pre.asp')" target =page>Confirm/Deny&nbsp; Requests (Part 3)</A> by 
            Ticket.&nbsp; 
            </FONT></TD>  
             <TR>
    <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><A 
            href="xca_Part3ViewPre.asp">View Part 
            3</A></font></td></tr>
            <p>&nbsp;</p></TD></tr>
    <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><a href="xca_Part4Pre.asp">Input Part 4 by NPA-NXX</font></A></td></tr>
    <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><A HREF="xca_RptPrtsFrmsPre.asp">Request Forms 
            Report: Part1, Part3, Part4</A></font></td></tr>
    <tr>
        <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
        <td><font face="Arial"><A HREF="xca_RptWebNPAStat.asp">NPA - CO Codes 
            Availability List</A> by NPA. &nbsp;Select here to see a list of 
            available CO Codes within a NPA.</font></td></tr>
    <tr>
        <td><font face="Arial"></font></td>
        <td><font face="Arial"></font></td></tr>
    <tr>
        <td><font face="Arial"></font></td>
        <td></td></tr>
    <tr>
        <td><font face="Arial"></font></td>
        <td><font face="Arial"></font></td></tr>
    <tr>
        <td></td>
        <td></td>
            </tr>
    <tr>
        <td></td>
        <td></td>
    <tr>
        <td><font face="Arial"></font>
        <td></td><font face="Arial"></font>
    <tr>
        <td><font face="Arial"></font>
        <td></td><font face="Arial"></font>
    <tr>
        <td><font face="Arial"></font>
        <td></td><font face="Arial"></font>
    <tr>
        <td><font face="Arial"></font>
        <td><font face="Arial"></font></td><font face="Arial"></font></tr></table></TABLE></FORM>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

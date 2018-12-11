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
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<%
NPA=session("aNPA")
NXX=session("aNXX")
RC=session("aRC")
SET objConnectionTix = server.createobject("ADODB.connection")
SET rstTixQry =server.createobject("ADODB.recordset")
objConnectionTix.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlTixQry = "Select Max(Tix) As Tix From xca_MBIPart1 Where Tix < 999999999 And NPA=" &NPA & " And NXX=" & NXX
Set rstTixQry = objConnectionTix.execute(sqlTixQry)
Tix=rstTixQry("Tix")

SET objConnection = server.createobject("ADODB.connection")
SET rstQry =server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
sqlQry = "Select count(*) As Number From xca_MBI Where Tix="  & Tix
Set rstQry = objConnection.execute(sqlQry)
Number=rstQry("Number")
%>
<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

Sub btnOK_onclick()
Response.Redirect "xca_MenuMBI.asp"
End Sub


</script>

<title></title>
</head>

<body leftmargin="15" bgColor="#d7c7a4" bgProperties="fixed" text="black">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<table align="center" WIDTH="75%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <tr>
    <td colSpan="3">&nbsp;&nbsp;&nbsp;&nbsp; </td>
  </tr>
  <tr>
    <td colSpan="3"><strong><font face="Arial">You have completed a <br>
    Part 1: MBI Block Assignment for:</font> </strong></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td nowrap align="left"><font face="Arial" size="3">NPA : </font></td>
    <td><font face="Arial" size="3"><strong><%Response.write NPA%></strong></font></td>
    <td></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td nowrap align="left"><font face="Arial" size="3">NXX : </font></td>
    <td><strong><font face="Arial" size="3"><%Response.write NXX%></font></strong></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td nowrap align="left"><font face="Arial" size="3">Rate Center : </font></td>
    <td><strong><font face="Arial" size="3"><%Response.write RC%></font></strong></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td nowrap align="left"><font face="Arial" size="3">Ticket #:</font></td>
    <td><strong><font face="Arial" size="3"><%Response.write Tix%></font></strong> </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
  </tr>
    <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td nowrap align="left"><font face="Arial" size="3">Number:</font><size="1">(if not 10 there is a problem)</td>
    <td><strong><font face="Arial" size="3"><%Response.write Number%></font></strong> </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
  </tr>
  <tr>
    <td align="left" noWrap>&nbsp;&nbsp; </td>
    <td noWrap></td>
    <td></td>
  </tr>
  <tr>
    <td><!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnOK style="HEIGHT: 27px; LEFT: 10px; TOP: 34px; WIDTH: 36px" 
            width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="OK">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnOK()
{
	btnOK.value = 'OK';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan--> </td>
    <td></td>
    <td></td>
  </tr>
</table>

<p>&nbsp;</p>
<%objConnectionTix.close 
objConnection.close%>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

<%@ Language=VBScript %>
<%
Response.Buffer = true
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>Administration Reports Menu</title>
<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

Sub btnReturntoMain_onclick()
	Response.Redirect "xca_MenuMainPost.asp"
End Sub

</script>
</head>

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>

<form action="xca_MenuInt.asp" method="post" id="formP4" name="formP4">
</form>

<p align="center"><font face="Arial Black" color="maroon" size="5">Administration Reports
Menu</font> </p>

<table align="center" border="0" style="HEIGHT: 200px; WIDTH: 45%">
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"> </td>
    <td><font face="Arial"><a href="xca_RptCOCStat.asp">CO Code Status Report</a></font></td>
  </tr>
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
    <td><font face="Arial"><a HREF="xca_RptPrtsFrmsPre.asp">CO Code Request Report: (All
    Part1, Part3, and Part4)</a></font></td>
  </tr>
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
    <td><font face="Arial"><a href="xca_RptCOCodesActLog.asp" target>CO Code Activity Log
    Report</a></font></td>
  </tr>
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
    <td><font face="Arial"><a href="xca_RptFormsActLog.asp" target>CO Code Request&nbsp;
    Activity Log Report</a></font></td>
  </tr>
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
    <td><font face="Arial"><a HREF="xca_RptWebNPAStat.asp">Available CO Code Report</a> </font></td>
  </tr>
  <tr>
    <td><img SRC="../images/ball25.gif" WIDTH="35" HEIGHT="36"></td>
    <td><font face="Arial"><a HREF="xca_CNA_ReportsMenu.asp">CNA Reports Menu</a> </font></td>
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

<p>&nbsp; </p>

<p>&nbsp;</p>

<p>&nbsp;</p>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

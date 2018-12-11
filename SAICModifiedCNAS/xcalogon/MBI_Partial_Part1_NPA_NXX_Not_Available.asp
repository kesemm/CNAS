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
aNPA=session("aNPA")
aNXX=session("aNXX")

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
    <td colSpan="3"><strong><font face="Arial"><%=aNPA %>-<%= aNXX%> that you have selected does not have any blocks available for assignment.</td>
	</tr>
  <tr>
    <td colSpan="3">&nbsp;&nbsp;&nbsp;&nbsp; </td>
  </tr>
<tr>
<td colSpan="3"><strong><font face="Arial">Please try again.</font> </strong></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp; </td>
    <td></td>
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
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

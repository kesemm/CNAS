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
</form>
<form action="xca_Part1app.asp" method="post" id="Part1AppPre" name="Part1AppPre">

<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
session("BlankP1")="BlankP1"


If session("UserEntityType") = "a" then
		Response.Redirect "xca_Part1adminPre.asp
else
		Response.Redirect "xca_Part1appPre.asp		
end if
%>
<P><center><font face="Arial Black" color=maroon size=5><strong>Input Part 1</strong></font></center></P>
<P>&nbsp;<P>
<P><font face="Arial" size="3"><font face="Arial"><strong><em>
Please enter the NPA 
for your Part 1 request to INPUT....</P>
<br><br>
</font>
<p>&nbsp; 

<table align="center" border="0" cellPadding="1" cellSpacing="1" width="50%">
    <tr>
        <td>
            <DIV align=right><font face="Arial"><STRONG>NPA:</STRONG> 
            </font></DIV>
        <td>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=NPA style="HEIGHT: 21px; LEFT: 10px; TOP: 34px; WIDTH: 96px" 
            width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="NPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="GetPart1NPA">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPA()
{
	GetPart1NPA.advise(RS_ONDATASETCOMPLETE, 'NPA.setRowSource(GetPart1NPA, \'NPA\', \'NPA\');');
}
function _NPA_ctor()
{
	CreateListbox('NPA', _initNPA, null);
}
</script>
<% NPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        <td>
<input type="submit" value="Go" id="button1" name="submit">
    <tr>
        <td></td>
        <td></td>
        <td><STRONG></STRONG></td></tr>
    <tr>
        <td>
        <td>

        <td></td></tr></table> 
<p>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=75%>
    
    <TR>
        <TD align=right><STRONG><FONT face="" size=2>Administrators Enter Applicant Entity Name:&nbsp;</FONT></STRONG> </TD>
        <TD>
          <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=EntityName 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 55px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="EntityName">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityName()
{
	EntityName.setStyle(TXT_TEXTBOX);
	EntityName.setMaxLength(35);
	EntityName.setColumnCount(35);
}
function _EntityName_ctor()
{
	CreateTextbox('EntityName', _initEntityName, null);
}
</script>
<% EntityName.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
            </TD>
        <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TD></TR>
    <TR>
        <TD></TD>
        <TD></TD>
        <TD></TD></TR>
    <TR>
        <TD></TD>
        <TD></TD>
        <TD></TD></TR></TABLE></p>
<hr align="left">
</form>

</EM></STRONG></font>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>

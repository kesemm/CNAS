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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%
Tix=session("P3TixCook")
NPA=session("P3NPACook")
NXX=session("P3NXXCook")
twoEmail=session("P3TwoEmailsCook")
%>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnOK_onclick()
Response.Redirect Session("HereP3")
End Sub


</SCRIPT>
</HEAD>
<body leftmargin=15 bgColor="#d7c7a4" bgProperties="fixed" text="black">
<TABLE align=center WIDTH=75% BORDER=0 CELLSPACING=0 CELLPADDING=0>
    
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp;&nbsp; 
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp;&nbsp; 
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp; 
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp;
	<TR>
		<TD colSpan=3></TD>
	</TR>
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp;&nbsp;
    <TR>
        <TD colSpan=3><FONT face=Arial><STRONG>You have successfully Denied or 
            Suspended a 
        <br>
        Part 3: Canadian CNA's 
            Response/Confirmation Form for:</FONT></FONT> </STRONG>
    <TR>
        <TD>&nbsp;&nbsp; 
        <TD>
        <TD>
	<TR>
		<TD nowrap align=left>
            <font face=Arial size=3>CO Code:</font></TD>
		<TD><font face=Arial size=3>
            <STRONG>
            <%Response.write NPA%>
            &nbsp;
            <%Response.write NXX%></font></STRONG></TD>
		<TD></TD>
	</TR>
    <TR>
        <TD align=left noWrap>&nbsp;&nbsp; 
        <TD>
        <TD>
    <TR>
        <TD nowrap align=left>
            <font face=Arial size=3>Ticket #:</font></TD>
        <TD><font face=Arial size=3>
            <STRONG>
            <%Response.write Tix%></font></STRONG>
        <TD><STRONG></STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            
    <TR>
        <TD align=left noWrap>&nbsp;&nbsp; 
        <TD noWrap>
        <TD>
    <TR>
        <TD nowrap align=left><font face=Arial size=3>Email <br>is being sent 
            to:</font></TD>
        <TD nowrap><font face=Arial size=3>
            <STRONG>
            <%Response.write twoEmail%></font></STRONG>
        <TD>
    <TR>
        <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <TD>
        <TD>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnOK style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 36px" 
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

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD></TD>
		<TD></TD>
	</TR>
</TABLE>

<P>&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

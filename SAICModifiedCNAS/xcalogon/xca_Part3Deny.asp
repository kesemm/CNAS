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
		<TD colSpan=2><IMG alt="" src="../images/stop.gif"
            height=52 
            style="HEIGHT: 52px; WIDTH: 56px" width=56></TD>
        <TD></TD>
    <TR>
        <TD colSpan=3>&nbsp;&nbsp;&nbsp;&nbsp;
    <TR>
        <TD colSpan=3><FONT face=Arial><STRONG>Sorry that CO Code is not 
            Available. Press the browser's back 
            button and input/change these fields.</FONT></FONT> </STRONG>
    <TR>
        <TD>&nbsp;&nbsp; 
        <TD>
        <TD>
	<TR>
		<TD nowrap align=left></TD>
		<TD><font face=Arial size=3>
            <STRONG>
            </font></STRONG></TD>
		<TD></TD>
	</TR>
    <TR>
        <TD align=left noWrap>&nbsp;&nbsp; 
        <TD>
        <TD>
    <TR>
        <TD nowrap align=left></TD>
        <TD><font face=Arial size=3>
            <STRONG>
            </font></STRONG>
        <TD><STRONG></STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            
    <TR>
        <TD align=left noWrap>&nbsp;&nbsp; 
        <TD noWrap>
        <TD>
    <TR>
        <TD nowrap align=left></TD>
        <TD nowrap><font face=Arial size=3>
            <STRONG>
            </font></STRONG>
        <TD>
    <TR>
        <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <TD>
        <TD>
	<TR>
		<TD>
            
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

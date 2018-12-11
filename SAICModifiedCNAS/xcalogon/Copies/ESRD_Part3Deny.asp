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
%>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnOK_onclick()
Response.Redirect "xca_MenuNANPCAN.asp"
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
        <TD colSpan=3><FONT face=Arial><STRONG>Sorry there is a disconnect between your logon and the Company that has been assigned that ESRD Block. Press the OK Button to return.</FONT></FONT> 
            </STRONG>
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
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnOK style="HEIGHT: 27px; LEFT: 10px; TOP: 34px; WIDTH: 36px" 
	width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="OK">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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

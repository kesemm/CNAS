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
Getdiffnum=session("P1diffnum")
P1DiffErr=session("P1DiffErr")
CenterNPAErr=session("CenterNPAErr")
RouteNPAErr=session("RouteNPAErr")
'RouteNPA1=session("RouteNPA")
'RouteNXX1=session("RouteNXX")
'CenterNPA1=session("CenterNPA")
'CenterNXX1=session("CenterNXX")
Part1Days.setCaption(Getdiffnum)
if P1DiffErr ="true" then

	RequestedDate.hide
	RequestedDateMsg.hide	
	Part1Days.hide
	Label1.hide
	
end if

if CenterNPAErr ="" then

	
	CenterNPA.hide
	NumErr.hide
	
end if

if RouteNPAErr ="" then
	RouteNPA.hide
	NXXDateMsg.hide
end if



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
        <TD colSpan=3>
    <TR>
        <TD colSpan=3><FONT face=Arial><STRONG>The Part 1 Request was not 
            complete!! The following fields are missing/incorrect.&nbsp; Press 
            the browser's back button and input/change these 
            fields.&nbsp;</STRONG></FONT>
    <TR>
        <TD>&nbsp;&nbsp; 
        <TD>
        <TD>
	<TR>
		<TD nowrap align=left><FONT 
            face=Arial><STRONG><U>Affected fields:</U></STRONG></FONT></TD>
		<TD><font face=Arial size=3>
            <STRONG>
            </font></STRONG></TD>
		<TD></TD>
	</TR>
    <TR>
        <TD align=left noWrap><FONT 
            face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=CenterNPA 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 34px; WIDTH: 147px" width=147>
	<PARAM NAME="_ExtentX" VALUE="3889">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CenterNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Use Same Rate Center">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCenterNPA()
{
	CenterNPA.setCaption('Use Same Rate Center');
}
function _CenterNPA_ctor()
{
	CreateLabel('CenterNPA', _initCenterNPA, null);
}
</script>
<% CenterNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->

            
 </FONT> 
        <TD><FONT face=Arial></FONT>
        <TD><FONT face=Arial></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NumErr style="HEIGHT: 20px; LEFT: 10px; TOP: 54px; WIDTH: 310px" 
            width=310>
	<PARAM NAME="_ExtentX" VALUE="8202">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NumErr">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Is not in Service, please select another CO Code">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNumErr()
{
	NumErr.setCaption('Is not in Service, please select another CO Code');
}
function _NumErr_ctor()
{
	CreateLabel('NumErr', _initNumErr, null);
}
</script>
<% NumErr.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD nowrap align=left><FONT face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RouteNPA 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 74px; WIDTH: 131px" width=131>
	<PARAM NAME="_ExtentX" VALUE="3466">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RouteNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Route Same as NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRouteNPA()
{
	RouteNPA.setCaption('Route Same as NPA');
}
function _RouteNPA_ctor()
{
	CreateLabel('RouteNPA', _initRouteNPA, null);
}
</script>
<% RouteNPA.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->

             
            </FONT></TD>
        <TD>
        <TD><STRONG><FONT 
            face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXXDateMsg 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 94px; WIDTH: 310px" width=310>
	<PARAM NAME="_ExtentX" VALUE="8202">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXDateMsg">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Is not in Service, please select another CO Code">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNXXDateMsg()
{
	NXXDateMsg.setCaption('Is not in Service, please select another CO Code');
}
function _NXXDateMsg_ctor()
{
	CreateLabel('NXXDateMsg', _initNXXDateMsg, null);
}
</script>
<% NXXDateMsg.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
            </FONT></STRONG> 
            
    <TR>
        <TD align=left noWrap>
        
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RequestedDate style="HEIGHT: 20px; WIDTH: 162px" width=162><PARAM NAME="_ExtentX" VALUE="4286"><PARAM NAME="_ExtentY" VALUE="529"><PARAM NAME="id" VALUE="RequestedDate"><PARAM NAME="DataSource" VALUE=""><PARAM NAME="DataField" VALUE="Requested Effective Date"><PARAM NAME="FontFace" VALUE="Arial"><PARAM NAME="FontSize" VALUE="2"><PARAM NAME="FontColor" VALUE="red"><PARAM NAME="FontBold" VALUE="-1"><PARAM NAME="FontItalic" VALUE="0"><PARAM NAME="Visible" VALUE="-1"><PARAM NAME="FormatAsHTML" VALUE="0"><PARAM NAME="Platform" VALUE="256"><PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedDate()
{
	RequestedDate.setCaption('Requested Effective Date');
}
function _RequestedDate_ctor()
{
	CreateLabel('RequestedDate', _initRequestedDate, null);
}
</script>
<% RequestedDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> 
        <TD noWrap><FONT face=Arial></FONT>
        <TD><FONT face=Arial></FONT>

            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=RequestedDateMsg 
	style="HEIGHT: 20px; LEFT: 10px; TOP: 151px; WIDTH: 53px" width=53>
	<PARAM NAME="_ExtentX" VALUE="1402">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RequestedDateMsg">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Must be">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRequestedDateMsg()
{
	RequestedDateMsg.setCaption('Must be');
}
function _RequestedDateMsg_ctor()
{
	CreateLabel('RequestedDateMsg', _initRequestedDateMsg, null);
}
</script>
<% RequestedDateMsg.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Part1Days 
	style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 73px" width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Part1Days">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="getDiffnum">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="BLUE">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="BLUE"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPart1Days()
{
	Part1Days.setCaption('getDiffnum');
}
function _Part1Days_ctor()
{
	CreateLabel('Part1Days', _initPart1Days, null);
}
</script>
<% Part1Days.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 id=Label1 style="HEIGHT: 20px; LEFT: 10px; TOP: 174px; WIDTH: 200px" 
	width=200>
	<PARAM NAME="_ExtentX" VALUE="5292">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="days past the Application Date.">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('days past the Application Date.');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
    <TR>
        <TD nowrap align=left></TD>
        <TD nowrap><font face=Arial size=3>
            <STRONG><FONT>
            
            </FONT></font><FONT></STRONG></FONT>
        <TD><FONT></FONT>
    <TR>
        <TD>
        <TD><FONT></FONT>
        <TD><FONT></FONT>
    <TR>
        <TD>
        <TD>
        <TD>
	<TR>
		<TD>
</TD>
		<TD><FONT></FONT></TD>
		<TD>
</TD>
	</TR>
</TABLE><FONT></FONT>

<P>&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

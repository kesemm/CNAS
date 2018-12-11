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
AssignedNXXErr=session("AssignedNXXErr")
EFFDateErr=session("EFFDateErr")
RRDescriptionErr=session("RRDescriptionErr")
LERGDateErr=session("LERGDateErr")
RRReturnErr=session("RRReturnErr")
ReservedNXXErr=session("ReservedNXXErr")
CNAResponsibleErr=session("CNAResponsibleErr")


if AssignedNXXErr ="" then
	AssignedNXX.hide
	NumErr.hide
end if

if EFFDateErr ="" then
	EFFDate.hide
	NXXDateMsg.hide
end if

if RRDescriptionErr ="" then
	RRDescription.hide
end if

if LERGDateErr ="" then
	LERGDate.hide
end if

if RRReturnErr ="" then
	RRReturnDate.hide
end if

if RRDescriptionErr ="" then
	RRDescription.hide
end if

if ReservedNXXErr ="" then
	ReservedNXX.hide
	NumErr1.hide
end if

if CNAResponsibleErr ="" then
	CNAResponsible.hide
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
        <TD colSpan=3><FONT face=Arial><STRONG>The Part 3 Response was not 
            complete!! The following fields are missing.&nbsp; Press the 
            browser's back button and input/change these 
            fields.&nbsp;</STRONG></FONT>
    <TR>
        <TD>&nbsp;&nbsp; 
        <TD>
        <TD>
	<TR>
		<TD nowrap align=left><FONT 
            face=Arial><STRONG><U>Missing fields:</U></STRONG></FONT></TD>
		<TD><font face=Arial size=3>
            <STRONG>
            <%Response.write NXX%></font></STRONG></TD>
		<TD></TD>
	</TR>
    <TR>
        <TD align=left noWrap><FONT 
            face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=AssignedNXX 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 92px" width=92>
	<PARAM NAME="_ExtentX" VALUE="2434">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="AssignedNXX">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Assigned NXX">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAssignedNXX()
{
	AssignedNXX.setCaption('Assigned NXX');
}
function _AssignedNXX_ctor()
{
	CreateLabel('AssignedNXX', _initAssignedNXX, null);
}
</script>
<% AssignedNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </FONT> 
        <TD><FONT face=Arial></FONT>
        <TD><FONT face=Arial></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NumErr style="HEIGHT: 20px; LEFT: 10px; TOP: 54px; WIDTH: 265px" 
            width=265>
	<PARAM NAME="_ExtentX" VALUE="7011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NumErr">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Is missing, not a number, or less than 200">
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
	NumErr.setCaption('Is missing, not a number, or less than 200');
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
            id=EFFDate style="HEIGHT: 20px; LEFT: 10px; TOP: 54px; WIDTH: 122px" 
            width=122>
	<PARAM NAME="_ExtentX" VALUE="3228">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="EFFDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NXX Effective Date">
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
function _initEFFDate()
{
	EFFDate.setCaption('NXX Effective Date');
}
function _EFFDate_ctor()
{
	CreateLabel('EFFDate', _initEFFDate, null);
}
</script>
<% EFFDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan--> 
            </FONT></TD>
        <TD>
        <TD><STRONG><FONT 
            face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NXXDateMsg 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 94px; WIDTH: 364px" width=364>
	<PARAM NAME="_ExtentX" VALUE="9631">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NXXDateMsg">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Must be equal to or greater than the Part 1 Effective Date">
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
	NXXDateMsg.setCaption('Must be equal to or greater than the Part 1 Effective Date');
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
        <TD align=left noWrap><FONT 
            face=Arial>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RRDescription 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 201px" width=201>
	<PARAM NAME="_ExtentX" VALUE="5318">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RRDescription">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Routing and Rating Description">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRRDescription()
{
	RRDescription.setCaption('Routing and Rating Description');
}
function _RRDescription_ctor()
{
	CreateLabel('RRDescription', _initRRDescription, null);
}
</script>
<% RRDescription.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </FONT> 
        <TD noWrap><FONT face=Arial></FONT>
        <TD><FONT face=Arial></FONT>
    <TR>
        <TD nowrap align=left><FONT 
            face=Arial><FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=LERGDate style="HEIGHT: 20px; WIDTH: 72px" width=72>
	<PARAM NAME="_ExtentX" VALUE="1905">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="LERGDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="LERG Date">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Red">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<FONT FACE="Arial" SIZE="2" COLOR="Red"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLERGDate()
{
	LERGDate.setCaption('LERG Date');
}
function _LERGDate_ctor()
{
	CreateLabel('LERGDate', _initLERGDate, null);
}
</script>
<% LERGDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </FONT></FONT><FONT></FONT></TD>
        <TD nowrap><font face=Arial size=3>
            <STRONG><FONT>
            <%Response.write twoEmail%>
            </FONT></font><FONT></STRONG></FONT>
        <TD><FONT></FONT>
    <TR>
        <TD><FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=RRReturnDate 
            style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 205px" width=205>
	<PARAM NAME="_ExtentX" VALUE="5424">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="RRReturnDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Routing and Rating Return Date">
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
function _initRRReturnDate()
{
	RRReturnDate.setCaption('Routing and Rating Return Date');
}
function _RRReturnDate_ctor()
{
	CreateLabel('RRReturnDate', _initRRReturnDate, null);
}
</script>
<% RRReturnDate.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
 </FONT>
        <TD><FONT></FONT>
        <TD><FONT></FONT>
    <TR>
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=CNAResponsible 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 134px; WIDTH: 112px" 
width=112>
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="CNAResponsible">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="CNA Responsible">
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
function _initCNAResponsible()
{
	CNAResponsible.setCaption('CNA Responsible');
}
function _CNAResponsible_ctor()
{
	CreateLabel('CNAResponsible', _initCNAResponsible, null);
}
</script>
<% CNAResponsible.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD>
        <TD>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=ReservedNXX style="HEIGHT: 20px; WIDTH: 94px" width=94>
	<PARAM NAME="_ExtentX" VALUE="2487">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="ReservedNXX">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Reserved NXX">
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
function _initReservedNXX()
{
	ReservedNXX.setCaption('Reserved NXX');
}
function _ReservedNXX_ctor()
{
	CreateLabel('ReservedNXX', _initReservedNXX, null);
}
</script>
<% ReservedNXX.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD><FONT></FONT></TD>
		<TD><FONT></FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=20 
            id=NumErr1 
            style="HEIGHT: 20px; LEFT: 10px; TOP: 194px; WIDTH: 265px" 
width=265>
	<PARAM NAME="_ExtentX" VALUE="7011">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="NumErr1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Is missing, not a number, or less than 200">
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
function _initNumErr1()
{
	NumErr1.setCaption('Is missing, not a number, or less than 200');
}
function _NumErr1_ctor()
{
	CreateLabel('NumErr1', _initNumErr1, null);
}
</script>
<% NumErr1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><FONT></FONT>

<P>&nbsp;</P>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

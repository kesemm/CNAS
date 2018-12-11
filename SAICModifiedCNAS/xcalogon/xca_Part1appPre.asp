<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</form>
<form action="xca_Part1app.asp" method="post" id="Part1AppPre" name="Part1AppPre">
<!--#include file="xca_CNASLib.inc"-->
<html>
<head>
<META HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
session("BlankP1")="Applicant"
session("Here")="xca_Part1appPre.asp"



%>
<P><center><font face="Arial Black" color=maroon size=5><strong>Input Part 
1</strong></font></center>
<P></P>
<P>&nbsp;<P>
<P><font face="Arial" size="3"><font face="Arial"><strong><em>
Please enter the NPA for your Part 1 request to INPUT....</P>
<br><br></FONT>
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
        <td>
            <DIV align=right><STRONG>Type of </STRONG><STRONG>Request:</STRONG> 
            </DIV></td>
        <td>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=SelectedTypeOfReq 
            style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96><PARAM NAME="_ExtentX" VALUE="2170"><PARAM NAME="_ExtentY" VALUE="556"><PARAM NAME="id" VALUE="SelectedTypeOfReq"><PARAM NAME="DataSource" VALUE=""><PARAM NAME="DataField" VALUE=""><PARAM NAME="ControlStyle" VALUE="0"><PARAM NAME="Enabled" VALUE="-1"><PARAM NAME="Visible" VALUE="-1"><PARAM NAME="Platform" VALUE="256"><PARAM NAME="UsesStaticList" VALUE="-1"><PARAM NAME="CLSize" VALUE="3"><PARAM NAME="CLED1" VALUE="Assign"><PARAM NAME="CLEV1" VALUE="A"><PARAM NAME="CLED2" VALUE="Update"><PARAM NAME="CLEV2" VALUE="U"><PARAM NAME="CLED3" VALUE="Reserve"><PARAM NAME="CLEV3" VALUE="R"><PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initSelectedTypeOfReq()
{
	SelectedTypeOfReq.addItem('Assign', 'A');
	SelectedTypeOfReq.addItem('Update', 'U');
	SelectedTypeOfReq.addItem('Reserve', 'R');
}
function _SelectedTypeOfReq_ctor()
{
	CreateListbox('SelectedTypeOfReq', _initSelectedTypeOfReq, null);
}
</script>
<% SelectedTypeOfReq.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td><STRONG></STRONG></td></tr>
    <tr>
        <td>
        <td>

        <td></td></tr>
        </table>
<hr align="left">
</FORM>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1NPA 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sStatus='S'\sorder\sby\sNPA\q,TCControlID_Unmatched=\qGetPart1NPA\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sStatus='S'\sorder\sby\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1NPA()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasapp_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasapp_CommandTimeout');
	DBConn.CursorLocation = Application('cnasapp_CursorLocation');
	DBConn.Open(Application('cnasapp_ConnectionString'), Application('cnasapp_RuntimeUserName'), Application('cnasapp_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'select distinct NPA from xca_COCode where Status=\'S\' order by NPA';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart1NPA.setRecordSource(rsTmp);
	GetPart1NPA.open();
	if (thisPage.getState('pb_GetPart1NPA') != null)
		GetPart1NPA.setBookmark(thisPage.getState('pb_GetPart1NPA'));
}
function _GetPart1NPA_ctor()
{
	CreateRecordset('GetPart1NPA', _initGetPart1NPA, null);
}
function _GetPart1NPA_dtor()
{
	GetPart1NPA._preserveState();
	thisPage.setState('pb_GetPart1NPA', GetPart1NPA.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
</EM></STRONG></FONT>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>

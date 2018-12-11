<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
</form>
<form action="ESRDInput.asp" method="post" id="ESRD" name="ESRD">
<!--#include file="xca_CNASLib.inc"-->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<%
session("BlankP1")="Admin"
session("Here")="ESRD.asp"
%>

<P><center><font face="Arial Black" size=5 color=maroon><strong>Input ESRD Data</strong></font></center>
<P></P>

<P>&nbsp;<P>

<p><font face="Arial" size="3"><font face="Arial"><strong><em>
Please enter the NPA for your ESRD request ...</p>
<br><br><br><br></FONT>

<div align="center">
  <center>
  <table border="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="80%">
    <tr>
      <td width="100%" align="center">            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=NPA511 style="HEIGHT: 21px; LEFT: 10px; TOP: 34px; WIDTH: 130px" 
            width=130>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="NPA511">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="GetPart1NPA511">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPA511()
{
	GetPart1NPA511.advise(RS_ONDATASETCOMPLETE, 'NPA511.setRowSource(GetPart1NPA511, \'NPA\', \'NPA\');');
}
function _NPA511_ctor()
{
	CreateListbox('NPA511', _initNPA511, null);
}
</script>
<% NPA511.display %> <strong> - 511 ESRDs</strong>

<!--METADATA TYPE="DesignerControl" endspan--></td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">           <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=NPA211 style="HEIGHT: 21px; LEFT: 10px; TOP: 34px; WIDTH: 130px" 
            width=130>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="NPA211">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="GetPart1NPA211">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initNPA211()
{
	GetPart1NPA211.advise(RS_ONDATASETCOMPLETE, 'NPA211.setRowSource(GetPart1NPA211, \'NPA\', \'NPA\');');
}
function _NPA211_ctor()
{
	CreateListbox('NPA211', _initNPA211, null);
}
</script>
<% NPA211.display %> <strong> - 211 ESRDs</strong>
</td>
<!--METADATA TYPE="DesignerControl" endspan--></td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">
      <p align="center"><font size="3" face="Arial"><strong>Choose to submit for 
      511 or 211 ESRDs and click Go</strong></font></td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">
      <p align="center"><font size="2" face="Arial"><strong>(NOTE: 511 ESRDs MUST be exhausted in the NPA before choosing 211 ESRDs)</strong></font></td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="center">
      <p align="center"><font size="3" face="Arial"><strong>
      <input id="NXXSelect" name="NXXSelect" value="511" checked style type="radio">511 &nbsp;&nbsp;&nbsp;&nbsp;or&nbsp; &nbsp;&nbsp;&nbsp;<input id="NXXSelect" name="NXXSelect" value="211" style type="radio">211&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <input type="submit" value="Go" id="button1" name="submit"></strong></font></td>
    </tr>
    </table>
  </center>
</div>




</FORM>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1NPA511 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sNXX=511 And Status='S'\sorder\sby\sNPA\q,TCControlID_Unmatched=\qGetPart1NPA511\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sNXX=511 And Status='S'\sorder\sby\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1NPA511()
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
	cmdTmp.CommandText = 'select distinct NPA from xca_ESRD where NXX=511 And Status=\'S\' order by NPA';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart1NPA511.setRecordSource(rsTmp);
	GetPart1NPA511.open();
	if (thisPage.getState('pb_GetPart1NPA511') != null)
		GetPart1NPA511.setBookmark(thisPage.getState('pb_GetPart1NPA511'));
}
function _GetPart1NPA511_ctor()
{
	CreateRecordset('GetPart1NPA511', _initGetPart1NPA511, null);
}
function _GetPart1NPA511_dtor()
{
	GetPart1NPA511._preserveState();
	thisPage.setState('pb_GetPart1NPA511', GetPart1NPA511.getBookmark());
}
</SCRIPT>
<!--METADATA TYPE="DesignerControl" endspan-->


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPart1NPA211 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sNXX=211 And Status='S'\sorder\sby\sNPA\q,TCControlID_Unmatched=\qGetPart1NPA211\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\sdistinct\sNPA\sfrom\sxca_COCode\swhere\sNXX=211 And Status='S'\sorder\sby\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPart1NPA211()
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
	cmdTmp.CommandText = 'select distinct NPA from xca_ESRD where NXX=211 And Status=\'S\' order by NPA';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPart1NPA211.setRecordSource(rsTmp);
	GetPart1NPA211.open();
	if (thisPage.getState('pb_GetPart1NPA211') != null)
		GetPart1NPA211.setBookmark(thisPage.getState('pb_GetPart1NPA211'));
}
function _GetPart1NPA211_ctor()
{
	CreateRecordset('GetPart1NPA211', _initGetPart1NPA211, null);
}
function _GetPart1NPA211_dtor()
{
	GetPart1NPA211._preserveState();
	thisPage.setState('pb_GetPart1NPA211', GetPart1NPA211.getBookmark());
}
</SCRIPT>
<!--METADATA TYPE="DesignerControl" endspan-->

</EM></STRONG></FONT>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>

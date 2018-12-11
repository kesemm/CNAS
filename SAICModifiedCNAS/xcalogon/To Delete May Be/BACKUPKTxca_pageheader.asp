<%@ Language=VBScript%>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<title></title>
</head>

<body leftmargin="2" bgColor="#d7c7a4" bgProperties="fixed" text="black" bottomMargin="1"
topMargin="0">

<form name="thisForm" METHOD="post">
</form>
<%
sql="select Value from xca_Parms where Name='IP'"
getIP.getSQLText(sql)
getIP.open
IPaddress=trim(getIP.fields.getValue("Value"))
IPaddress="209.195.96.131"


%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=getIP 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Parms\swhere\sName=\s'IP'\q,TCControlID_Unmatched=\qgetIP\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Parms\swhere\sName=\s'IP'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initgetIP()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('cnasadmin_ConnectionTimeout');
	DBConn.CommandTimeout = Application('cnasadmin_CommandTimeout');
	DBConn.CursorLocation = Application('cnasadmin_CursorLocation');
	DBConn.Open(Application('cnasadmin_ConnectionString'), Application('cnasadmin_RuntimeUserName'), Application('cnasadmin_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'Select * from xca_Parms where Name= \'IP\'';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	getIP.setRecordSource(rsTmp);
	if (thisPage.getState('pb_getIP') != null)
		getIP.setBookmark(thisPage.getState('pb_getIP'));
}
function _getIP_ctor()
{
	CreateRecordset('getIP', _initgetIP, null);
}
function _getIP_dtor()
{
	getIP._preserveState();
	thisPage.setState('pb_getIP', getIP.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<table align="left" cellPadding="0" cellSpacing="0" width="506" text="black" border="0"
size="60%" height="55" style="HEIGHT: 55px; WIDTH: 506px">
  <tr>
    <td noWrap colSpan="4" vAlign="bottom"
    style="PADDING-BOTTOM: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; PADDING-TOP: 0px"><img
    align="textTop" src="../images/canada1.gif" height="23" width="70" border="0"
    style="HEIGHT: 23px; WIDTH: 70px"> <strong><font face="Arial" size="4" color="darkgreen"><font
    face="Arial Black">C</font>anadian <font face="Arial Black">N</font>umbering <font
    face="Arial Black">A</font>dministration <font face="Arial Black">S</font>ystem</font></strong>
    </td>
    <td noWrap></td>
    <td noWrap align="middle" vAlign="bottom"></td>
    <td noWrap align="middle" vAlign="bottom"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;</font>
    </td>
  </tr>
  <tr>
    <td noWrap width="150" align="left" style="WIDTH: 150px" vAlign="baseline"></td>
    <td align="left" vAlign="baseline"><a
    href="http://<%=ipaddress%>/xcalogon/xca_MenuMainPost.asp" target="page"
    style="TEXT-DECORATION: none; VERTICAL-ALIGN: top"><font face="Verdana" size="2">MAIN MENU</font></a>
    </td>
    <td align="middle" vAlign="baseline"><a
    href="http://<%=ipaddress%>/xcalogon/xca_MenuMain.asp" target="page"
    style="TEXT-DECORATION: none; VERTICAL-ALIGN: top"><font face="Verdana" size="2">ENTRY
    PAGE</font></a> </td>
    <td align="middle" vAlign="baseline"><a href="http://<%=ipaddress%>/Default.asp"
    target="page" style="TEXT-DECORATION: none; VERTICAL-ALIGN: top"><font face="Verdana"
    size="2">EXIT</font></a></td>
  </tr>
</table>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</html>

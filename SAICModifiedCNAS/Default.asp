<%@ Language=VBScript %>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: Default.asp,v $
'* Commit Date:   $Date: 2015/01/19 13:49:19 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
<%
response.buffer=true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</form>
<form action="xca_logincheck.asp" method="post" id="form1" name="form1">
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>CNAS Home Page</title>
<%
If session("UserEntityType") <> ""   then
		
        Session.Abandon
        
 end if 
 
 GetPhone="Select Value from xca_Parms where Name= 'HotLine'"
	
	GetPhoneNum.setSQLText(GetPhone)
	GetPhoneNum.open
	
		Phone = GetPhoneNum.fields.getValue("Value")
 
 
 
 
 %>

</head>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetPhoneNum style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\s*\sfrom\sxca+Parms\swhere\sName=?\q,TCControlID_Unmatched=\qGetPhoneNum\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\s*\sfrom\sxca+Parms\swhere\sName=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))"></OBJECT>
-->
<!--#INCLUDE FILE="_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetPhoneNum()
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
	cmdTmp.CommandText = 'select * from xca+Parms where Name=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetPhoneNum.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetPhoneNum') != null)
		GetPhoneNum.setBookmark(thisPage.getState('pb_GetPhoneNum'));
}
function _GetPhoneNum_ctor()
{
	CreateRecordset('GetPhoneNum', _initGetPhoneNum, null);
}
function _GetPhoneNum_dtor()
{
	GetPhoneNum._preserveState();
	thisPage.setState('pb_GetPhoneNum', GetPhoneNum.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<body leftmargin=15 bgColor="#d7c7a4" bgProperties="fixed" text="black">
<p>
<table border="0" cellPadding="1" cellSpacing="1" height="113" style="HEIGHT: 113px; WIDTH: 517px" width="96.28%" align="center">
   
    <tr>
        <td colSpan="3"><font face="Arial Black"><font size="7">W<font size="2"><font face="Arial">elcome to Leidos Canada Modified CNAS.&nbsp; 
            </font></font></font></font> </td></tr>
    <tr>
        <td colSpan="3"><font face="Arial"><font size="2">&nbsp;</font></font> 



    <tr>
        <td colSpan="3"><font face="arial"><strong>Please enter your CNAS Logon ID 
            and Password....</strong></font></td></tr>
    <tr>
        <td>
        <td>
        <td>
    <tr>
        <td>
        <td>
        <td>
    <tr>
        <td align="right"><font face="Arial">User ID:
            </font>
        <td>

<input id="text" name="UID">
        <td>
    <tr>
        <td></td>
        <td></td>
        <td></td></tr>
    <tr>
        <td align="right"><font face="Arial">Password 
            </font>
        <td>
<input type="password" id="password1" name="PWD">
        <td>
    <tr>
        <td align="right">
        <td>
        <td>
    <tr>
        <td align="right"></td>
        <td>
            <p>
<input type="submit" value="Go" id="button1" name="submit"></p>
            <p>&nbsp;</p></td>
        <td></td></tr></table></FONT></FONT><font size="2"></font>
    <tr>
        <td colSpan = "3"><font face="Arial"><font size="2">
<table border=0 cellpadding=0 cellspacing=0 height=90 style="HEIGHT: 90px; WIDTH: 603px" width=75% align=center>
    
    <tr>
        <td colSpan="3"><font face=arial size=2>If you have trouble logging in, please contact 
            the CNAS Administrator @&nbsp;
            <%=phone%>
             during weekday business hours (8:00am - 5:00pm EST)</font> </td></tr>
    <tr>
        <td colSpan="3">&nbsp;</td></tr>
    <tr>
        <td colSpan="3"><font face="Arial"><font size="2"><strong>Best when viewed using 
            Microsoft Internet Explorer version 4.0  or higher with 'cookies' and 'request new versions 
            of stored pages always' enabled.&nbsp; 
            </strong></font> </font>
    <TR>
        <TD colSpan=3>&nbsp;
    <tr>
        <td colSpan="3">&nbsp; 
    <TR>
        <TD colSpan=3>&nbsp; 
    <TR>
        <TD colSpan=3>
            <DIV align=center><STRONG></STRONG></DIV> 
  
     <tr>
        <td colSpan="3">
                <tr>
        <td colSpan = "3"><font face="Arial"><font size="2"></font></font><font size="2"></font></td></tr>
    <tr>
        <td colSpan = "3">
    <tr>
        <td><font>&nbsp; </font> 
    <tr>
        <td><font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  </font>
            </td>
            </tr></TABLE><font></font></STRONG>
            </p>
            
     
<p><STRONG>&nbsp; </STRONG></p><STRONG>
<hr align="left">
</TABLE></STRONG>
<center><STRONG></STRONG><table>
<tr><td><STRONG>&nbsp;

<a HREF="http://www.microsoft.com/ie"><img src="ie_anim.gif" border="0" WIDTH="88" HEIGHT="31"></a> 
</STRONG> 
    

</td>
</tr></table><STRONG></STRONG></center><STRONG></FORM></STRONG>
            
            
            
            

<p><font><STRONG>&nbsp; 
</font></FONT></STRONG></p> 



</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
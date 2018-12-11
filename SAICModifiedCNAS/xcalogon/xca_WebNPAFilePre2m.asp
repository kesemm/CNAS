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
<!--#include file="xca_CNASLib.inc"-->
<form action="xca_GenerateNPAFilesConfirm.asp" method="post" id="form1" name="form1" LANGUAGE="javascript">

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<%

Sub btnSUBMIT1_onclick()
Dim fname, str,fileStr,idxtmp

Dim npaCnt 


	str=""		
	sqlNPALookUp = "SELECT DISTINCT NPA FROM xca_COCode Where Status <> 's' Order By NPA"
	NPALookUp.setSQLText(sqlNPALookUp)
	NPALookUp.open()
	
	npaCnt = NPALookUp.getCount() - 1
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim NPANameArray()
	ReDim NPANameArray(npaCnt)
''this B O C populates the the NPANameArray array with the valid NPA numbers
	for  i = 0 to npaCnt step 1       
		'Response.Write(i)
		NPANameArray(i) = NPALookUp.fields.getValue("NPA")			
		NPALookUp.moveNext()
	next 
'' This B O C iterates 0 to the number of NPAs - 1 making  a new file for each npa with the NXXs Utilized
	For i = 0 to npaCnt
		str = ""
		str = str + "NPA" + chr(9) 
		str = str + "NXX" + chr(13)+ chr(10)
		nm = NPANameArray(i)
		
		sqlNPAUtil = "SELECT DISTINCT NPA,NXX FROM xca_COCode Where NPA = '"&nm&"' and Status <> 's' Order By NXX"
		NPANXXUtilRec.setSQLText(sqlNPAUtil)
		NPANXXUtilRec.open
		NPANXXUtilRec.moveFirst
		
		Do While not NPANXXUtilRec.EOF 						
			str = str + NPANXXUtilRec.fields.getValue("NPA") + chr(9) 
			str = str + NPANXXUtilRec.fields.getValue("NXX") +  chr(13)+ chr(10)
			NPANXXUtilRec.moveNext
		Loop

		fileStr = str
		npaFileName = NPANameArray(i)		
		fname = "c:\xca_CNASNPAFiles\npaFileUtil_"+npaFileName+".txt"
		set	objNf = fso.CreateTextFile(fname, True)
		objNf.Write(fileStr)
		objNf.Close
		set gf2 = fso.GetFile(fname)
		NPANXXUtilRec.close

	Next
	Response.Redirect("xca_GenerateAllNPAFilesUtilConfirm.asp")
End Sub


Sub btnSUBMIT2_onclick()
Dim fname, str,fileStr,idxtmp

Dim npaCnt 


	str=""		
	sqlNPALookUp = "SELECT DISTINCT NPA FROM xca_COCode Where Status <> 's' Order By NPA"
	NPALookUp.setSQLText(sqlNPALookUp)
	NPALookUp.open()
	
	npaCnt = NPALookUp.getCount() - 1
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	Dim NPANameArray()
	ReDim NPANameArray(npaCnt)
''this B O C populates the the NPANameArray array with the valid NPA numbers
	for  i = 0 to npaCnt step 1       
		'Response.Write(i)
		NPANameArray(i) = NPALookUp.fields.getValue("NPA")			
		NPALookUp.moveNext()
	next 
'' This B O C iterates 0 to the number of NPAs - 1 making  a new file for each npa with the NXXs Utilized
	For i = 0 to npaCnt
		str = ""
		'str = str + "NPA" + chr(9) + chr(9)
		'str = str + "NXX" + chr(10) 
		
		str = str + "NPA" + chr(9)
		str = str + "NXX" + chr(13)+ chr(10) 
		
		nm = NPANameArray(i)		
		sqlNPAAvl = "SELECT DISTINCT NPA,NXX FROM xca_COCode Where NPA = '"&nm&"' and Status = 's' Order By NXX"
		NPANXXAvlRec.setSQLText(sqlNPAAvl)
		NPANXXAvlRec.open
		NPANXXAvlRec.moveFirst		
		Do While not NPANXXAvlRec.EOF 						
			'str = str + NPANXXAvlRec.fields.getValue("NPA") + chr(9) + chr(9)
			'str = str + NPANXXAvlRec.fields.getValue("NXX") + chr(10) 
			
			str = str + NPANXXAvlRec.fields.getValue("NPA") + chr(9)
			str = str + NPANXXAvlRec.fields.getValue("NXX") + chr(13)+ chr(10) 
			
			NPANXXAvlRec.moveNext
		Loop

		fileStr = str
		npaFileName = NPANameArray(i)		
		fname = "c:\xca_CNASNPAFiles\npaFileAvl_"+npaFileName+".txt"
		set	objNf = fso.CreateTextFile(fname, True)
		objNf.Write(fileStr)
		objNf.Close
		set gf2 = fso.GetFile(fname)
		NPANXXAvlRec.close

	Next
	Response.Redirect("xca_GenerateAllNPAFilesConfirm.asp")
End Sub

%>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub Button2_onclick()
Response.Redirect "xca_WebNPAFileMenu.htm"
End Sub


</SCRIPT>
</head>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPANXXUtilRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s\l\g\s's'\sOrder\sBy\sNXX\q,TCControlID_Unmatched=\qNPANXXUtilRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s\l\g\s's'\sOrder\sBy\sNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNPANXXUtilRec()
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
	cmdTmp.CommandText = 'Select * From xca_COCodes Where NPA=? and Staus <> \'s\' Order By NXX';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NPANXXUtilRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_NPANXXUtilRec') != null)
		NPANXXUtilRec.setBookmark(thisPage.getState('pb_NPANXXUtilRec'));
}
function _NPANXXUtilRec_ctor()
{
	CreateRecordset('NPANXXUtilRec', _initNPANXXUtilRec, null);
}
function _NPANXXUtilRec_dtor()
{
	NPANXXUtilRec._preserveState();
	thisPage.setState('pb_NPANXXUtilRec', NPANXXUtilRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPANXXAvlRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sDISTINCT\sNPA,NXX\sFROM\sxca_COCode\sWhere\sNPA\s=?'\sand\sStatus\s=\s's'\sOrder\sBy\sNXX\q,TCControlID_Unmatched=\qNPANXXAvlRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sDISTINCT\sNPA,NXX\sFROM\sxca_COCode\sWhere\sNPA\s=?'\sand\sStatus\s=\s's'\sOrder\sBy\sNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNPANXXAvlRec()
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
	cmdTmp.CommandText = 'SELECT DISTINCT NPA,NXX FROM xca_COCode Where NPA =?\' and Status = \'s\' Order By NXX';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NPANXXAvlRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_NPANXXAvlRec') != null)
		NPANXXAvlRec.setBookmark(thisPage.getState('pb_NPANXXAvlRec'));
}
function _NPANXXAvlRec_ctor()
{
	CreateRecordset('NPANXXAvlRec', _initNPANXXAvlRec, null);
}
function _NPANXXAvlRec_dtor()
{
	NPANXXAvlRec._preserveState();
	thisPage.setState('pb_NPANXXAvlRec', NPANXXAvlRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPALookUp style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sDISTINCT\sNPA\sFROM\sxca_COCode\sWhere\sStatus\s\l\g\s's'\sOrder\sBy\sNPA\q,TCControlID_Unmatched=\qNPALookUp\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sDISTINCT\sNPA\sFROM\sxca_COCode\sWhere\sStatus\s\l\g\s's'\sOrder\sBy\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNPALookUp()
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
	cmdTmp.CommandText = 'SELECT DISTINCT NPA FROM xca_COCode Where Status <> \'s\' Order By NPA';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NPALookUp.setRecordSource(rsTmp);
	if (thisPage.getState('pb_NPALookUp') != null)
		NPALookUp.setBookmark(thisPage.getState('pb_NPALookUp'));
}
function _NPALookUp_ctor()
{
	CreateRecordset('NPALookUp', _initNPALookUp, null);
}
function _NPALookUp_dtor()
{
	NPALookUp._preserveState();
	thisPage.setState('pb_NPALookUp', NPALookUp.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">

<div align="center"><center>
<table border="0" cellpadding="2"><tr>
  <td nowrap><font color="maroon" face="Arial Black" size="5"><strong>
 Generate All NPA Files</strong></font> 
		</td>
	</tr>
</table>
</center></div>

<p>&nbsp;</p>

<table align="center" border="0">
	<tr>
		<td align="right" nowrap>
            <p><font face="Arial" Color="black" Size="3"><strong>Generate All NPA 
            file with Utilized CO Codes:</strong><br></font>&nbsp;</p>
		</td><td>
		</td><td>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSUBMIT1 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSUBMIT1">
	<PARAM NAME="Caption" VALUE="SUBMIT">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnSUBMIT1()
{
	btnSUBMIT1.value = 'SUBMIT';
	btnSUBMIT1.setStyle(0);
}
function _btnSUBMIT1_ctor()
{
	CreateButton('btnSUBMIT1', _initbtnSUBMIT1, null);
}
</script>
<% btnSUBMIT1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>
	</tr>
	<tr>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td align="right" nowrap><font face="Arial" Color="black" Size="3"><strong>Generate All NPA file with Available CO Codes:</strong><br></font>
             
		</td><td>
		</td><td>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSUBMIT2 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSUBMIT2">
	<PARAM NAME="Caption" VALUE="SUBMIT">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnSUBMIT2()
{
	btnSUBMIT2.value = 'SUBMIT';
	btnSUBMIT2.setStyle(0);
}
function _btnSUBMIT2_ctor()
{
	CreateButton('btnSUBMIT2', _initbtnSUBMIT2, null);
}
</script>
<% btnSUBMIT2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</td>
	</tr>
	<tr>
		<td align="right" nowrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=Button2 style="HEIGHT: 27px; LEFT: 10px; TOP: 325px; WIDTH: 80px" 
	width=80>
	<PARAM NAME="_ExtentX" VALUE="2117">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button2">
	<PARAM NAME="Caption" VALUE="RETURN">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton2()
{
	Button2.value = 'RETURN';
	Button2.setStyle(0);
}
function _Button2_ctor()
{
	CreateButton('Button2', _initButton2, null);
}
</script>
<% Button2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
		<td>

</td>
		<td></td>
	</tr>
		
</table>


<p>&nbsp;</p></FORM>

</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

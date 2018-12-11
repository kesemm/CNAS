<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
</FORM>
<FORM action="xca_GenerateNPAFilesConfirm.asp" method=POST id=form1 name=form1>
<!--#include file="xca_CNASLib.inc"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</HEAD>

<BODY bgColor=#d7c7a4 bgProperties="fixed" text="black">
<P>
<%

Sub NPANXXUtilRec_onbeforeopen()
txtNPAUtil=Request.Form("txtNPAUtil")
NPAsql = "SELECT * FROM xca_COCode WHERE NPA = '"&txtNPAUtil&"' and Status <> 's' Order By NXX"
NPANXXUtilRec.setSQLText(NPAsql)
End Sub


Sub NPANXXAvlRec_onbeforeopen()
txtNPAAvl=Request.Form("txtNPAAvl")
NPAsql2 = "SELECT * FROM xca_COCode WHERE NPA = '"&txtNPAAvl&"' and Status = 's' Order By NXX"
NPANXXAvlRec.setSQLText(NPAsql2)
End Sub


Sub NPANXXUtilRec_ondatasetcomplete()

txtNPAUtil=NPANXXUtilRec.fields.getValue("NPA")
Dim fname, str,fileStr
	str=""


	
		str = str + "NPA" + chr(9) 
		str = str + "NXX" +  chr(13)+ chr(10) 
		Do While not NPANXXUtilRec.EOF 
		str = str + txtNPAUtil + chr(9) 
		str = str + NPANXXUtilRec.fields.getValue("NXX") +  chr(13)+ chr(10)
		fileStr = str
		NPANXXUtilRec.moveNext
		Loop


'Response.Write(txtNPAUtil)
'Response.Write("Entering CreateObject")
fname = "c:\xca_CNASNPAFiles\npaFileUtil_"+txtNPAUtil+".txt"
set fso = Server.CreateObject("Scripting.FileSystemObject")
set objNf = fso.CreateTextFile(fname, True)
	'Response.Write ("n syde  create")
	'set gf2 = fso.GetFile(fname)
	'Response.Write ("owt syde  create")
	'Response.Write ("gf
objNf.Write(fileStr)
objNf.Close
	'Response.Write("Leaveing CreateObject")
End Sub

Sub NPANXXAvlRec_ondatasetcomplete()

txtNPAAvl=NPANXXAvlRec.fields.getValue("NPA")
Dim fname, str,fileStr
	str=""

			str = str + "NPA" + chr(9) 
			str = str + "NXX" +  chr(13)+ chr(10)
			Do While not NPANXXAvlRec.EOF 
			str = str + txtNPAAvl + chr(9) 
			str = str + NPANXXAvlRec.fields.getValue("NXX") +  chr(13)+ chr(10)
			NPANXXAvlRec.moveNext		
			Loop
	fileStr = str
	fname = "c:\xca_CNASNPAFiles\npaFileAvl_"+txtNPAAvl+".txt"
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set objNf = fso.CreateTextFile(fname, True)
	set gf2 = fso.GetFile(fname)
	
	objNf.Write(fileStr)
	objNf.Close	
	End Sub
	%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPANXXUtilRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s\l\g\s's'\sOrder\sBy\sNXX\q,TCControlID_Unmatched=\qNPANXXUtilRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s\l\g\s's'\sOrder\sBy\sNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
	NPANXXUtilRec.open();
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPANXXAvlRec 
style="LEFT: 10px; TOP: 95px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s=\s's'\sOrder\sBy\sNXX\q,TCControlID_Unmatched=\qNPANXXAvlRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sFrom\sxca_COCodes\sWhere\sNPA=?\sand\sStaus\s=\s's'\sOrder\sBy\sNXX\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
	cmdTmp.CommandText = 'Select * From xca_COCodes Where NPA=? and Staus = \'s\' Order By NXX';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NPANXXAvlRec.setRecordSource(rsTmp);
	NPANXXAvlRec.open();
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
</P>
<P>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
Processing Form...</P>
<P>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%Response.Redirect "xca_GenerateNPAFilesConfirm.asp"
Display "<p>gf2</p>"
%>
</P>
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

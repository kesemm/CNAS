<%@ Language=VBScript %>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: ESRDInput.asp,v $
'* Commit Date:   $Date: 2016/03/15 14:54:10 $ (UTC)
'* Committed by:  $Author: walshkel $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name$ (Version/Build)
'**************************************************************************************** 
%>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">
</form>
<!--#include file="xca_CNASLib.inc"-->
<form action="ESRDPost.asp" method="post" id="ESRDInput" name="ESRDInput" onSubmit="return validateForm()">
<html>
<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<title>Part 1 - ESRD  (Required)</title>
<script LANGUAGE="JavaScript"> <!--

       
    function checkdate(a) {						
				var err=0,result
				if (a.length != 10) err=1
					d = a.substring(0, 2)//day  was-> b = a.substring(0, 2)// day
					c = a.substring(2, 3)// '/'
					b = a.substring(3, 5)//month was->d = a.substring(3, 5)// month
					e = a.substring(5, 6)// '/'
					f = a.substring(6, 10)// year
				if (b<1 || b>12) err = 1
				if (c != '/') err = 1
				if (d<1 || d>31) err = 1
				if (e != '/') err = 1
				if (f<1999) err = 1
				if (b==4 || b==6 || b==9 || b==11){
				if (d==31) err=1
				}
				if (b==2){
				var g=parseInt(f/4)
				if (isNaN(g)) {
				err=1
				}
				if (d>29) err=1
				if (d==29 && ((f/4)!=parseInt(f/4))) err=1
				}
				if (err==1) {
				return false;
				}
				else {
					return true;
			   }
		}  
 

 function validateForm()
        {    
getValues_JS(); 
           if (document.ESRDInput.AuthorizedRep.value == "") {
                alert("You have not filled in the Authorized Rep field. Please type in an Authorized Name and submit again");
                document.ESRDInput.AuthorizedRep.focus();               
                return false;
            }
            if (document.ESRDInput.AuthorizedRepTitle.value == "") {
                alert("You have not filled in the Authorized Rep Title field. Please type in an Authorized Name Title and submit again");
                document.ESRDInput.AuthorizedRepTitle.focus();
                return false;
            }

            if (document.ESRDInput.OCN.value == "") {
                alert("You have not filled in the OCN field. Please type in a 4 characters and submit again");
                document.ESRDInput.OCN.focus();
                return false;
            }

           if (document.ESRDInput.OCN.value.length <4) {
                alert("The OCN field must be 4 characters. Please type in a 4 characters and submit again");
                document.ESRDInput.OCN.focus();               
                return false; 
            }
            if (document.ESRDInput.ApplicationDate.value == "") {
                alert("You have not filled in the Application Date field. Please type in a valid date and submit again");
                document.ESRDInput.ApplicationDate.focus();
                return false;
            }
			
            var result=checkdate(document.ESRDInput.ApplicationDate.value) //this one             
            if (result==false)	{
				alert("The Application Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
                document.ESRDInput.ApplicationDate.focus();
                return false;
			}
			if (document.ESRDInput.CorrespondenceDate.value == "") {
                alert("You have not filled in the Correspondence Date field. Please type in a valid date and submit again");
                document.ESRDInput.CorrespondenceDate.focus();
                return false;
            }
			var result=checkdate(document.ESRDInput.CorrespondenceDate.value) //this one             
            if (result==false)	{
				alert("The Correspondence Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again");
                document.ESRDInput.CorrespondenceDate.focus();
                return false;
		}
}
        // end hiding -->
// app-b    
function getValues_JS(){
A=Number(document.ESRDInput.CurrentlyHeld.value);
B=Number(document.ESRDInput.CurrentlyUsed.value);
C=Number(document.ESRDInput.Required.value);
Additional=Math.ceil(((B+C)-(A*100))/100,1);
document.ESRDInput.Additional.value=Additional

}

</script>
 <meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%

'GET BOTH NPA-NXX SELECTIONS & the NXX Chosen
NPA511=Request.Form("NPA511")
NPA211=Request.Form("NPA211")
SelectedNXX=Request.Form("NXXSelect")

' Set this NXX to session
session("CurrentESRDNXX")=SelectedNXX

' Figure out which NXX to use
if SelectedNXX="211" Then
	SelectedNPA=NPA211
Else
	SelectedNPA=NPA511
end if

'' WRITE OUT VARIABLES FOR TESTING
' Response.Write("=================<br>")
' Response.Write("<STRONG>TESTING VARIABLES</STRONG><br>")
' Response.Write("--------------------------<br>")
' Response.Write("<br>")

' Response.Write("<STRONG>NPA511=</STRONG>"&NPA511&"<br><STRONG>NPA211=</STRONG>"&NPA211&"<br><STRONG>SelectedNXX=</STRONG>"&SelectedNXX&"<br><br><STRONG>SelectedNPA=</STRONG>"&SelectedNPA)

' Response.Write("<br>")
' Response.Write("=================<br>")


UserEntityID=session("UserEntityID")
uname = session("UserUserName")
Sub btnGoToMainFrm_onclick()
	Response.Redirect "xca_MenuNANPCAN.asp"
End Sub

		sqlcheckEntity="Select * from xca_Entity,xca_User where xca_Entity.EntityName='"&UserEntityID&"' "
			GetUserEntityName.setSQLText(sqlcheckEntity)
			GetUserEntityName.Open
			checkEntityName= GetUserEntityName.fields.getValue("EntityName")
			GetUserEntityName.close
sqlADMIN="Select * from xca_Entity where EntityName ='Admin'"
GetAdminEntityName.setSQLText(sqlADMIN)
GetAdminEntityName.Open
sql = "Select * from xca_Entity,xca_User where xca_Entity.EntityID = '"&UserEntityID&"' and xca_User.UserName= '"&uname&"' "
   GetUserEntityName.setSQLText(sql)
	GetUserEntityName.Open
session("P1CONPA")=SelectedNPA
' *** ADD NXX SELECTION TO THE WHERE CLAUSE BELOW
sql1 = "Select  ESRD from xca_ESRD where (Status='S' and (ReserveID='"&UserEntityID&"' or ReserveID=0) and NPA='"&SelectedNPA&"' and NXX='"&SelectedNXX&"')ORDER BY ESRD ASC"

   ESRDAssignLookup.setSQLText(sql1)
	 ESRDAssignLookup.Open
	ESRDAssignLookup.fields.getValue("ESRD"),ESRDAssignLookup.fields.getValue("ESRD"),cnt-1
 %>

</head>
<FORM>
<body leftmargin="20" rightmargin="20" bgColor="#d7c7a4" text="black" LANGUAGE=javascript>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=P1Parms 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCControlID_Unmatched=\qP1Parms\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\q\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sValue\sfrom\sxca_Parms\swhere\sName=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q25\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersP1Parms()
{
}
function _initP1Parms()
{
	P1Parms.advise(RS_ONBEFOREOPEN, _setParametersP1Parms);
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
	cmdTmp.CommandText = 'Select Value from xca_Parms where Name=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	P1Parms.setRecordSource(rsTmp);
	if (thisPage.getState('pb_P1Parms') != null)
		P1Parms.setBookmark(thisPage.getState('pb_P1Parms'));
}
function _P1Parms_ctor()
{
	CreateRecordset('P1Parms', _initP1Parms, null);
}
function _P1Parms_dtor()
{
	P1Parms._preserveState();
	thisPage.setState('pb_P1Parms', P1Parms.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetUserEntityName style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetUserEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetUserEntityName()
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
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetUserEntityName.setRecordSource(rsTmp);
}
function _GetUserEntityName_ctor()
{
	CreateRecordset('GetUserEntityName', _initGetUserEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetSelectedEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetSelectedEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=0,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetSelectedEntityName()
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
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
	rsTmp.CacheSize = 10;
	rsTmp.MaxRecords = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetSelectedEntityName.setRecordSource(rsTmp);
}
function _GetSelectedEntityName_ctor()
{
	CreateRecordset('GetSelectedEntityName', _initGetSelectedEntityName, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetAdminEntityName 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCControlID_Unmatched=\qGetAdminEntityName\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Entity\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_Entity\swhere\sxca_Entity.EntityName\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetAdminEntityName()
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
	cmdTmp.CommandText = 'Select * from xca_Entity where xca_Entity.EntityName =?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetAdminEntityName.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetAdminEntityName') != null)
		GetAdminEntityName.setBookmark(thisPage.getState('pb_GetAdminEntityName'));
}
function _GetAdminEntityName_ctor()
{
	CreateRecordset('GetAdminEntityName', _initGetAdminEntityName, null);
}
function _GetAdminEntityName_dtor()
{
	GetAdminEntityName._preserveState();
	thisPage.setState('pb_GetAdminEntityName', GetAdminEntityName.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=ESRDAssignLookup 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sESRD\sfrom\sxca_ESRD\swhere\sStatus='S'\sand\sNPA=?\q,TCControlID_Unmatched=\qESRDAssignLookup\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sESRD\sfrom\sxca_ESRD\swhere\sStatus='S'\sand\sNPA=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initESRDAssignLookup()
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
	cmdTmp.CommandText = 'Select Distinct ESRD from xca_ESRD where Status=\'S\' and NPA=?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	ESRDAssignLookup.setRecordSource(rsTmp);
	if (thisPage.getState('pb_ESRDAssignLookup') != null)
		ESRDAssignLookup.setBookmark(thisPage.getState('pb_ESRDAssignLookup'));
}
function _ESRDAssignLookup_ctor()
{
	CreateRecordset('ESRDAssignLookup', _initESRDAssignLookup, null);
}
function _ESRDAssignLookup_dtor()
{
	ESRDAssignLookup._preserveState();
	thisPage.setState('pb_ESRDAssignLookup', ESRDAssignLookup.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=GetESRDData style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sfrom\sxca_ESRD\swhere\sxcaESRD.Tix\s=?\q,TCControlID_Unmatched=\qGetESRDData\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_ESRD\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sfrom\sxca_ESRD\swhere\sxcaESRD.Tix\s=?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initGetESRDData()
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
	cmdTmp.CommandText = 'Select * from xca_ESRD where xca_ESRD.Tix =?';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	GetESRDData.setRecordSource(rsTmp);
	if (thisPage.getState('pb_GetESRDData') != null)
		GetESRDData.setBookmark(thisPage.getState('pb_GetESRDData'));
}
function _GetESRDData_ctor()
{
	CreateRecordset('GetESRDData', _initGetESRDData, null);
}
function _GetESRDData_dtor()
{
	GetESRDData._preserveState();
	thisPage.setState('pb_GetESRDData', GetESRDData.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->


<table border="0" cellpadding="0"><tr>
	<td wrap><font color="maroon" face="Arial Black" size="4"><strong>

	
	
	<br>
	
	Part 1 - Request for ESRD Block Assignment</strong></font>
            </td></tr>
            </table>
<font face="arial" size="2">

<p>Please complete the following form.  Use one form per OCN per NPA.  Multiple ESRD Blocks may be requested on a single form, provided these ESRD Blocks are for the same OCN and NPA.  Mail, fax or E-mail the completed form to the ESRD Block Administrator.</p>

<p>The assignment and use of ESRD Blocks is subject to the Canadian ESRD Block Assignment Guideline (the Guideline) which is available from the ESRD Block Administrator.  An ESRD Block assigned to an ESRD Block Applicant should be placed in use within 6 months of the assignment date, with possibility for a three month extension.</p>

<p>The Guideline may be modified from time-to-time.  The Guideline in effect shall apply equally to all ESRD Block Applicants and all existing ESRD Block Holders.</p>

<p>The ESRD Block Administrator will treat the information provided on this form as confidential in accordance with the Guideline.</p>

<p>I hereby certify that the information provided in this request for an ESRD Block is true and accurate to the best of my knowledge and that this application has been prepared in accordance with the current version of the Guideline.  I certify that the requested ESRD Block(s) will be used only for the purpose of providing wireless Enhanced 9-1-1 service in accordance with the Guideline and that all applicable regulatory authorization has been obtained.</p>

<p>I will return the ESRD Block to the ESRD Block Administrator if the resource is no longer in use by the ESRD Block Applicant, no longer required for the service for which it was intended, not in use within the time frame specified in the Guideline (including provision for extension), or not used in conformity with the Guideline.</p></font>
<p>
<br>
<table align="center" border="0" cellPadding="0" cellSpacing="0">
<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Authorized Representative 
            Name:&nbsp;&nbsp;</strong></font></label></td>
<td align="left" wrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=AuthorizedRep 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="AuthorizedRep">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetESRDData">
	<PARAM NAME="DataField" VALUE="AuthorizedRep">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRep()
{
	AuthorizedRep.setStyle(TXT_TEXTBOX);
	AuthorizedRep.setDataSource(GetESRDData);
	AuthorizedRep.setDataField('AuthorizedRep');
	AuthorizedRep.setMaxLength(35);
	AuthorizedRep.setColumnCount(35);
}
function _AuthorizedRep_ctor()
{
	CreateTextbox('AuthorizedRep', _initAuthorizedRep, null);
}
</script>
<% AuthorizedRep.display %>

<!--METADATA TYPE="DesignerControl" endspan-->          
</td></tr>
<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Title:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=AuthorizedRepTitle 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="AuthorizedRepTitle">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetESRDData">
	<PARAM NAME="DataField" VALUE="AuthorizedRepTitle">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAuthorizedRepTitle()
{
	AuthorizedRepTitle.setStyle(TXT_TEXTBOX);
	AuthorizedRepTitle.setDataSource(GetESRDData);
	AuthorizedRepTitle.setDataField('AuthorizedRepTitle');
	AuthorizedRepTitle.setMaxLength(35);
	AuthorizedRepTitle.setColumnCount(35);
}
function _AuthorizedRepTitle_ctor()
{
	CreateTextbox('AuthorizedRepTitle', _initAuthorizedRepTitle, null);
}
</script>
<% AuthorizedRepTitle.display %>

<!--METADATA TYPE="DesignerControl" endspan-->        
</td></tr>
<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Application Date:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=ApplicationDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="ApplicationDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetESRDData">
	<PARAM NAME="DataField" VALUE="ApplicationDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initApplicationDate()
{
	ApplicationDate.setStyle(TXT_TEXTBOX);
	ApplicationDate.setDataSource(GetESRDData);
	ApplicationDate.setDataField('ApplicationDate');
	ApplicationDate.setMaxLength(10);
	ApplicationDate.setColumnCount(10);
}
function _ApplicationDate_ctor()
{
	CreateTextbox('ApplicationDate', _initApplicationDate, null);
}
</script>
<% 
' ***********************************************************
' KT CHANGED 2013-06-14: 
'ApplicationDate.Value= FormatDateTime(now(),vbShortDate) 'GetESRDData.fields.getValue("ApplicationDate")                 'FormatDateTime(now(),vbShortDate)
'  I GUESS WE DON'T NEED TO FORMAT ANY DATE TEXT BECAUSE IT NEVER PUT ONE IN THERE!


ApplicationDate.display



 %>

<!--METADATA TYPE="DesignerControl" endspan-->
<font face="arial" size="1">dd/mm/ccyy</font></font></strong> 
     
</td></tr>

<tr>
<td align="right" wrap><label><font face="arial" size="2"><strong>Correspondence Date:&nbsp;&nbsp;</strong></font></label></td>
<td wrap align="left">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=CorrespondenceDate 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="CorrespondenceDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetESRDData">
	<PARAM NAME="DataField" VALUE="CorrespondenceDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCorrespondenceDate()
{
	CorrespondenceDate.setStyle(TXT_TEXTBOX);
	CorrespondenceDate.setDataSource(GetESRDData);
	CorrespondenceDate.setDataField('CorrespondenceDate');
	CorrespondenceDate.setMaxLength(10);
	CorrespondenceDate.setColumnCount(10);
}
function _CorrespondenceDate_ctor()
{
	CreateTextbox('CorrespondenceDate', _initCorrespondenceDate, null);
}
</script>
<% CorrespondenceDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<font face="arial" size="1">dd/mm/ccyy</font></font></strong> 
     
</td></tr>


</table>
<br><br>
<br><br>
<table align="left" border="0" cellPadding="0" cellSpacing="1">
<tr>
        <td wrap style="FONT-WEIGHT: bold"><label><strong><font size="3" face="arial" color="#993300">1.1 ESRD Block Applicant Contact Information:</font></strong></label> 
 
 </td></tr>
 
 </table>
 <br>
 <br>


<table align="center" border="0" cellPadding="1" cellSpacing="1">
    <tbody>
    
    <tr>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">Code Applicant 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
        <td align="left" wrap><font face="Arial"> </font>
        <td align="left" colSpan="2" wrap>
            <div align="center"><strong><u><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">CNA 
            Info:</font></u></strong></div><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font></td>
    </tr><tr> 
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity 
            Name</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityname 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityname">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityname()
{
	AppEntityname.setDataSource(GetUserEntityName);
	AppEntityname.setDataField('EntityName');
}
function _AppEntityname_ctor()
{
	CreateLabel('AppEntityname', _initAppEntityname, null);
}
</script>
<% AppEntityname.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>&nbsp;&nbsp;&nbsp;&nbsp;
        <td align="right" wrap> <font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Entity Name 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityName 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 65px" width=65>
	<PARAM NAME="_ExtentX" VALUE="1720">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityName">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityName()
{
	AdminEntityName.setDataSource(GetAdminEntityName);
	AdminEntityName.setDataField('EntityName');
}
function _AdminEntityName_ctor()
{
	CreateLabel('AdminEntityName', _initAdminEntityName, null);
}
</script>
<% AdminEntityName.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>

</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityContact">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityContact()
{
	AppEntityContact.setDataSource(GetUserEntityName);
	AppEntityContact.setDataField('UserName');
}
function _AppEntityContact_ctor()
{
	CreateLabel('AppEntityContact', _initAppEntityContact, null);
}
</script>
<% AppEntityContact.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Contact 
            Name</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityContact 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 76px" width=76>
	<PARAM NAME="_ExtentX" VALUE="2011">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityContact">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityContact">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityContact()
{
	AdminEntityContact.setDataSource(GetAdminEntityName);
	AdminEntityContact.setDataField('EntityContact');
}
function _AdminEntityContact_ctor()
{
	CreateLabel('AdminEntityContact', _initAdminEntityContact, null);
}
</script>
<% AdminEntityContact.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityAddress()
{
	AppEntityAddress.setDataSource(GetUserEntityName);
	AppEntityAddress.setDataField('UserAddress');
}
function _AppEntityAddress_ctor()
{
	CreateLabel('AppEntityAddress', _initAppEntityAddress, null);
}
</script>
<% AppEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Street 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityAddress 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityAddress">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityAddress">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityAddress()
{
	AdminEntityAddress.setDataSource(GetAdminEntityName);
	AdminEntityAddress.setDataField('EntityAddress');
}
function _AdminEntityAddress_ctor()
{
	CreateLabel('AdminEntityAddress', _initAdminEntityAddress, null);
}
</script>
<% AdminEntityAddress.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityCity">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityCity()
{
	AppEntityCity.setDataSource(GetUserEntityName);
	AppEntityCity.setDataField('UserCity');
}
function _AppEntityCity_ctor()
{
	CreateLabel('AppEntityCity', _initAppEntityCity, null);
}
</script>
<% AppEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">City 
            </font></font> 
            </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityCity 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 55px" width=55>
	<PARAM NAME="_ExtentX" VALUE="1455">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityCity">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityCity">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityCity()
{
	AdminEntityCity.setDataSource(GetAdminEntityName);
	AdminEntityCity.setDataField('EntityCity');
}
function _AdminEntityCity_ctor()
{
	CreateLabel('AdminEntityCity', _initAdminEntityCity, null);
}
</script>
<% AdminEntityCity.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityProvince()
{
	AppEntityProvince.setDataSource(GetUserEntityName);
	AppEntityProvince.setDataField('UserProvince');
}
function _AppEntityProvince_ctor()
{
	CreateLabel('AppEntityProvince', _initAppEntityProvince, null);
}
</script>
<% AppEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Province</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityProvince 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 82px" width=82>
	<PARAM NAME="_ExtentX" VALUE="2170">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityProvince">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityProvince">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityProvince()
{
	AdminEntityProvince.setDataSource(GetAdminEntityName);
	AdminEntityProvince.setDataField('EntityProvince');
}
function _AdminEntityProvince_ctor()
{
	CreateLabel('AdminEntityProvince', _initAdminEntityProvince, null);
}
</script>
<% AdminEntityProvince.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
         
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal 
            Code</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityPostalCode()
{
	AppEntityPostalCode.setDataSource(GetUserEntityName);
	AppEntityPostalCode.setDataField('UserPostalCode');
}
function _AppEntityPostalCode_ctor()
{
	CreateLabel('AppEntityPostalCode', _initAppEntityPostalCode, null);
}
</script>
<% AppEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font size="2"><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Postal Code 
            </font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityPostalCode 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 97px" width=97>
	<PARAM NAME="_ExtentX" VALUE="2566">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityPostalCode">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityPostalCode">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityPostalCode()
{
	AdminEntityPostalCode.setDataSource(GetAdminEntityName);
	AdminEntityPostalCode.setDataField('EntityPostalCode');
}
function _AdminEntityPostalCode_ctor()
{
	CreateLabel('AdminEntityPostalCode', _initAdminEntityPostalCode, null);
}
</script>
<% AdminEntityPostalCode.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
           
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail Address 
            </font></font> </font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"></font></font></td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 59px" width=59>
	<PARAM NAME="_ExtentX" VALUE="1561">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityEmail()
{
	AppEntityEmail.setDataSource(GetUserEntityName);
	AppEntityEmail.setDataField('UserEmail');
}
function _AppEntityEmail_ctor()
{
	CreateLabel('AppEntityEmail', _initAppEntityEmail, null);
}
</script>
<% AppEntityEmail.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">E-Mail 
            Address</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> 
            </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityEmail 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityEmail">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityEmail">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityEmail()
{
	AdminEntityEmail.setDataSource(GetAdminEntityName);
	AdminEntityEmail.setDataField('EntityEmail');
}
function _AdminEntityEmail_ctor()
{
	CreateLabel('AdminEntityEmail', _initAdminEntityEmail, null);
}
</script>
<% AdminEntityEmail.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>    
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 48px" width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityFax">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityFax()
{
	AppEntityFax.setDataSource(GetUserEntityName);
	AppEntityFax.setDataField('UserFax');
}
function _AppEntityFax_ctor()
{
	CreateLabel('AppEntityFax', _initAppEntityFax, null);
}
</script>
<% AppEntityFax.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Facsimile</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityFax 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 52px" width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityFax">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityFax">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityFax()
{
	AdminEntityFax.setDataSource(GetAdminEntityName);
	AdminEntityFax.setDataField('EntityFax');
}
function _AdminEntityFax_ctor()
{
	CreateLabel('AdminEntityFax', _initAdminEntityFax, null);
}
</script>
<% AdminEntityFax.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
          
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 89px" width=89>
	<PARAM NAME="_ExtentX" VALUE="2355">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityTelephone()
{
	AppEntityTelephone.setDataSource(GetUserEntityName);
	AppEntityTelephone.setDataField('UserTelephone');
}
function _AppEntityTelephone_ctor()
{
	CreateLabel('AppEntityTelephone', _initAppEntityTelephone, null);
}
</script>
<% AppEntityTelephone.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Telephone</font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityTelephone 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 93px" width=93>
	<PARAM NAME="_ExtentX" VALUE="2461">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityTelephone">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityTelephone">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityTelephone()
{
	AdminEntityTelephone.setDataSource(GetAdminEntityName);
	AdminEntityTelephone.setDataField('EntityTelephone');
}
function _AdminEntityTelephone_ctor()
{
	CreateLabel('AdminEntityTelephone', _initAdminEntityTelephone, null);
}
</script>
<% AdminEntityTelephone.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
            
            
</td></tr>
    <tr>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AppEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 84px" width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AppEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetUserEntityName">
	<PARAM NAME="DataField" VALUE="UserExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAppEntityExtension()
{
	AppEntityExtension.setDataSource(GetUserEntityName);
	AppEntityExtension.setDataField('UserExtension');
}
function _AppEntityExtension_ctor()
{
	CreateLabel('AppEntityExtension', _initAppEntityExtension, null);
}
</script>
<% AppEntityExtension.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
        </td>
        <td align="right" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4"></font>
        <td align="right" wrap><font face ="" size="2" 
           ><font face="Arial"><font style="BACKGROUND-COLOR: #d7c7a4">Extension</STRONG></font></font></font><font style="BACKGROUND-COLOR: #d7c7a4"><font face="Arial"> </font></font> </td>
        <td align="left" wrap><font face="Arial" style="BACKGROUND-COLOR: #d7c7a4">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=18 id=AdminEntityExtension 
	style="HEIGHT: 18px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="476">
	<PARAM NAME="id" VALUE="AdminEntityExtension">
	<PARAM NAME="DataSource" VALUE="GetAdminEntityName">
	<PARAM NAME="DataField" VALUE="EntityExtension">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT FACE="Arial"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initAdminEntityExtension()
{
	AdminEntityExtension.setDataSource(GetAdminEntityName);
	AdminEntityExtension.setDataField('EntityExtension');
}
function _AdminEntityExtension_ctor()
{
	CreateLabel('AdminEntityExtension', _initAdminEntityExtension, null);
}
</script>
<% AdminEntityExtension.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td></tr></tbody>
</table>
  
    <tr>
        <td align="left" colSpan="8"><strong><font face="arial" color="#993300" size="3">
	1.2 ESRD Block Request Information and Worksheet:</font></strong>
    <tr>
        <td align="right" colSpan="8">
            <div align="left">&nbsp; </div>
    <tr>
        <td align="left" colSpan="2" width="100"><strong><font face="arial" size="2">&nbsp;NPA:&nbsp; 
           <font color="blue">
            <%Response.Write(session("P1CONPA"))%></font>
			&nbsp;&nbsp;NXX:&nbsp; 
           <font color="blue">
            <%Response.Write(SelectedNXX)%></font></strong></FONT></td>
        <td align="left" colSpan="4" width="100"><strong><font face="arial" size="2">&nbsp;OCN:&nbsp;</font></strong>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=OCN style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 24px" 
	width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="OCN">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="GetESRDData">
	<PARAM NAME="DataField" VALUE="OCN">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="4">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initOCN()
{
	OCN.setStyle(TXT_TEXTBOX);
	OCN.setDataSource(GetESRDData);
	OCN.setDataField('OCN');
	OCN.setMaxLength(4);
	OCN.setColumnCount(4);
}
function _OCN_ctor()
{
	CreateTextbox('OCN', _initOCN, null);
}
</script>
<% OCN.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<p> </p>
<tr>
<td align="left" colSpan="3"><strong><font face="Arial" size="3" color="#993300">For this OCN in this NPA the Applicant must Provide:</font></strong></td></tr></table>
<p> </p>
<table border=0 background="">
<TR><TD>
<STRONG><FONT face=Arial size=2>
A.&nbsp; Quantity of ESRD Blocks currently held:</FONT></STRONG>&nbsp;</TD><TD>
<INPUT id=CurrentlyHeld name=CurrentlyHeld size=3  maxlength=3 value=0></TD> 
</TR>
<TR><TD>
<STRONG><FONT face=Arial size=2>
B.&nbsp; Quantity of ten-digit ESRDs currently in use:</FONT></STRONG>&nbsp;</TD><TD>
<INPUT id=CurrentlyUsed name=CurrentlyUsed size=5  maxlength=5 value=0></TD> 
</TR>
<TR><TD>
<STRONG><FONT face=Arial size=2>
C.&nbsp; Quantity of additional ten-digit ESRDs required:</FONT></STRONG>&nbsp;</TD><TD>
<INPUT id=Required name=Required size=5  maxlength=5 value=0> </TD> 
</TR>
<TR><TD>
<STRONG><FONT face=Arial size=2>
D.&nbsp; Quantity of additional ESRD Blocks required:</FONT></STRONG>&nbsp;</TD><TD>
</TR>
<TR><TD>
<FONT face=Arial size=2>&nbsp;(B+C-(A*100))/100 rounded up to the nearest integral =                         
<INPUT id=Additional name=Additional readonly size =3 maxlength=3 > 
</FONT>
                       
                        		</TD>
	</TR>
     <TR>

   <td align="left"><font face="Arial" size="2"><strong>
			Specific ESRB Block requested:</strong></font>
            
<%

ESRDAssignLookup.moveFirst
Response.Write"<SELECT id=ESRDAssign name=ESRDAssign>"
while (Not ESRDAssignLookup.EOF) 
Response.Write"<OPTION>"  
Response.Write ESRDAssignLookup.fields.getValue("ESRD") 
Response.Write"</OPTION>"
ESRDAssignLookup.moveNext
wend
Response.Write"</SELECT>"

%>
</SELECT>
</table>
<table>
        <TD align=left><STRONG><FONT face=Arial size=2>Supporting explanation (if desired):&nbsp;</FONT></STRONG>
<% 'Changed SupportingExplanation input box to 50 to match database and noted on screen /ktw 2016-03-15 %>                        
<INPUT id=SupportingExplanation name=SupportingExplanation size=65 maxlength=50 
                       > (50 chars)
                
		</TD>
	</TR>
</table>
<P>Notes:</P>
<table>
<TR><TD VAlign="top"><P>1)</P></TD>
<TD colspan="3"> <P> This worksheet is to be submitted to the ESRD Block Administrator with each request for an initial or additional ESRD Block.  For audit purposes, a copy of this worksheet must be retained in the ESRD Block Applicant’s files. </P> </TD> </TR>
<TR><TD><P>2)</P></TD>
<TD colspan="3"> <P> Definitions of terms may be found in the Glossary section of the Canadian ESRD Block Assignment Guideline.</P> </TD> </TR>
<TR><TD><P>3)</P></TD>
<TD colspan="3"> <P> An ESRD Block consists of 100 consecutive ten-digit ESRDs.  A ten-digit ESRD is in the form NPA-<%Response.Write(SelectedNXX)%>-XXXX.</P> </TD> </TR>
<TR><TD><P>4)</P></TD>
<TD colspan="3"> <P> Requests for specific ESRD Blocks will be considered at the discretion of the ESRD Block Administrator consistent with Section 3.4 of the Canadian ESRD Block Assignment Guideline.</P> </TD> </TR>

</table>

<P>&nbsp;<P></P>
		</TD>
	</tr>
    <TR>
        <TD align = left colSpan = 3>
    <TR>
        <TD align = left colSpan = 3>
    <TR>
        <TD align=left colSpan=3>
    <tr>
        <td align = "left" colSpan = "3">
    <tr>

		<td align = "left" colSpan = "3"> <input type="submit" value="Submit" name="submit">
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoToMainFrm 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoToMainFrm">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoToMainFrm()
{
	btnGoToMainFrm.value = 'Return';
	btnGoToMainFrm.setStyle(0);
}
function _btnGoToMainFrm_ctor()
{
	CreateButton('btnGoToMainFrm', _initbtnGoToMainFrm, null);
}
</script>
<% btnGoToMainFrm.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<tr>
<td align = "left" colSpan = "3" wrap>
	</td></tr></TBODY></TABLE></FORM>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

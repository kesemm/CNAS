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
<title>CO Codes Activity Log</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>
Function validDates (sdate,edate)
sd = mid(sdate,4,2) 'sd = start date
ed = mid(edate,4,2) 'ed = end date
'Response.Write str
	if isdate(sdate) and isdate(edate) then
	  if isnumeric(sd) and isnumeric(ed) then
	   if int(sd) < 13 and int(ed)< 13 then
		if DateValue(sdate) <= DateValue(edate)then
			validDates = true
		else 
			validDates = false
		end if
	   else 
		 validDates = false		
	   end if
	  else 
		 validDates = false
	  end if
    else
    validDates = false
'  Response.Write "false <br>"
end if	
end function

Function validDate (vdate)
str = mid(vdate,4,2)
'Response.Write str
if isdate(vdate)  then
  if isnumeric(str) then
    if int(str) < 13 then
		validDate = true
	else 
		validDate = false
	end if
  else 
	validDate = false		
  end if
	'   Response.Write "false <br>"
else 
	validDate = false
end if	
end function

Sub btnSubmit_onclick()
If NxxActLogRec.isOpen() then NxxActLogRec.close
''Dim MyDate, MyCheck
''txt1 = txtStartDate.getDataSource

MyDate = txtStartDate.value
'MyCheck=IsDate(MyDate)
MyDate1 = txtEndDate.value
'YourCheck=IsDate(MyDate1)
'	If MyCheck Then
'If  MyDate <> "" and MyDate1 <> "" Then 
'Response.Write "b 4  validate<br>"
  if validDates(MyDate,MyDate1) then
	sql3 = MyDate
	sql4 = MyDate1
	If sql3 <> "" And sql4 <> "" Then
		If sql3 = sql4 Then
			SQL9 = sql3 & "," & sql4
		ElseIf sql3 <> sql4 Then
			SQL9 = "Between" & sql3 & "," & sql4
		Else
			SQL9 = "" 
		End If
	End If
	
'Response.Write "after  validate<br>"
	TXT = ListSort5.selectedIndex
	TXT = ListSort5.getvalue (TXT)
	If TXT <> "" Then
		SQL1 = TXT
	Else
		SQL1 = ""
	End If
	TXT = ListSort6.selectedIndex
	TXT = ListSort6.getvalue (TXT)
	If TXT <> "" Then
		SQL2 = TXT
	Else
		SQL2 = ""
	End If

	fullstartdate=txtStartDate.value' & " 00:00:00"
	fullenddate=txtEndDate.value & " 23:59:59"
	RecSQL = "SELECT xca_Logs.*, xca_User.UserLogon  as Logon FROM xca_User, xca_Logs   WHERE (Date1 >= '"&fullstartdate&"' AND Date1 <= '"&fullenddate&"') AND LogType = 'C'and xca_User.UserID = xca_Logs.UserID" 
'	RecSQL = "SELECT * from xca_Logs  WHERE (Date1 >= '"&fullstartdate&"' AND Date1 <= '"&fullenddate&"') AND LogType = 'C'" 
'    RecSQL = "SELECT * FROM xca_Logs  Where (Date1 >=  '"&txt1&"' And Date1 <=  '"&txt2&"') and LogType = 'C'" 
	If SQL1 <> "" And SQL2 <> "" Then
		If SQL1 = SQL2 Then
			SQL = " Order By " & SQL1 
		Else
			SQL = " Order By " & SQL1 & "," & SQL2 
		End If
	ElseIF SQL1 <> "" Then
		SQL = " Order By " & SQL1 
	ElseIf SQL2 <> "" Then
		SQL = " Order By " & SQL2 
	Else
			SQL = "" 
	End If	
	
	RecSQL = RecSQL & SQL
	NxxActLogRec.setSQLText(RecSQL)
'	Response.Write (NxxActLogRec.getSQLText())
	NxxActLogRec.open
'  end if
'else
Session("Error") = ""
	ElseIF MyDate = "" and  MyDate1 = "" then session("Error")="Missing Date(s)"
	  ElseIF MyDate = "" then session("Error")="Start Date Needed"
		'Response.Write"<p align=center>The Start Date field must not be blank. Please type in a valid date (including leading zeros and 4 digit year) and submit again</p>"
		ElseIF not validDate(MyDate) then session("Error")="Wrong Date Format"
		'Response.Write"<p align=center>The Start Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again</p>"
			ElseIF MyDate1 = "" then session("Error")="End Date Needed"
			'Response.Write"<p align=center>The End Date field must not be blank. Please type in a valid date (including leading zeros and 4 digit year) and submit again</p>"
		  		ElseIF not validDate(MyDate1) then session("Error")="Wrong Date Format"
	  			'Response.Write"<p align=center>The End Date field is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again</p>"
					ElseIf DateValue(MyDate1) <= DateValue(MyDate)then session("Error")="Wrong Date Order"
					'Response.Write"<p align=center>The End Date field  is not greater than or equal to that of the Start Date . Please type in a valid End Date greater than the Start Date (including leading zeros and 4 digit year) and submit again</p>"
						
	Else
		session("Error")=""	
				' handle error here
'	Response.Write "<p align=center>The Start Date field or End Date  is invalid. Please type in a valid date (including leading zeros and 4 digit year) and submit again</p>"
	End if
End Sub


Sub btnClose_onclick()
	Response.Redirect "xca_RptAdminMenu.asp"
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor=#d7c7a4 bgProperties="fixed" text="black">
<% Select Case session("Error")  %>

<% Case "Wrong Date Format" %>
		<SCRIPT LANGUAGE="JavaScript">
		alert("Incorrect format of date field(s) entered.")
		</SCRIPT>

<% Case "Start Date Needed" %>
		<SCRIPT LANGUAGE="JavaScript">
		alert("Start date is missing.  Please enter a beginning date for the report.")
		</SCRIPT>

<% Case "Wrong Date Order" %>
		<SCRIPT LANGUAGE="JavaScript">
		alert("The End Date field  is not greater than or equal to that of the Start Date.")
		</SCRIPT>

<% Case "End Date Needed" %>
		<SCRIPT LANGUAGE="JavaScript">
		alert("End date is missing.  Please enter an ending date for the report.")
		</SCRIPT>

<% Case "Missing Date(s)" %>
		<SCRIPT LANGUAGE="JavaScript">
		alert("Missing date(s).  Please enter a Starting and Ending Date for this report.")
		</SCRIPT>

<% Case Else Session("Error") = "" %>

<% End Select %>

<%Session("Error") = "" %>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NxxActLogRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_Logs.*,\sxca_User.UserLogon\sAS\sLogon\sFROM\sxca_Logs,\sxca_User\q,TCControlID_Unmatched=\qNxxActLogRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Logs\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_Logs.*,\sxca_User.UserLogon\sAS\sLogon\sFROM\sxca_Logs,\sxca_User\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNxxActLogRec()
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
	cmdTmp.CommandText = 'SELECT xca_Logs.*, xca_User.UserLogon AS Logon FROM xca_Logs, xca_User';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NxxActLogRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_NxxActLogRec') != null)
		NxxActLogRec.setBookmark(thisPage.getState('pb_NxxActLogRec'));
}
function _NxxActLogRec_ctor()
{
	CreateRecordset('NxxActLogRec', _initNxxActLogRec, null);
}
function _NxxActLogRec_dtor()
{
	NxxActLogRec._preserveState();
	thisPage.setState('pb_NxxActLogRec', NxxActLogRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<div align="center"><center><table border="0" cellpadding="2"><tr>
	<td nowrap><font color=maroon face="Arial Black" size="5"><strong>
CO Code Activity 
Log</strong></font></td></tr></table></center></div><p></p>

<TABLE Align=right nowrap border=0 cellPadding=0 cellSpacing=0 height=23 style="HEIGHT: 23px; WIDTH: 226px" width=226>
	<TR>
		<TD><FONT face="Arial" size=4 color=black><STRONG>Created:
            <% Response.write "" & Date() %></STRONG></FONT>
        </TD></TR></TABLE>
		
		<p>&nbsp;</p>


<TABLE WIDTH=31.99% BGCOLOR=#d7c7a4  ALIGN=center BORDER=0 CELLSPACING=2 CELLPADDING=1 height=81 style="HEIGHT: 81px; WIDTH: 222px">
	<TR>
		<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label1 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 116px" 
            width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="First Order By:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('First Order By:');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD NOWRAP align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=ListSort5 style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 114px" 
            width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort5">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="8">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="NPA">
	<PARAM NAME="CLEV2" VALUE="NPA">
	<PARAM NAME="CLED3" VALUE="CO Code">
	<PARAM NAME="CLEV3" VALUE="NXX">
	<PARAM NAME="CLED4" VALUE="Logon ID">
	<PARAM NAME="CLEV4" VALUE="LOGON">
	<PARAM NAME="CLED5" VALUE="Log Date/Time">
	<PARAM NAME="CLEV5" VALUE="Date1">
	<PARAM NAME="CLED6" VALUE="Process">
	<PARAM NAME="CLEV6" VALUE="Process">
	<PARAM NAME="CLED7" VALUE="Action">
	<PARAM NAME="CLEV7" VALUE="Action">
	<PARAM NAME="CLED8" VALUE="ActionText">
	<PARAM NAME="CLEV8" VALUE="ActionText">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort5()
{
	ListSort5.addItem('', '');
	ListSort5.addItem('NPA', 'NPA');
	ListSort5.addItem('CO Code', 'NXX');
	ListSort5.addItem('Logon ID', 'LOGON');
	ListSort5.addItem('Log Date/Time', 'Date1');
	ListSort5.addItem('Process', 'Process');
	ListSort5.addItem('Action', 'Action');
	ListSort5.addItem('ActionText', 'ActionText');
}
function _ListSort5_ctor()
{
	CreateListbox('ListSort5', _initListSort5, null);
}
</script>
<% ListSort5.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                        
</TD>
		<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label2 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 141px" 
            width=141>
	<PARAM NAME="_ExtentX" VALUE="3731">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Second Order By:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setCaption('Second Order By:');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD NOWRAP align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=ListSort6 style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 114px" 
            width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort6">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="8">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="NPA">
	<PARAM NAME="CLEV2" VALUE="NPA">
	<PARAM NAME="CLED3" VALUE="CO Code">
	<PARAM NAME="CLEV3" VALUE="NXX">
	<PARAM NAME="CLED4" VALUE="Logon ID">
	<PARAM NAME="CLEV4" VALUE="LOGON">
	<PARAM NAME="CLED5" VALUE="Log Date/Time">
	<PARAM NAME="CLEV5" VALUE="Date1">
	<PARAM NAME="CLED6" VALUE="Process">
	<PARAM NAME="CLEV6" VALUE="Process">
	<PARAM NAME="CLED7" VALUE="Action">
	<PARAM NAME="CLEV7" VALUE="Action">
	<PARAM NAME="CLED8" VALUE="ActionText">
	<PARAM NAME="CLEV8" VALUE="ActionText">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort6()
{
	ListSort6.addItem('', '');
	ListSort6.addItem('NPA', 'NPA');
	ListSort6.addItem('CO Code', 'NXX');
	ListSort6.addItem('Logon ID', 'LOGON');
	ListSort6.addItem('Log Date/Time', 'Date1');
	ListSort6.addItem('Process', 'Process');
	ListSort6.addItem('Action', 'Action');
	ListSort6.addItem('ActionText', 'ActionText');
}
function _ListSort6_ctor()
{
	CreateListbox('ListSort6', _initListSort6, null);
}
</script>
<% ListSort6.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
	</TR>
		<TR>
		<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label3 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 85px" 
            width=85>
	<PARAM NAME="_ExtentX" VALUE="2249">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Start Date:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel3()
{
	Label3.setCaption('Start Date:');
}
function _Label3_ctor()
{
	CreateLabel('Label3', _initLabel3, null);
}
</script>
<% Label3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
		<TD NOWRAP align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtStartDate 
            style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtStartDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtStartDate()
{
	txtStartDate.setStyle(TXT_TEXTBOX);
	txtStartDate.setMaxLength(10);
	txtStartDate.setColumnCount(10);
}
function _txtStartDate_ctor()
{
	CreateTextbox('txtStartDate', _inittxtStartDate, null);
}
</script>
<% txtStartDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->               
		</TD>
	<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label4 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 80px" 
            width=80>
	<PARAM NAME="_ExtentX" VALUE="2117">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label4">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="End Date:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel4()
{
	Label4.setCaption('End Date:');
}
function _Label4_ctor()
{
	CreateLabel('Label4', _initLabel4, null);
}
</script>
<% Label4.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
		</TD>
		<TD NOWRAP align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=txtEndDate style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 60px" 
            width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtEndDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtEndDate()
{
	txtEndDate.setStyle(TXT_TEXTBOX);
	txtEndDate.setMaxLength(10);
	txtEndDate.setColumnCount(10);
}
function _txtEndDate_ctor()
{
	CreateTextbox('txtEndDate', _inittxtEndDate, null);
}
</script>
<% txtEndDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
		</TD>
	</TR>
	<TR>
		<TD></TD>
		<TD nowrap align=left><font face=Arial size=2 color=black>(dd/mm/ccyy)</font></TD>
		<TD></TD>
		<TD nowrap align=left><font face=Arial size=2 color=black>(dd/mm/ccyy)</font></TD>
	</TR>
</TABLE>
<TABLE WIDTH=138 BGCOLOR=#d7c7a4 ALIGN=center BORDER=0 CELLSPACING=2 CELLPADDING=1 height=20 style="HEIGHT: 20px; WIDTH: 138px" background="">
    
    
	 <TR>
        <TD noWrap align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSubmit 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 285px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSubmit">
	<PARAM NAME="Caption" VALUE="Submit">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnSubmit()
{
	btnSubmit.value = 'Submit';
	btnSubmit.setStyle(0);
}
function _btnSubmit_ctor()
{
	CreateButton('btnSubmit', _initbtnSubmit, null);
}
</script>
<% btnSubmit.display %>

<!--METADATA TYPE="DesignerControl" endspan-->           
		</TD>
        <TD noWrap align=left>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnClose 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 312px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnClose">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
    </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnClose()
{
	btnClose.value = 'Return';
	btnClose.setStyle(0);
}
function _btnClose_ctor()
{
	CreateButton('btnClose', _initbtnClose, null);
}
</script>
<% btnClose.display %>

<!--METADATA TYPE="DesignerControl" endspan-->           
</TD></TR>
</TABLE>


<P>&nbsp;</P>

<TABLE WIDTH=75% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=absmiddle NOWRAP>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" 
            height=147 id=Grid1 
            style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 692px" width=692>
	<PARAM NAME="_ExtentX" VALUE="18309">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="NxxActLogRec">
	<PARAM NAME="CtrlName" VALUE="Grid1">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="ffff25">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="0">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="0">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPA","NXX","LOGON","Date1","Process","Action","ActionText"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4,5,6">
	<PARAM NAME="displayWidth" VALUE="65,80,91,150,75,80,175">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPA","CO Code","Logon ID","Log Date/Time","Process","Action","ActionText"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,,,">
	<PARAM NAME="ColumnCount" VALUE="7">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="2">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="2">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="9999">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="692">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="452653">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 9999;
Grid1.setDataSource(NxxActLogRec);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=7 rules=ALL WIDTH=692';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=65';
Grid1.headerWidth[1] = ' WIDTH=80';
Grid1.headerWidth[2] = ' WIDTH=91';
Grid1.headerWidth[3] = ' WIDTH=150';
Grid1.headerWidth[4] = ' WIDTH=75';
Grid1.headerWidth[5] = ' WIDTH=80';
Grid1.headerWidth[6] = ' WIDTH=175';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPA\'';
Grid1.colHeader[1] = '\'CO Code\'';
Grid1.colHeader[2] = '\'Logon ID\'';
Grid1.colHeader[3] = '\'Log Date/Time\'';
Grid1.colHeader[4] = '\'Process\'';
Grid1.colHeader[5] = '\'Action\'';
Grid1.colHeader[6] = '\'ActionText\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Center bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Center bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=65';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'NxxActLogRec.fields.getValue(\'NPA\')';
Grid1.colAttributes[1] = '  WIDTH=80';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'NxxActLogRec.fields.getValue(\'NXX\')';
Grid1.colAttributes[2] = '  WIDTH=91';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'NxxActLogRec.fields.getValue(\'LOGON\')';
Grid1.colAttributes[3] = '  WIDTH=150';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'NxxActLogRec.fields.getValue(\'Date1\')';
Grid1.colAttributes[4] = '  WIDTH=75';
Grid1.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[4] = 'NxxActLogRec.fields.getValue(\'Process\')';
Grid1.colAttributes[5] = '  WIDTH=80';
Grid1.colFormat[5] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[5] = 'NxxActLogRec.fields.getValue(\'Action\')';
Grid1.colAttributes[6] = '  WIDTH=175';
Grid1.colFormat[6] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[6] = 'NxxActLogRec.fields.getValue(\'ActionText\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(0,1);
Grid1.hasPageNumber = true;
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->

</TD>
	</TR>
</TABLE>


</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

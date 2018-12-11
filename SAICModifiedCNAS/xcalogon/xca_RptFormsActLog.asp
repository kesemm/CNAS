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
<title>Request Forms Activity Log</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>
Function validDates (sdate,edate)
sd = mid(sdate,4,2) 'sd = start date
ed = mid(edate,4,2) 'ed = end date
'Response.Write str
	if isdatereal(sdate) and isdatereal(edate) then
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
if isdatereal(vdate)  then
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
if COCRqtFrmRec.isopen() then COCRqtFrmRec.close()

	'Dim MyDate, MyCheck
	'txt1 = txtStartDate.getDataSource
	MyDate=txtStartDate.value
	MyDate1=txtEndDate.value
'	MyCheck=isdatereal(MyDate)
	if validDates(MyDate,MyDate1) then
		TXT = ListSort3.selectedIndex
		TXT = ListSort3.getvalue(TXT)
		If TXT <> "" Then
			SQL1 = TXT
		Else
			SQL1 = ""
		End If


		TXT = ListSort4.selectedIndex
		TXT = ListSort4.getvalue(TXT)
		If TXT <> "" Then
			SQL2 = TXT
		Else
			SQL2 = ""
		End If

		fullstartdate=txtStartDate.value '& " 00:00:00"
		fullenddate=txtEndDate.value & " 23:59:59"
		RecSQL = "SELECT xca_Logs.*, xca_User.UserLogon AS Logon FROM xca_Logs, xca_User   WHERE (Date1 >= '"&fullstartdate&"' AND Date1 <= '"&fullenddate&"') AND LogType = 'R' and xca_Logs.UserID = xca_User.UserID"
'		RecSQL = "SELECT * FROM xca_Logs WHERE (Date1 >= '"&fullstartdate&"' AND Date1 <= '"&fullenddate&"') AND LogType = 'R'" 
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
		COCRqtFrmRec.setSQLText (RecSQL)
		COCRqtFrmRec.open
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


Sub COCRqtFrmRec_ondatasetcomplete()
	'Response.Redirect "xca_RptFormsActLog.asp"
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=COCRqtFrmRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\s\sFROM\sxca_Logs\s\s\s\q,TCControlID_Unmatched=\qCOCRqtFrmRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\s\sFROM\sxca_Logs\s\s\s\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCOCRqtFrmRec()
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
	cmdTmp.CommandText = 'SELECT *  FROM xca_Logs   ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	COCRqtFrmRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_COCRqtFrmRec') != null)
		COCRqtFrmRec.setBookmark(thisPage.getState('pb_COCRqtFrmRec'));
}
function _COCRqtFrmRec_ctor()
{
	CreateRecordset('COCRqtFrmRec', _initCOCRqtFrmRec, null);
}
function _COCRqtFrmRec_dtor()
{
	COCRqtFrmRec._preserveState();
	thisPage.setState('pb_COCRqtFrmRec', COCRqtFrmRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->


<div align="center"><center>
<table border="0" cellpadding="2">
	<tr>
		<td nowrap><font color=maroon face="Arial Black" size="5"><strong>
Request Forms Activity Log</strong></font> 
		</td>
	</tr>
</table>
</center></div>

<p>&nbsp;</p>

<TABLE Align=right nowrap border=0 cellPadding=0 cellSpacing=0 height=23 style="HEIGHT: 23px; WIDTH: 226px" width=226>
	<TR>
		<TD><FONT face="Arial" size=4 color=black><STRONG>Created
            <% Response.write "" & Date() %></STRONG></FONT>
        </TD>
	</TR>
</TABLE>
		
		<p>&nbsp;</p>


<TABLE WIDTH=31.99% BGCOLOR=#d7c7a4  ALIGN=center BORDER=0 CELLSPACING=2 CELLPADDING=1 height=81 style="HEIGHT: 81px; WIDTH: 222px">
	<TR>
		<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=26 
            id=Label1 style="HEIGHT: 26px; LEFT: 0px; TOP: 0px; WIDTH: 127px" 
            width=127>
	<PARAM NAME="_ExtentX" VALUE="3360">
	<PARAM NAME="_ExtentY" VALUE="688">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="First Sort By:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="4">
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
<FONT FACE="Arial" SIZE="4" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('First Sort By:');
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
            id=ListSort3 style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 114px" 
            width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="9">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="Ticket #">
	<PARAM NAME="CLEV2" VALUE="Tix">
	<PARAM NAME="CLED3" VALUE="Logon ID">
	<PARAM NAME="CLEV3" VALUE="Logon">
	<PARAM NAME="CLED4" VALUE="Log Date/Time">
	<PARAM NAME="CLEV4" VALUE="Date1">
	<PARAM NAME="CLED5" VALUE="NPA">
	<PARAM NAME="CLEV5" VALUE="NPA">
	<PARAM NAME="CLED6" VALUE="CO Code">
	<PARAM NAME="CLEV6" VALUE="NXX">
	<PARAM NAME="CLED7" VALUE="Process">
	<PARAM NAME="CLEV7" VALUE="Process">
	<PARAM NAME="CLED8" VALUE="Action">
	<PARAM NAME="CLEV8" VALUE="Action">
	<PARAM NAME="CLED9" VALUE="ActionText">
	<PARAM NAME="CLEV9" VALUE="ActionText">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort3()
{
	ListSort3.addItem('', '');
	ListSort3.addItem('Ticket #', 'Tix');
	ListSort3.addItem('Logon ID', 'Logon');
	ListSort3.addItem('Log Date/Time', 'Date1');
	ListSort3.addItem('NPA', 'NPA');
	ListSort3.addItem('CO Code', 'NXX');
	ListSort3.addItem('Process', 'Process');
	ListSort3.addItem('Action', 'Action');
	ListSort3.addItem('ActionText', 'ActionText');
}
function _ListSort3_ctor()
{
	CreateListbox('ListSort3', _initListSort3, null);
}
</script>
<% ListSort3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
</TD>
<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=26 
            id=Label2 style="HEIGHT: 26px; LEFT: 0px; TOP: 0px; WIDTH: 156px" 
            width=156>
	<PARAM NAME="_ExtentX" VALUE="4128">
	<PARAM NAME="_ExtentY" VALUE="688">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Second Sort By:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="4" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setCaption('Second Sort By:');
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
            id=ListSort4 style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 114px" 
            width=114>
	<PARAM NAME="_ExtentX" VALUE="3016">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort4">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="9">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="Ticket #">
	<PARAM NAME="CLEV2" VALUE="Tix">
	<PARAM NAME="CLED3" VALUE="Logon ID">
	<PARAM NAME="CLEV3" VALUE="Logon">
	<PARAM NAME="CLED4" VALUE="Log Date/Time">
	<PARAM NAME="CLEV4" VALUE="Date1">
	<PARAM NAME="CLED5" VALUE="NPA">
	<PARAM NAME="CLEV5" VALUE="NPA">
	<PARAM NAME="CLED6" VALUE="CO Code">
	<PARAM NAME="CLEV6" VALUE="NXX">
	<PARAM NAME="CLED7" VALUE="Process">
	<PARAM NAME="CLEV7" VALUE="Process">
	<PARAM NAME="CLED8" VALUE="Action">
	<PARAM NAME="CLEV8" VALUE="Action">
	<PARAM NAME="CLED9" VALUE="ActionText">
	<PARAM NAME="CLEV9" VALUE="ActionText">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort4()
{
	ListSort4.addItem('', '');
	ListSort4.addItem('Ticket #', 'Tix');
	ListSort4.addItem('Logon ID', 'Logon');
	ListSort4.addItem('Log Date/Time', 'Date1');
	ListSort4.addItem('NPA', 'NPA');
	ListSort4.addItem('CO Code', 'NXX');
	ListSort4.addItem('Process', 'Process');
	ListSort4.addItem('Action', 'Action');
	ListSort4.addItem('ActionText', 'ActionText');
}
function _ListSort4_ctor()
{
	CreateListbox('ListSort4', _initListSort4, null);
}
</script>
<% ListSort4.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</TD>
	</TR>
		<TR>
		<TD NOWRAP align=right>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=26 
            id=StartDate style="HEIGHT: 26px; LEFT: 0px; TOP: 0px; WIDTH: 102px" 
            width=102>
	<PARAM NAME="_ExtentX" VALUE="2699">
	<PARAM NAME="_ExtentY" VALUE="688">
	<PARAM NAME="id" VALUE="StartDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Start Date:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="4" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initStartDate()
{
	StartDate.setCaption('Start Date:');
}
function _StartDate_ctor()
{
	CreateLabel('StartDate', _initStartDate, null);
}
</script>
<% StartDate.display %>
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
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=26 
            id=EndDate style="HEIGHT: 26px; LEFT: 0px; TOP: 0px; WIDTH: 95px" 
            width=95>
	<PARAM NAME="_ExtentX" VALUE="2514">
	<PARAM NAME="_ExtentY" VALUE="688">
	<PARAM NAME="id" VALUE="EndDate">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="End Date:">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<FONT FACE="Arial" SIZE="4" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEndDate()
{
	EndDate.setCaption('End Date:');
}
function _EndDate_ctor()
{
	CreateLabel('EndDate', _initEndDate, null);
}
</script>
<% EndDate.display %>
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
<TABLE  ALIGN=center BORDER=0 CELLSPACING=2 CELLPADDING=1 background="">
     <TR>
        <TD noWrap align=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSubmit 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 297px; WIDTH: 63px" width=63>
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
	<TD noWrap align=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnClose 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 324px; WIDTH: 61px" width=61>
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

<TABLE align=center nowrap border=0 cellPadding=0 cellSpacing=0 height=149 style="HEIGHT: 149px; WIDTH: 703px" width=703>
<TR>
	<TD noWrap align=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" 
            height=147 id=COCALogRtpGrd 
            style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 880px" width=880>
	<PARAM NAME="_ExtentX" VALUE="23283">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="COCRqtFrmRec">
	<PARAM NAME="CtrlName" VALUE="COCALogRtpGrd">
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
	<PARAM NAME="HiliteColor" VALUE="#ffff25">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"Tix","Logon","Date1","NPA","NXX","Process","Action","ActionText"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4,5,6,7">
	<PARAM NAME="displayWidth" VALUE="90,85,150,70,90,80,90,200">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Ticket #","Logon ID","Log Date/Time","NPA","CO Code","Process","Action","Action Text"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,,,,">
	<PARAM NAME="ColumnCount" VALUE="8">
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
	<PARAM NAME="BorderSize" VALUE="0">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="880">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="436269">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCOCALogRtpGrd()
{
COCALogRtpGrd.pageSize = 9999;
COCALogRtpGrd.setDataSource(COCRqtFrmRec);
COCALogRtpGrd.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=0 cols=8 rules=ALL WIDTH=880 nowrap';
COCALogRtpGrd.headerAttributes = '   bgcolor=Maroon align=Center nowrap';
COCALogRtpGrd.headerWidth[0] = ' WIDTH=90';
COCALogRtpGrd.headerWidth[1] = ' WIDTH=85';
COCALogRtpGrd.headerWidth[2] = ' WIDTH=150';
COCALogRtpGrd.headerWidth[3] = ' WIDTH=70';
COCALogRtpGrd.headerWidth[4] = ' WIDTH=90';
COCALogRtpGrd.headerWidth[5] = ' WIDTH=80';
COCALogRtpGrd.headerWidth[6] = ' WIDTH=90';
COCALogRtpGrd.headerWidth[7] = ' WIDTH=200';
COCALogRtpGrd.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
COCALogRtpGrd.colHeader[0] = '\'Ticket #\'';
COCALogRtpGrd.colHeader[1] = '\'Logon ID\'';
COCALogRtpGrd.colHeader[2] = '\'Log Date/Time\'';
COCALogRtpGrd.colHeader[3] = '\'NPA\'';
COCALogRtpGrd.colHeader[4] = '\'CO Code\'';
COCALogRtpGrd.colHeader[5] = '\'Process\'';
COCALogRtpGrd.colHeader[6] = '\'Action\'';
COCALogRtpGrd.colHeader[7] = '\'Action Text\'';
COCALogRtpGrd.rowAttributes[0] = '  bgcolor = Silver align=Center nowrap  bordercolor=White';
COCALogRtpGrd.rowAttributes[1] = '  bgcolor = White align=Center nowrap  bordercolor=White';
COCALogRtpGrd.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
COCALogRtpGrd.colAttributes[0] = '  WIDTH=90';
COCALogRtpGrd.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[0] = 'COCRqtFrmRec.fields.getValue(\'Tix\')';
COCALogRtpGrd.colAttributes[1] = '  WIDTH=85';
COCALogRtpGrd.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[1] = 'COCRqtFrmRec.fields.getValue(\'Logon\')';
COCALogRtpGrd.colAttributes[2] = '  WIDTH=150';
COCALogRtpGrd.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[2] = 'COCRqtFrmRec.fields.getValue(\'Date1\')';
COCALogRtpGrd.colAttributes[3] = '  WIDTH=70';
COCALogRtpGrd.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[3] = 'COCRqtFrmRec.fields.getValue(\'NPA\')';
COCALogRtpGrd.colAttributes[4] = '  WIDTH=90';
COCALogRtpGrd.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[4] = 'COCRqtFrmRec.fields.getValue(\'NXX\')';
COCALogRtpGrd.colAttributes[5] = '  WIDTH=80';
COCALogRtpGrd.colFormat[5] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[5] = 'COCRqtFrmRec.fields.getValue(\'Process\')';
COCALogRtpGrd.colAttributes[6] = '  WIDTH=90';
COCALogRtpGrd.colFormat[6] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[6] = 'COCRqtFrmRec.fields.getValue(\'Action\')';
COCALogRtpGrd.colAttributes[7] = '  WIDTH=200';
COCALogRtpGrd.colFormat[7] = '<Font Size=2 Face="Arial" Color=Black >';
COCALogRtpGrd.colData[7] = 'COCRqtFrmRec.fields.getValue(\'ActionText\')';
COCALogRtpGrd.navbarAlignment = 'Right';
var objPageNavbar = COCALogRtpGrd.showPageNavbar(0,1);
COCALogRtpGrd.hasPageNumber = true;
}
function _COCALogRtpGrd_ctor()
{
	CreateDataGrid('COCALogRtpGrd',_initCOCALogRtpGrd);
}
</SCRIPT>

<%	COCALogRtpGrd.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
	</TD>
</TR>
</TABLE>

<p>&nbsp;</p>



</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

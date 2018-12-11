<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires =0

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
'Session("Pending")="xca_Pending.asp"
Session("Here")="xca_Pending.asp"
Dim DateresponseDue

Sub btnGoToTicket_onclick()
session("Tix")=P3gotoTix.Value
Response.Redirect "xca_Part3.asp"
End Sub

if session("NoTixSent")="DidNotSend" then 
	%><script LANGUAGE="JavaScript">
alert("That Ticket does not exist or is closed.  Please try again.....");

	</Script>
<%
end if
session("NoTixSent")=""




Sub btnSubmit_onclick()

if COCodeRec.isopen() then COCodeRec.close()
	ValueSort=PendingSort.selectedIndex
	ValueSort=PendingSort.getvalue(ValueSort)
	
	ValueFilter=PendingFilter.selectedIndex
	ValueFilter=PendingFilter.getvalue(ValueFilter)
	
'	if ValueSort="" then
'		ValueSort="Tix"
'	end if
	if ValueSort <> "" then
		SQL1 = ValueSort
	else
		SQL1 = ""
	end if
	
	
	select case ValueFilter
	case "" 
	SQLfilter="RequestStatus='NW' or RequestStatus='RS' or RequestStatus='AS' or RequestStatus='UP'	"
	case "AL"
	SQLfilter="RequestStatus='NW' or RequestStatus='RS' or RequestStatus='AS' or RequestStatus='UP'	"
	case "AA"
	SQLfilter="RequestStatus='AS'"
	case "RS"
	SQLfilter="RequestStatus='RS'"
	case "NW"
	SQLfilter="RequestStatus='NW'"
	case "UU"
	SQLfilter="RequestStatus='UP'"
	end select
	
'ReqSQL = "Select xca_Part1.* From xca_Part1 where " & SQLfilter 
RecSQL = "select xca_Part1.* "	
RecSQL= RecSQL & "from xca_Part1 where "
RecSQL= RecSQL & SQLfilter 
If SQL1 <> "" Then
	SQL =" ORDER BY " & SQL1
Else
	SQL = ""
End If

RecSQL= RecSQL & SQL
'RecSQL= RecSQL &" order by " & ValueSort
	COCodeRec.setsqltext(RecSQL)
	COCodeRec.open
	
	if COCodeRec.getCount()=0 then  %>

	<SCRIPT Language="javascript">
	alert("No Values Exist from your selection criteria  Please try again.")
	</SCRIPT>
<%	
	'give them everything
	RecSQL = "select * from xca_Part1 where (RequestStatus='NW' or RequestStatus='AS' or RequestStatus='RS' or RequestStatus='UP') order by Tix"
	COCodeRec.setsqltext(RecSQL)
	COCodeRec.open

			
	End if
End Sub

%>

</HEAD>
<BODY bgColor=#d7c7a4 bgProperties="fixed" text="black">
<table align=center border="0" cellpadding="0">
<tr>
	<td nowrap align=center colSpan=3><font color="maroon" face="Arial Black" size="5"><strong>
Pending Requests</strong></font></td>
	<tr>
		<td colspan=3></td>
	</tr>
	<tr>
		<td colspan=3></td>
	</tr>
	<tr>
		<td colspan=3></td>
	</tr>
    <TR>
        <TD><FONT face=Arial><STRONG>Column Sort</STRONG></FONT>
        <TD><FONT size=+0><STRONG><FONT size=+0><FONT 
            face=Arial><FONT>Part 1 Status 
            Filter&nbsp;</FONT></FONT></FONT></FONT><FONT><FONT face=Arial><FONT 
            size=+0> </FONT></FONT></FONT></STRONG>
            
            
        <TD>
    <TR>
        <TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=PendingSort 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 146px" width=146>
	<PARAM NAME="_ExtentX" VALUE="3863">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="PendingSort">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="3">
	<PARAM NAME="CLED1" VALUE="Date Response Due">
	<PARAM NAME="CLEV1" VALUE="DueDate">
	<PARAM NAME="CLED2" VALUE="Part 1 Request Status">
	<PARAM NAME="CLEV2" VALUE="RequestStatus">
	<PARAM NAME="CLED3" VALUE="Ticket">
	<PARAM NAME="CLEV3" VALUE="Tix">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPendingSort()
{
	PendingSort.addItem('Date Response Due', 'DueDate');
	PendingSort.addItem('Part 1 Request Status', 'RequestStatus');
	PendingSort.addItem('Ticket', 'Tix');
}
function _PendingSort_ctor()
{
	CreateListbox('PendingSort', _initPendingSort, null);
}
</script>
<% PendingSort.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        </TD><TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=PendingFilter 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 88px" width=88>
	<PARAM NAME="_ExtentX" VALUE="2328">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="PendingFilter">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="5">
	<PARAM NAME="CLED1" VALUE="New">
	<PARAM NAME="CLEV1" VALUE="NW">
	<PARAM NAME="CLED2" VALUE="Reserved">
	<PARAM NAME="CLEV2" VALUE="RS">
	<PARAM NAME="CLED3" VALUE="Update">
	<PARAM NAME="CLEV3" VALUE="UU">
	<PARAM NAME="CLED4" VALUE="Assigned">
	<PARAM NAME="CLEV4" VALUE="AA">
	<PARAM NAME="CLED5" VALUE="All">
	<PARAM NAME="CLEV5" VALUE="AL">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initPendingFilter()
{
	PendingFilter.addItem('New', 'NW');
	PendingFilter.addItem('Reserved', 'RS');
	PendingFilter.addItem('Update', 'UU');
	PendingFilter.addItem('Assigned', 'AA');
	PendingFilter.addItem('All', 'AL');
}
function _PendingFilter_ctor()
{
	CreateListbox('PendingFilter', _initPendingFilter, null);
}
</script>
<% PendingFilter.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
 </TD><TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSubmit style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 63px" 
            width=63>
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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            
    <TR>
        <TD colSpan=3><FONT face="Arial Narrow">Select 
            the Column Sort and Part 1 Staus Filter criteria above and click SUBMIT 
            to display the results.&nbsp; 
</FONT>
    <TR>
        <TD colSpan=3><FONT face="Arial Narrow">Then, Enter the Ticket # 
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=P3gotoTix 
            style="HEIGHT: 19px; LEFT: 145px; TOP: 8px; WIDTH: 54px" width=54><PARAM NAME="_ExtentX" VALUE="1429"><PARAM NAME="_ExtentY" VALUE="503"><PARAM NAME="id" VALUE="P3gotoTix"><PARAM NAME="ControlType" VALUE="0"><PARAM NAME="Lines" VALUE="3"><PARAM NAME="DataSource" VALUE=""><PARAM NAME="DataField" VALUE=""><PARAM NAME="Enabled" VALUE="-1"><PARAM NAME="Visible" VALUE="-1"><PARAM NAME="MaxChars" VALUE="9"><PARAM NAME="DisplayWidth" VALUE="9"><PARAM NAME="Platform" VALUE="256"><PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initP3gotoTix()
{
	P3gotoTix.setStyle(TXT_TEXTBOX);
	P3gotoTix.setMaxLength(9);
	P3gotoTix.setColumnCount(9);
}
function _P3gotoTix_ctor()
{
	CreateTextbox('P3gotoTix', _initP3gotoTix, null);
}
</script>
<% P3gotoTix.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
you need to update and 
            then click
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnGoToTicket 
            style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 36px" width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoToTicket">
	<PARAM NAME="Caption" VALUE="Go">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoToTicket()
{
	btnGoToTicket.value = 'Go';
	btnGoToTicket.setStyle(0);
}
function _btnGoToTicket_ctor()
{
	CreateButton('btnGoToTicket', _initbtnGoToTicket, null);
}
</script>
<% btnGoToTicket.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></TR>

</table>

<TABLE align=center nowrap border=0 cellPadding=0 cellSpacing=0 height=149 style="HEIGHT: 149px; WIDTH: 703px" width=703>
<TR>
	<TD noWrap align=middle><FONT>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 400px" 
	width=400>
	<PARAM NAME="_ExtentX" VALUE="10583">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="COCodeRec">
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
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"Tix","RequestStatus","DueDate"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="68,68,68">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Ticket","Request Status","Date Response Due"'>
	<PARAM NAME="DetailAlignment" VALUE=",,">
	<PARAM NAME="HeaderAlignment" VALUE=",,">
	<PARAM NAME="DetailBackColor" VALUE=",,">
	<PARAM NAME="HeaderBackColor" VALUE=",,">
	<PARAM NAME="HeaderFont" VALUE=",,">
	<PARAM NAME="HeaderFontColor" VALUE=",,">
	<PARAM NAME="HeaderFontSize" VALUE=",,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,">
	<PARAM NAME="DetailFont" VALUE=",,">
	<PARAM NAME="DetailFontColor" VALUE=",,">
	<PARAM NAME="DetailFontSize" VALUE=",,">
	<PARAM NAME="DetailFontStyle" VALUE=",,">
	<PARAM NAME="ColumnCount" VALUE="3">
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
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="10">
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
	<PARAM NAME="GridWidth" VALUE="400">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 10;
Grid1.setDataSource(COCodeRec);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=0 cols=3 rules=ALL WIDTH=400';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=68';
Grid1.headerWidth[1] = ' WIDTH=68';
Grid1.headerWidth[2] = ' WIDTH=68';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'Ticket\'';
Grid1.colHeader[1] = '\'Request Status\'';
Grid1.colHeader[2] = '\'Date Response Due\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=68';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'COCodeRec.fields.getValue(\'Tix\')';
Grid1.colAttributes[1] = '  WIDTH=68';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'COCodeRec.fields.getValue(\'RequestStatus\')';
Grid1.colAttributes[2] = '  WIDTH=68';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'COCodeRec.fields.getValue(\'DueDate\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid1.hasPageNumber = true;
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
	</TR>
</TABLE>
</BODY>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=COCodeRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sfrom\sxca_part1\s\q,TCControlID_Unmatched=\qCOCodeRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sfrom\sxca_part1\s\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCOCodeRec()
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
	cmdTmp.CommandText = 'SELECT * from xca_part1 ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	COCodeRec.setRecordSource(rsTmp);
	if (thisPage.getState('pb_COCodeRec') != null)
		COCodeRec.setBookmark(thisPage.getState('pb_COCodeRec'));
}
function _COCodeRec_ctor()
{
	CreateRecordset('COCodeRec', _initCOCodeRec, null);
}
function _COCodeRec_dtor()
{
	COCodeRec._preserveState();
	thisPage.setState('pb_COCodeRec', COCodeRec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->


<%
'COCodeRec.close
%>

<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

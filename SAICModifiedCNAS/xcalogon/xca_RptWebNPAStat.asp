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


<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnSubmit_onclick()
if COCodeRec.isopen() then COCodeRec.close()


'Want to declare status as "S"
Dim Status
  Status="S"
	TXT=ListNPA.selectedIndex
	TXT=ListNPA.gettext(TXT)
If TXT<>"" Then
	SQL0=" And NPA = '" & TXT & "' Order By NXX"
Else
	SQL0=""
End If

'SQL statement is checking for all NPA/NXXs with a status of "S" and outputting the value 'Available'.  Can
'be checked using the recordsets' SQL builder
RecSQL="select NPA, NXX, 'Available' from xca_cocode Where Status =  '" & Status & " '"  & SQL0
 

	RecSQL = RecSQL & SQL	
	COCodeRec.setsqltext(RecSQL)
	COCodeRec.open
	
End Sub

Sub btnClose_onclick()
	Response.Redirect "xca_MenuRptPost.asp"
End Sub

'With Grid1, wherever a NPA and NXX has a status of "S"(available), the unbound column prints out the word
'"Available".  The unbound column is not actually reading the status of a CO Code.


</SCRIPT>
</HEAD>
<BODY bgColor=#d7c7a4 bgProperties="fixed" text="black">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=COCodeRec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sNPA,\sNXX,\s'Available'\sFROM\sxca_cocode\sWHERE\sStatus\s=\s'S'\q,TCControlID_Unmatched=\qCOCodeRec\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sNPA,\sNXX,\s'Available'\sFROM\sxca_cocode\sWHERE\sStatus\s=\s'S'\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initCOCodeRec()
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
	cmdTmp.CommandText = 'SELECT NPA, NXX, \'Available\' FROM xca_cocode WHERE Status = \'S\'';
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=NPARec 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasapp\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\q,TCControlID_Unmatched=\qNPARec\q,TCPPConn=\qcnasapp\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNPARec()
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
	cmdTmp.CommandText = 'Select Distinct NPA from xca_COCode';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	NPARec.setRecordSource(rsTmp);
	NPARec.open();
	if (thisPage.getState('pb_NPARec') != null)
		NPARec.setBookmark(thisPage.getState('pb_NPARec'));
}
function _NPARec_ctor()
{
	CreateRecordset('NPARec', _initNPARec, null);
}
function _NPARec_dtor()
{
	NPARec._preserveState();
	thisPage.setState('pb_NPARec', NPARec.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<div align="center"><center><table border="0" cellpadding="2"><tr>
	<td nowrap><font color=maroon face="Arial Black" size="5"><strong>
NPA - CO Codes Availability List</strong></font></td></tr></table></center></div>

<p></p>

<TABLE Align=right nowrap border=0 cellPadding=0 cellSpacing=0 height=23 style="HEIGHT: 23px; WIDTH: 226px" width=226>
	<TR>
		<TD><FONT face="Arial" size=4 color=black><STRONG>Created
            <%
		Response.write "" & Date() 
		
		%></STRONG></FONT></TD></TR></TABLE>
		
		<p>&nbsp;</p>

<TABLE width=60% ALIGN=center border=0 cellspacing=5 cellpadding=1>
<TR>
		<TD align=right nowrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 id=Label1 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
	width=38>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="3">
	<PARAM NAME="FontColor" VALUE="Black">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('NPA');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD align=left nowrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 
            id=ListNPA style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" 
            width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListNPA">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="NPARec">
	<PARAM NAME="BoundColumn" VALUE="NPA">
	<PARAM NAME="ListField" VALUE="NPA">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListNPA()
{
	NPARec.advise(RS_ONDATASETCOMPLETE, 'ListNPA.setRowSource(NPARec, \'NPA\', \'NPA\');');
}
function _ListNPA_ctor()
{
	CreateListbox('ListNPA', _initListNPA, null);
}
</script>
<% ListNPA.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<FONT color=black Face="Arial" size=2>(<em>Select a NPA to view the available CO Codes</em>)</FONT> 
    
		</TD>
	</TR>
</TABLE>

<TABLE WIDTH=5% ALIGN=center BORDER=0 CELLSPACING=5 CELLPADDING=1>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnSubmit 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 236px; WIDTH: 63px" width=63>
	<PARAM NAME="_ExtentX" VALUE="1667">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSubmit">
	<PARAM NAME="Caption" VALUE="Submit">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnClose 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 263px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnClose">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
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
</TD>
	</TR>
</TABLE> 

<p></p>

<TABLE align=center nowrap border=0 cellPadding=0 cellSpacing=0 height=149 style="HEIGHT: 149px; WIDTH: 703px" width=703>
<TR>
	<TD noWrap align=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" 
            height=147 id=Grid1 
            style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 256px" width=256>
	<PARAM NAME="_ExtentX" VALUE="6773">
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
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE="&quot;NPA&quot;,&quot;NXX&quot;,&quot;='Available'&quot;">
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="85,85,85">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,1,0">
	<PARAM NAME="DisplayName" VALUE='"NPA","CO Code","Status"'>
	<PARAM NAME="DetailAlignment" VALUE=",2,">
	<PARAM NAME="HeaderAlignment" VALUE=",0,">
	<PARAM NAME="DetailBackColor" VALUE=",Silver,">
	<PARAM NAME="HeaderBackColor" VALUE=",#669999,">
	<PARAM NAME="HeaderFont" VALUE=",Arial,">
	<PARAM NAME="HeaderFontColor" VALUE=",Black,">
	<PARAM NAME="HeaderFontSize" VALUE=",4,">
	<PARAM NAME="HeaderFontStyle" VALUE=",1,">
	<PARAM NAME="DetailFont" VALUE=",Arial,">
	<PARAM NAME="DetailFontColor" VALUE=",Black,">
	<PARAM NAME="DetailFontSize" VALUE=",2,">
	<PARAM NAME="DetailFontStyle" VALUE=",0,">
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
	<PARAM NAME="RowAlignment" VALUE="2">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="20">
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
	<PARAM NAME="GridWidth" VALUE="256">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="437229">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 20;
Grid1.setDataSource(COCodeRec);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=256 nowrap';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center nowrap';
Grid1.headerWidth[0] = ' WIDTH=85';
Grid1.headerWidth[1] = ' WIDTH=85';
Grid1.headerWidth[2] = ' WIDTH=85';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPA\'';
Grid1.colHeader[1] = '\'CO Code\'';
Grid1.colHeader[2] = '\'Status\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Center nowrap  bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Center nowrap  bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=85';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'COCodeRec.fields.getValue(\'NPA\')';
Grid1.colAttributes[1] = '  WIDTH=85 bgcolor=Silver align=Center nowrap';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'COCodeRec.fields.getValue(\'NXX\')';
Grid1.colAttributes[2] = '  WIDTH=85';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = '\'Available\'';
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
</TD>
	</TR>
</TABLE>

<P>&nbsp;</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

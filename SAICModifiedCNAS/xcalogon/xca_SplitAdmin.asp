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
<%
'<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Connection</TITLE>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>


dim RecIndex

Sub btnUpdate_onclick()
	session("pSplitID")=Rec1.fields.getValue("NPASplitID")	
	session("pCurrentNPA")=Rec1.fields.getValue("CurrentNPA")	
	session("pNewNPA")=Rec1.fields.getValue("NewNPA")
	session("pPDPStartDate")=Rec1.fields.getValue("PDPStartDate")
	session("pPDPEndDate")=Rec1.fields.getValue("PDPEndDate")
	if session("pSplitID")<>"" then
		Response.Redirect "xca_Split.asp"
	end if	
End Sub

Sub btnAdd_onclick()
	session("pSplitID")=""	
	session("pCurrentNPA")=""
	session("pNewNPA")=""
	session("pPDPStartDate")=""
	session("pPDPEndDate")=""
	Response.Redirect "xca_Split.asp"
End Sub

Sub btnDelete_onclick()
	'rec1.deleteRecord
End Sub

function c()
	RecIndex=RecIndex+1
	c=RecIndex
End Function

Sub btngoto_onclick()
	Rec1.movefirst
	Rec1.move cInt(txtNum.value)-1
End Sub

Sub btnReturnToMain_onclick()
	Response.Redirect "xca_MenuC0CAdmin.asp"

End Sub

</SCRIPT>

</HEAD>

<BODY bgColor="#d7c7a4"><BR>


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject=\qViews\q,TCDBObjectName=\qv_xca_NPASplit_P\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qViews\q,TCPPDBObjectName=\qv_xca_NPASplit_P\q,TCSQLStatement_Unmatched=\qselect\s*\sfrom\sv_xca_NPASplit_P\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRec1()
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
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '"v_xca_NPASplit_P"';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Rec1.setRecordSource(rsTmp);
	Rec1.open();
	if (thisPage.getState('pb_Rec1') != null)
		Rec1.setBookmark(thisPage.getState('pb_Rec1'));
}
function _Rec1_ctor()
{
	CreateRecordset('Rec1', _initRec1, null);
}
function _Rec1_dtor()
{
	Rec1._preserveState();
	thisPage.setState('pb_Rec1', Rec1.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->



<TABLE WIDTH=75% ALIGN=center border=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD ALIGN=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=37 id=lblTitle 
	style="HEIGHT: 37px; LEFT: 10px; TOP: 113px; WIDTH: 355px" width=355>
	<PARAM NAME="_ExtentX" VALUE="9393">
	<PARAM NAME="_ExtentY" VALUE="979">
	<PARAM NAME="id" VALUE="lblTitle">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA Split Administration">
	<PARAM NAME="FontFace" VALUE="Arial Black">
	<PARAM NAME="FontSize" VALUE="5">
	<PARAM NAME="FontColor" VALUE="Maroon">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial Black" SIZE="5" COLOR="Maroon"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlblTitle()
{
	lblTitle.setCaption('NPA Split Administration');
}
function _lblTitle_ctor()
{
	CreateLabel('lblTitle', _initlblTitle, null);
}
</script>
<% lblTitle.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><BR>

<BR>
<TABLE ALIGN=center border=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnAdd style="HEIGHT: 27px; LEFT: 10px; TOP: 150px; WIDTH: 62px" 
	width=62>
	<PARAM NAME="_ExtentX" VALUE="1640">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnAdd">
	<PARAM NAME="Caption" VALUE="  New  ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnAdd()
{
	btnAdd.value = '  New  ';
	btnAdd.setStyle(0);
}
function _btnAdd_ctor()
{
	CreateButton('btnAdd', _initbtnAdd, null);
}
</script>
<% btnAdd.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnUpdate 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 177px; WIDTH: 51px" width=51>
	<PARAM NAME="_ExtentX" VALUE="1349">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE=" Edit ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnUpdate()
{
	btnUpdate.value = ' Edit ';
	btnUpdate.setStyle(0);
}
function _btnUpdate_ctor()
{
	CreateButton('btnUpdate', _initbtnUpdate, null);
}
</script>
<% btnUpdate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</TD><TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnReturnToMain 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 204px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturnToMain">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturnToMain()
{
	btnReturnToMain.value = 'Return';
	btnReturnToMain.setStyle(0);
}
function _btnReturnToMain_ctor()
{
	CreateButton('btnReturnToMain', _initbtnReturnToMain, null);
}
</script>
<% btnReturnToMain.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</TD>
	
	</TR>
</TABLE>

<BR><BR><TABLE ALIGN=center border=1 cellspacing=1 cellpadding=1 bgcolor=white>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 797px" 
	width=797>
	<PARAM NAME="_ExtentX" VALUE="21087">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="Rec1">
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
	<PARAM NAME="EnableRowNav" VALUE="-1">
	<PARAM NAME="HiliteColor" VALUE="LimeGreen">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPASplitID","CurrentNPA","NewNPA","PDPStartDate","PDPEndDate"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4">
	<PARAM NAME="displayWidth" VALUE="87,85,81,113,128">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPA Split ID","Current NPA","New NPA","PDP Start Date","PDP End Date"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,">
	<PARAM NAME="ColumnCount" VALUE="5">
	<PARAM NAME="CurStyle" VALUE="Basic Maroon">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="8388608">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="0">
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
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="16777215">
	<PARAM NAME="GridBackColor" VALUE="8388608">
	<PARAM NAME="AltRowBckgnd" VALUE="16777215">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="797">
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
Grid1.setDataSource(Rec1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=5 rules=ALL WIDTH=797';
Grid1.headerAttributes = '   bgcolor=Maroon align=Left';
Grid1.headerWidth[0] = ' WIDTH=87';
Grid1.headerWidth[1] = ' WIDTH=85';
Grid1.headerWidth[2] = ' WIDTH=81';
Grid1.headerWidth[3] = ' WIDTH=113';
Grid1.headerWidth[4] = ' WIDTH=128';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPA Split ID\'';
Grid1.colHeader[1] = '\'Current NPA\'';
Grid1.colHeader[2] = '\'New NPA\'';
Grid1.colHeader[3] = '\'PDP Start Date\'';
Grid1.colHeader[4] = '\'PDP End Date\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=87';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'NPASplitID\')';
Grid1.colAttributes[1] = '  WIDTH=85';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'CurrentNPA\')';
Grid1.colAttributes[2] = '  WIDTH=81';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'NewNPA\')';
Grid1.colAttributes[3] = '  WIDTH=113';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'Rec1.fields.getValue(\'PDPStartDate\')';
Grid1.colAttributes[4] = '  WIDTH=128';
Grid1.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[4] = 'Rec1.fields.getValue(\'PDPEndDate\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid1.hasPageNumber = true;
Grid1.hiliteAttributes = ' bgcolor=LimeGreen';
var objRecNavbar = Grid1.showRecordNavbar(40,1);
objRecNavbar.getButton(1).value = '   <   ';
objRecNavbar.getButton(2).value = '   >   ';
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

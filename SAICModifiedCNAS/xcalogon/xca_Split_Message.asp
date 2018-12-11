<%@ Language=VBScript %>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btnOK_onclick()
	
	Response.Redirect "xca_Split.asp"
	
End Sub


</SCRIPT>
</HEAD>
<BODY bgColor="#d7c7a4">

<%

dim NPA
dim SplitID
NPA=session("NPA_S_M")
SplitID=session("SplitID_S_M")
Rec1.open

%>

<BR>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Rec1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sNPA,\sNXX,\sTransferID,\sNPASplitID\sFROM\sxca_COCode\sWHERE\s(NPA\s=\s?)\sAND\s(TransferID\s\l\g\s'Excluded')\sAND\s(NPASplitID\s=\s?)\q,TCControlID_Unmatched=\qRec1\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_COCode\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sNPA,\sNXX,\sTransferID,\sNPASplitID\sFROM\sxca_COCode\sWHERE\s(NPA\s=\s?)\sAND\s(TransferID\s\l\g\s'Excluded')\sAND\s(NPASplitID\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=2,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1,CValue_Unmatched=\qNPA\q),Row2=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam2\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q10\q,CReq=0,CValue_Unmatched=\qSplitID\q)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersRec1()
{
	Rec1.setParameter(0,NPA);
	Rec1.setParameter(1,SplitID);
}
function _initRec1()
{
	Rec1.advise(RS_ONBEFOREOPEN, _setParametersRec1);
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
	cmdTmp.CommandText = 'SELECT NPA, NXX, TransferID, NPASplitID FROM xca_COCode WHERE (NPA = ?) AND (TransferID <> \'Excluded\') AND (NPASplitID = ?)';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Rec1.setRecordSource(rsTmp);
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
<TABLE WIDTH=75% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=27 id=Label1 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 891px" 
	width=891>
	<PARAM NAME="_ExtentX" VALUE="23574">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="NPA Split can not be completed because the CO Codes listed below are associated with an active transfer.">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT SIZE="4">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setCaption('NPA Split can not be completed because the CO Codes listed below are associated with an active transfer.');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE><BR>

<TABLE align=center border=0 cellPadding=1 cellSpacing=1 width="75%" id=TABLE1>
  
  <TR>
    <TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=27 id=Label2 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 1078px" 
	width=1078>
	<PARAM NAME="_ExtentX" VALUE="28522">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Either complete the transfer or exclude the code from the transfer and include in a new transfer following NPA Split completion.">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="4">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<FONT SIZE="4">
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setCaption('Either complete the transfer or exclude the code from the transfer and include in a new transfer following NPA Split completion.');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>
</FONT>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD></TR></TABLE><BR>
<P></P>
<TABLE  ALIGN=center BORDER=2 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 255px" 
	width=255>
	<PARAM NAME="_ExtentX" VALUE="6747">
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
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPA","NXX","TransferID"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="68,68,114">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPA","NXX","Transfer ID"'>
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
	<PARAM NAME="TitleAlignment" VALUE="0">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="12632256">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="800">
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
	<PARAM NAME="GridWidth" VALUE="255">
	<PARAM NAME="EnablePaging" VALUE="0">
	<PARAM NAME="ShowStatus" VALUE="0">
	<PARAM NAME="StyleValue" VALUE="452653">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 0;
Grid1.setDataSource(Rec1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=1 cols=3 rules=ALL WIDTH=255';
Grid1.headerAttributes = '   bgcolor=Maroon align=Left';
Grid1.headerWidth[0] = ' WIDTH=68';
Grid1.headerWidth[1] = ' WIDTH=68';
Grid1.headerWidth[2] = ' WIDTH=114';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPA\'';
Grid1.colHeader[1] = '\'NXX\'';
Grid1.colHeader[2] = '\'Transfer ID\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Left bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Left bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=68';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Rec1.fields.getValue(\'NPA\')';
Grid1.colAttributes[1] = '  WIDTH=68';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Rec1.fields.getValue(\'NXX\')';
Grid1.colAttributes[2] = '  WIDTH=114';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Rec1.fields.getValue(\'TransferID\')';
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
</TABLE><BR><BR>

<TABLE  ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnOK style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 52px" 
	width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnOK">
	<PARAM NAME="Caption" VALUE="  OK  ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnOK()
{
	btnOK.value = '  OK  ';
	btnOK.setStyle(0);
}
function _btnOK_ctor()
{
	CreateButton('btnOK', _initbtnOK, null);
}
</script>
<% btnOK.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD>
	</TR>
</TABLE>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>

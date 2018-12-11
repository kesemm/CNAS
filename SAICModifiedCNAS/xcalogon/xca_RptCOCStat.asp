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

	
	checkbox=chkFilterAll.getChecked(checkbox)
	TXT=ListNPA.selectedIndex
	TXT=ListNPA.gettext(TXT)
	If checkbox="True" Then
	SQL0=""
	Else
	If TXT<>"" Then
			SQL0=" Where NPA = '" & TXT & "'"
		End If
	End If
	
	TXT=ListSort1.selectedIndex
	TXT=ListSort1.getvalue(TXT)
	If TXT<>"" Then
		SQL1=TXT
	Else
		SQL1=""
	End If
			
		
	TXT=ListSort2.selectedIndex
	TXT=ListSort2.getvalue(TXT)
	If TXT<>"" Then
		SQL2=TXT
	Else
		SQL2=""
	End If
		
	RecSQL="SELECT xca_COCode.NPA, xca_COCode.NXX, xca_COCode.EntityID, xca_status_codes.COStatusDescription FROM xca_COCode INNER JOIN xca_status_codes ON xca_COCode.Status = xca_status_codes.COStatus  " & SQL0
	'RecSQL="SELECT * FROM xca_COCode, xca_status_codes " & SQL0
	If SQL1<>"" and SQL2<>"" Then
		If SQL1=SQL2 Then
			SQL=" ORDER BY " & SQL1
		Else
			SQL=" ORDER BY " & SQL1 & "," & SQL2
		End If
	
	Elseif SQL1<>"" Then
		SQL=" ORDER BY " & SQL1
	Elseif SQL2<>"" Then
		SQL=" ORDER BY " & SQL2
	Else
		SQL=""
		
	End If

	RecSQL = RecSQL & SQL	
	COCodeRec.setsqltext(RecSQL)
	COCodeRec.open
End Sub




Sub btnClose_onclick()
	Response.Redirect "xca_RptAdminMenu.asp"
End Sub

</SCRIPT>
</HEAD>
<BODY bgColor=#d7c7a4 bgProperties="fixed" text="black">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=COCodeRec style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\sxca_COCode.NPA,\sxca_COCode.NXX,\sxca_COCode.EntityID,\sxca_status_codes.COStatusDescription\sFROM\sxca_COCode\sINNER\sJOIN\sxca_status_codes\sON\sxca_COCode.Status\s=\sxca_status_codes.COStatus\sWHERE\s(xca_COCode.NPA\s=\s?)\q,TCControlID_Unmatched=\qCOCodeRec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\sxca_COCode.NPA,\sxca_COCode.NXX,\sxca_COCode.EntityID,\sxca_status_codes.COStatusDescription\sFROM\sxca_COCode\sINNER\sJOIN\sxca_status_codes\sON\sxca_COCode.Status\s=\sxca_status_codes.COStatus\sWHERE\s(xca_COCode.NPA\s=\s?)\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=1,Row1=(CType_Unmatched=\q?\q,CParName_Unmatched=\qParam1\q,CDataType_Unmatched=\qVarChar\q,CSize_Unmatched=\q3\q,CReq=1)))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _setParametersCOCodeRec()
{
}
function _initCOCodeRec()
{
	COCodeRec.advise(RS_ONBEFOREOPEN, _setParametersCOCodeRec);
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
	cmdTmp.CommandText = 'SELECT xca_COCode.NPA, xca_COCode.NXX, xca_COCode.EntityID, xca_status_codes.COStatusDescription FROM xca_COCode INNER JOIN xca_status_codes ON xca_COCode.Status = xca_status_codes.COStatus WHERE (xca_COCode.NPA = ?)';
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
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\sOrder\sBy\sNPA\q,TCControlID_Unmatched=\qNPARec\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qT1\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\sDistinct\sNPA\sfrom\sxca_COCode\sOrder\sBy\sNPA\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initNPARec()
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
	cmdTmp.CommandText = 'Select Distinct NPA from xca_COCode Order By NPA';
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
CO Code Status 
Report</strong></font></td></tr></table></center></div>

<p>&nbsp;</p>

<TABLE Align=right nowrap border=0 cellPadding=0 cellSpacing=0 height=23 style="HEIGHT: 23px; WIDTH: 226px" width=226>
	<TR>
		<TD><FONT face="Arial" size=4 color=black><STRONG>Created
            <% Response.write "" & Date() %></STRONG></FONT>
        </TD>
	</TR>
</TABLE>
		
		<p>&nbsp;</p>



<TABLE width=60% ALIGN=center border=0 cellspacing=5 cellpadding=1 background="" style="WIDTH: 60%">

<TR><TD align=middle noWrap colSpan=2>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label4 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 73px" 
            width=73>
	<PARAM NAME="_ExtentX" VALUE="1931">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label4">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Filter By:">
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
function _initLabel4()
{
	Label4.setCaption('Filter By:');
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
        <TD align=right noWrap vAlign=bottom><TD align=left noWrap vAlign=bottom>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label2 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 116px" 
            width=116>
	<PARAM NAME="_ExtentX" VALUE="3069">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label2">
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
<FONT FACE="Arial" SIZE="3" COLOR="Black"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setCaption('First Order By:');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD align=left nowrap vAlign=bottom>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label3 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 141px" 
            width=141>
	<PARAM NAME="_ExtentX" VALUE="3731">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label3">
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
function _initLabel3()
{
	Label3.setCaption('Second Order By:');
}
function _Label3_ctor()
{
	CreateLabel('Label3', _initLabel3, null);
}
</script>
<% Label3.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</TD><TD align=left nowrap vAlign=bottom> 
</TD></TR>
    <TR>
        <TD align=right noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E46C-DC5F-11D0-9846-0000F8027CA0" height=27 
            id=chkFilterAll 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 215px; WIDTH: 29px" width=29>
	<PARAM NAME="_ExtentX" VALUE="767">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="chkFilterAll">
	<PARAM NAME="Caption" VALUE="">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/CheckBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _chkFilterAll_ctor()
{
	CreateCheckbox('chkFilterAll', null, null);
}
</script>
<% chkFilterAll.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD align=left noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label6 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 70px" 
            width=70>
	<PARAM NAME="_ExtentX" VALUE="1852">
	<PARAM NAME="_ExtentY" VALUE="609">
	<PARAM NAME="id" VALUE="Label6">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="All NPAs">
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
function _initLabel6()
{
	Label6.setCaption('All NPAs');
}
function _Label6_ctor()
{
	CreateLabel('Label6', _initLabel6, null);
}
</script>
<% Label6.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD align=left noWrap>
        <TD align=left noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=ListSort1 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 85px" width=85>
	<PARAM NAME="_ExtentX" VALUE="2249">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="5">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="NPA">
	<PARAM NAME="CLEV2" VALUE="NPA">
	<PARAM NAME="CLED3" VALUE="CO Code">
	<PARAM NAME="CLEV3" VALUE="NXX">
	<PARAM NAME="CLED4" VALUE="Status">
	<PARAM NAME="CLEV4" VALUE="Status">
	<PARAM NAME="CLED5" VALUE="Entity ID">
	<PARAM NAME="CLEV5" VALUE="EntityID">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort1()
{
	ListSort1.addItem('', '');
	ListSort1.addItem('NPA', 'NPA');
	ListSort1.addItem('CO Code', 'NXX');
	ListSort1.addItem('Status', 'Status');
	ListSort1.addItem('Entity ID', 'EntityID');
}
function _ListSort1_ctor()
{
	CreateListbox('ListSort1', _initListSort1, null);
}
</script>
<% ListSort1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        <TD align=left noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=ListSort2 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 85px" width=85>
	<PARAM NAME="_ExtentX" VALUE="2249">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="ListSort2">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="-1">
	<PARAM NAME="CLSize" VALUE="5">
	<PARAM NAME="CLED1" VALUE="">
	<PARAM NAME="CLEV1" VALUE="">
	<PARAM NAME="CLED2" VALUE="NPA">
	<PARAM NAME="CLEV2" VALUE="NPA">
	<PARAM NAME="CLED3" VALUE="CO Code">
	<PARAM NAME="CLEV3" VALUE="NXX">
	<PARAM NAME="CLED4" VALUE="Status">
	<PARAM NAME="CLEV4" VALUE="Status">
	<PARAM NAME="CLED5" VALUE="Entity ID">
	<PARAM NAME="CLEV5" VALUE="EntityID">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListSort2()
{
	ListSort2.addItem('', '');
	ListSort2.addItem('NPA', 'NPA');
	ListSort2.addItem('CO Code', 'NXX');
	ListSort2.addItem('Status', 'Status');
	ListSort2.addItem('Entity ID', 'EntityID');
}
function _ListSort2_ctor()
{
	CreateListbox('ListSort2', _initListSort2, null);
}
</script>
<% ListSort2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
	<TR>
		<TD align=right nowrap>
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
</TD>
        <TD align=left noWrap>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=23 
            id=Label1 style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 38px" 
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
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
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
<TD align=left nowrap>
</TD>
		<TD align=left nowrap>
</TD>
		<TD align=left nowrap>
</TD>
	</TR>
</TABLE>

<TABLE WIDTH=5% ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnSubmit 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 394px; WIDTH: 63px" width=63>
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
		<TD>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
            id=btnClose 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 421px; WIDTH: 61px" width=61>
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
</TD>
	</TR>
</TABLE> 

<p>&nbsp;</p>

<TABLE align=center nowrap border=0 cellPadding=0 cellSpacing=0 height=149 style="HEIGHT: 149px; WIDTH: 703px" width=703>
<TR>
	<TD noWrap align=middle>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" 
            height=147 id=Grid1 
            style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 421px" width=421>
	<PARAM NAME="_ExtentX" VALUE="11139">
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
	<PARAM NAME="RecNavBarHasNextButton" VALUE="0">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="0">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE='"NPA","NXX","COStatusDescription","EntityID"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3">
	<PARAM NAME="displayWidth" VALUE="69,80,85,99">
	<PARAM NAME="Coltype" VALUE="1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"NPA","CO Code","Status","Entity ID"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,">
	<PARAM NAME="HeaderFont" VALUE=",,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,">
	<PARAM NAME="DetailFont" VALUE=",,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,">
	<PARAM NAME="ColumnCount" VALUE="4">
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
	<PARAM NAME="PageSize" VALUE="800">
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
	<PARAM NAME="GridWidth" VALUE="421">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453421">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
</OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 800;
Grid1.setDataSource(COCodeRec);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolor=White bgcolor=Maroon border=0 cols=4 rules=ALL WIDTH=421';
Grid1.headerAttributes = '   bgcolor=Maroon align=Center';
Grid1.headerWidth[0] = ' WIDTH=69';
Grid1.headerWidth[1] = ' WIDTH=80';
Grid1.headerWidth[2] = ' WIDTH=85';
Grid1.headerWidth[3] = ' WIDTH=99';
Grid1.headerFormat = '<Font face="Arial" size=4 color=White> <b>';
Grid1.colHeader[0] = '\'NPA\'';
Grid1.colHeader[1] = '\'CO Code\'';
Grid1.colHeader[2] = '\'Status\'';
Grid1.colHeader[3] = '\'Entity ID\'';
Grid1.rowAttributes[0] = '  bgcolor = Silver align=Center bordercolor=White';
Grid1.rowAttributes[1] = '  bgcolor = White align=Center bordercolor=White';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=69';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'COCodeRec.fields.getValue(\'NPA\')';
Grid1.colAttributes[1] = '  WIDTH=80';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'COCodeRec.fields.getValue(\'NXX\')';
Grid1.colAttributes[2] = '  WIDTH=85';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'COCodeRec.fields.getValue(\'COStatusDescription\')';
Grid1.colAttributes[3] = '  WIDTH=99';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'COCodeRec.fields.getValue(\'EntityID\')';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(40,1);
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
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

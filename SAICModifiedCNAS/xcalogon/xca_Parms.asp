<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<form name="thisForm" METHOD="post">

<!--#include file="xca_CNASLib.inc"-->
<!--<form name="Entity" METHOD="post" onrowexit="return validateForm()">-->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>CNAS Entity Security Administration</title>
<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">






	

Sub btnEntityUpdate_onclick()
	
		getParmsdata.updateRecord
	

End Sub






Sub btnReturntoMain_onclick()
	Response.Redirect "xca_MenuSecurityAdmin.asp"
End Sub

</script>
   
  </head>   
        
     

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=getParmsdata style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qcnasadmin\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qselect\s*\sfrom\sxca_Parms\sorder\sby\sName\q,TCControlID_Unmatched=\qgetParmsdata\q,TCPPConn=\qcnasadmin\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qxca_Parms\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qselect\s*\sfrom\sxca_Parms\sorder\sby\sName\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initgetParmsdata()
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
	cmdTmp.CommandText = 'select * from xca_Parms order by Name';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	getParmsdata.setRecordSource(rsTmp);
	getParmsdata.open();
	if (thisPage.getState('pb_getParmsdata') != null)
		getParmsdata.setBookmark(thisPage.getState('pb_getParmsdata'));
}
function _getParmsdata_ctor()
{
	CreateRecordset('getParmsdata', _initgetParmsdata, null);
}
function _getParmsdata_dtor()
{
	getParmsdata._preserveState();
	thisPage.setState('pb_getParmsdata', getParmsdata.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<body bgColor="#d7c7a4" bgProperties="fixed" text="black">
<P><center><font face="Arial Black" color=maroon size=5><strong>CNAS Parameter 
Form</strong></font></center>
<P></P>

<P>&nbsp;<P>

<table align="center" border="0" cellPadding="1" cellSpacing="0" height="225" > 
    <tr>
        <td noWrap bgColor="#993300" align="left">
      <font face="Arial" size="4" style="COLOR: snow; FONT-STYLE: normal; FONT-WEIGHT: bold">
        <strong> Parms</strong></font> 

</td>
        <td noWrap bgColor="#993300">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnEdit style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 51px" 
width=51>
	<PARAM NAME="_ExtentX" VALUE="1349">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnEdit">
	<PARAM NAME="Caption" VALUE="EDIT">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      										      										                     
    </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnEdit()
{
	btnEdit.value = 'EDIT';
	btnEdit.setStyle(0);
}
function _btnEdit_ctor()
{
	CreateButton('btnEdit', _initbtnEdit, null);
}
</script>
<% btnEdit.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

</td>

        <td noWrap bgColor="#993300">&nbsp; 
            

</td>
        <td noWrap bgColor="#993300">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnEntityUpdate style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 78px" 
      width=78>
	<PARAM NAME="_ExtentX" VALUE="2064">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnEntityUpdate">
	<PARAM NAME="Caption" VALUE="UPDATE">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      										      										                             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnEntityUpdate()
{
	btnEntityUpdate.value = 'UPDATE';
	btnEntityUpdate.setStyle(0);
}
function _btnEntityUpdate_ctor()
{
	CreateButton('btnEntityUpdate', _initbtnEntityUpdate, null);
}
</script>
<% btnEntityUpdate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;&nbsp; 

</td></tr>
    <tr>
        <td align="right" noWrap>
       </td><td>
            


</td>
        <td align="right" noWrap></td>
        <td align="left" noWrap vAlign="center">
</td></tr>
    <tr>
        <td align="right" noWrap></td>
        <td noWrap>
</td>
        <td noWrap>
            <div align="right">&nbsp;</div></td>
        <td noWrap>
</td></tr>
    <tr>
        <td align="right" noWrap><strong><font face="Arial" size="2"><font face="Arial" size="2"><strong> 
            Name&nbsp;&nbsp;</strong></font></font></strong></td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=24 
      id=Name style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 129px" width=129>
	<PARAM NAME="_ExtentX" VALUE="1085">
	<PARAM NAME="_ExtentY" VALUE="529">
	<PARAM NAME="id" VALUE="Name">
	<PARAM NAME="DataSource" VALUE="getParmsdata">
	<PARAM NAME="DataField" VALUE="Name">
	<PARAM NAME="FontFace" VALUE="Arial">
	<PARAM NAME="FontSize" VALUE="2">
	<PARAM NAME="FontColor" VALUE="Blue">
	<PARAM NAME="FontBold" VALUE="-1">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      															      														    </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<FONT FACE="Arial" SIZE="2" COLOR="Blue"><B>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initName()
{
	Name.setDataSource(getParmsdata);
	Name.setDataField('Name');
}
function _Name_ctor()
{
	CreateLabel('Name', _initName, null);
}
</script>
<% Name.display %>
</FONT></B>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
        <td noWrap align="left">
            <div align="right">
            <div align="right"><FONT face=Arial 
            size=2><STRONG>Value</STRONG></FONT></div></div>
</td>
        <td align="left" noWrap>
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
      id=Value style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 150px" 
width=150>
	<PARAM NAME="_ExtentX" VALUE="3969">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Value">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="getParmsdata">
	<PARAM NAME="DataField" VALUE="Value">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="25">
	<PARAM NAME="DisplayWidth" VALUE="25">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
      														      													    </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initValue()
{
	Value.setStyle(TXT_TEXTBOX);
	Value.setDataSource(getParmsdata);
	Value.setDataField('Value');
	Value.setMaxLength(25);
	Value.setColumnCount(25);
}
function _Value_ctor()
{
	CreateTextbox('Value', _initValue, null);
}
</script>
<% Value.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</td>
    <TR>
        <TD align=right noWrap>
        <TD noWrap>&nbsp;
        <TD align=left noWrap>
        <TD align=left noWrap>
    <tr>
        <td align="right" noWrap><FONT face=Arial 
            size=2></FONT></td>
        <td noWrap><FONT size=2><FONT 
            face=Arial><FONT>ADMIN = CNA Entity Name for 
            Part1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></FONT></FONT><FONT><FONT 
            face=Arial><FONT size=2> </FONT></FONT></FONT> 
</td>
        <td noWrap align="left">
            <div align="right">
            <div align="right"><FONT face=Arial 
            size=2></FONT>&nbsp;</div></div><FONT face=Arial size=2></FONT>
</td>
        <td align="left" noWrap><FONT face=Arial 
            size=2></FONT>
</td>
    <tr>
        <td align="right" noWrap><FONT face=Arial 
            size=2></FONT></td>
        <td align="left" noWrap><FONT size=2><FONT 
            face=Arial><FONT>IP = IP address of web server</FONT></FONT></FONT><FONT><FONT face=Arial><FONT size=2> 
            </FONT></FONT></FONT> 
</td>
        <td noWrap align="left">
            <div align="right"><FONT face=Arial 
            size=2></FONT>&nbsp;</div><FONT face=Arial size=2></FONT>
</td>
        <td noWrap><FONT face=Arial size=2></FONT>
		</td></tr>
	<tr>
        <td align="right" noWrap><FONT face=Arial 
            size=2></FONT></td>
        <td align="left" noWrap><FONT size=2><FONT 
            face=Arial><FONT>P1 Days = Part 1 Effective Date 
            Restriction</FONT></FONT></FONT><FONT><FONT face=Arial><FONT size=2> 
            </FONT></FONT></FONT> 
</td>
        <td noWrap align="left">
            <div align="right"><FONT face=Arial 
            size=2></FONT>&nbsp;</div><FONT face=Arial size=2></FONT>
</td>
        <td noWrap><FONT face=Arial size=2></FONT>
		</td></tr>
	<tr>
        <td align="right" noWrap>
            <div align="right"><FONT face=Arial 
            size=2></FONT>&nbsp;</div><FONT face=Arial size=2></FONT>
        <td align="left" noWrap><FONT size=2><FONT face=Arial><FONT>
			P4 Date = Override P4 In-Service Date Validation (Yes/No) 
            </FONT></FONT></FONT>
        <td noWrap align="right">
        <td noWrap align="left">
   <tr>
        <td align="right" noWrap>
        <td align="left" noWrap>&nbsp;
        <td noWrap align="right">
        <td noWrap align="left">
    </td>
    <tr>
        <td align="left" noWrap bgColor="#993300">&nbsp;</td>


        <td align="left" noWrap bgColor="#993300">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:58F3D268-FEDF-11D0-9C7F-0060081840F3" 
      id=EntityNavbar1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="4075">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="EntityNavbar1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="DataSource" VALUE="getParmsdata">
	<PARAM NAME="UpdateOnMove" VALUE="-1">
	<PARAM NAME="FirstCaption" VALUE=" |< ">
	<PARAM NAME="MoveFirst" VALUE="-1">
	<PARAM NAME="FirstImage" VALUE="0">
	<PARAM NAME="PrevCaption" VALUE="  <  ">
	<PARAM NAME="MovePrev" VALUE="-1">
	<PARAM NAME="PrevImage" VALUE="0">
	<PARAM NAME="NextCaption" VALUE="  >  ">
	<PARAM NAME="MoveNext" VALUE="-1">
	<PARAM NAME="NextImage" VALUE="0">
	<PARAM NAME="LastCaption" VALUE=" >| ">
	<PARAM NAME="MoveLast" VALUE="-1">
	<PARAM NAME="LastImage" VALUE="0">
	<PARAM NAME="Alignment" VALUE="1">
	<PARAM NAME="LocalPath" VALUE="../">
	
      																					      																					                
       </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initEntityNavbar1()
{
	EntityNavbar1.setAlignment(1);
	EntityNavbar1.setButtonStyles(170);
	EntityNavbar1.setDataSource(getParmsdata);
	EntityNavbar1.getButton(0).value = ' |< ';
	EntityNavbar1.getButton(1).value = '  <  ';
	EntityNavbar1.getButton(2).value = '  >  ';
	EntityNavbar1.getButton(3).value = ' >| ';
}
function _EntityNavbar1_ctor()
{
	CreateRecordsetNavbar('EntityNavbar1', _initEntityNavbar1, null);
}
</script>
<% EntityNavbar1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
            </td>
            
        <td align="right" noWrap bgColor="#993300">&nbsp;
        </td>
        <td align="right" noWrap bgColor="#993300">
      <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 
      id=btnReturntoMain 
      style="HEIGHT: 27px; LEFT: 10px; TOP: 233px; WIDTH: 61px" width=61>
	<PARAM NAME="_ExtentX" VALUE="1614">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnReturntoMain">
	<PARAM NAME="Caption" VALUE="Return">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
      										      									    
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnReturntoMain()
{
	btnReturntoMain.value = 'Return';
	btnReturntoMain.setStyle(0);
}
function _btnReturntoMain_ctor()
{
	CreateButton('btnReturntoMain', _initbtnReturntoMain, null);
}
</script>
<% btnReturntoMain.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
        </td></tr>
        </table>&nbsp;
        
<p></p>
<br>
<br>
<br>
<br>
<br>
<br>

<p>&nbsp;</p>
<p>&nbsp;</p>

<p>&nbsp;</p>

<p>&nbsp;</p>

<p>&nbsp;</p>

<p>&nbsp;</p>
<p>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:CEB04D01-0445-11D1-BB81-006097C553C8" height=23 
id=EntityFormManger style="HEIGHT: 23px; LEFT: 0px; TOP: 0px; WIDTH: 160px" 
width=160>
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="609">
	<PARAM NAME="State" VALUE="(txtName_Unmatched=\qEntityFormManger\q,txtNewMode_Unmatched=\q\q,grFormMode=(Rows=2,Row1=(txtMode_Unmatched=\qDisplay\q),Row2=(txtMode_Unmatched=\qEdit\q)),txtDefaultMode=\qDisplay\q,grMasterMode=(Rows=10,Row1=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qbtnEdit\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qfalse\q),Row2=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qbtnEntityUpdate\q,txtProperty_Unmatched=\qhide\q,txtValue_Unmatched=\q()\q),Row3=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qName\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qtrue\q),Row4=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qValue\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qtrue\q),Row5=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qEntityNavbar1\q,txtProperty_Unmatched=\qgetDataSource\q,txtValue_Unmatched=\q()\q),Row6=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qbtnEdit\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qtrue\q),Row7=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qbtnEntityUpdate\q,txtProperty_Unmatched=\qshow\q,txtValue_Unmatched=\q()\q),Row8=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qName\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qfalse\q),Row9=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qValue\q,txtProperty_Unmatched=\qdisabled\q,txtValue_Unmatched=\qfalse\q),Row10=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qEntityNavbar1\q,txtProperty_Unmatched=\qgetButton\q,txtValue_Unmatched=\q()\q)),grTransitions=(Rows=2,Row1=(txtCurrentMode=\qDisplay\q,txtObject=\qbtnEdit\q,txtEvent=\qonclick\q,txtNextMode=\qEdit\q),Row2=(txtCurrentMode=\qEdit\q,txtObject=\qbtnEntityUpdate\q,txtEvent=\qonclick\q,txtNextMode=\qDisplay\q)),grMasterStep=(Rows=2,Row1=(txtName_Unmatched=\q1\q,txtControl_Unmatched=\qgetParmsdata\q,txtAction_Unmatched=\qopen\q,txtValue_Unmatched=\q()\q),Row2=(txtName_Unmatched=\q2\q,txtControl_Unmatched=\qgetParmsdata\q,txtAction_Unmatched=\qupdateRecord\q,txtValue_Unmatched=\q()\q)))">
	
								     </OBJECT>
-->
<SCRIPT RUNAT=SERVER LANGUAGE="JavaScript">
function _EntityFormManger_ctor()
{
	thisPage.advise(PAGE_ONINIT, _EntityFormManger_init);
}
function _EntityFormManger_init()
{
	if (thisPage.getState("EntityFormManger_formmode") == null)
		_EntityFormManger_SetMode("Display");
	btnEdit.advise("onclick", "_EntityFormManger_btnEdit_onclick()");
	btnEntityUpdate.advise("onclick", "_EntityFormManger_btnEntityUpdate_onclick()");
}
function _EntityFormManger_SetMode(formmode)
{
	thisPage.setState("EntityFormManger_formmode", formmode);
	if (formmode == "Display")
	{
		btnEdit.disabled = false;
		btnEntityUpdate.hide();
		Name.disabled = true;
		Value.disabled = true;
		EntityNavbar1.getDataSource();
	}
	if (formmode == "Edit")
	{
		btnEdit.disabled = true;
		btnEntityUpdate.show();
		Name.disabled = false;
		Value.disabled = false;
		EntityNavbar1.getButton();
	}
}
function _EntityFormManger_btnEdit_onclick()
{
	if (thisPage.getState("EntityFormManger_formmode") == "Display")
	{
		getParmsdata.open();
		_EntityFormManger_SetMode("Edit");
	}
	else _EntityFormManger_SetMode(thisPage.getState("EntityFormManger_formmode"))
}
function _EntityFormManger_btnEntityUpdate_onclick()
{
	if (thisPage.getState("EntityFormManger_formmode") == "Edit")
	{
		getParmsdata.updateRecord();
		_EntityFormManger_SetMode("Display");
	}
	else _EntityFormManger_SetMode(thisPage.getState("EntityFormManger_formmode"))
}
</SCRIPT>


<!--METADATA TYPE="DesignerControl" endspan-->
</p>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</form>
</html>

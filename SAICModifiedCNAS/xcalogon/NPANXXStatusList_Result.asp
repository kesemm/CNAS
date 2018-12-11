<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA NXX Status List</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPANXXStatusList_Result.asp,v $
'* Commit Date:   $Date: 2015/08/11 11:56:10 $ (UTC)
'* Committed by:  $Author: browng $
'* CVS Revision:  $Revision: 1.3 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>NXX Status List for NPA: <%=request.querystring("NPA")%></BIG></B><BR></P>
<%
' SET UP THE CONNECTION AND RECORDSET

' SET THE FIRST PART OF THE QUERY TEXT

SQLQueryText = "SELECT [Status], [NPA], [NXX], [COStatusDescription] As [StatusDesc], ISNULL(RateCenter,'') AS RateCenter, ISNULL([EntityName], '') AS Company " & _
		"FROM xca_COCode " & _
		"Left JOIN xca_status_codes ON xca_COCode.status=xca_status_codes.COStatus " & _
		"LEFT JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID " & _
		"WHERE xca_COCode.NPA='" & request.querystring("NPA") & "' "

' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT

Select Case request.querystring("SortOrder")

Case "Company"
SQLQueryText = SQLQueryText & _
			"Order by Company, NXX"

Case "RateCenter"
SQLQueryText = SQLQueryText & _
	"Order by RateCenter, NXX"

Case "Status"
SQLQueryText = SQLQueryText & _
			"Order by StatusDesc, NXX"

Case "NXX"
SQLQueryText = SQLQueryText & _
			"Order by NXX"

Case Else
SQLQueryText = SQLQueryText & _
			"Order by NXX"

End Select

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPANXXStatusList = server.createObject("ADODB.recordset")
rstNPANXXStatusList.activeConnection = objConnection
rstNPANXXStatusList.CursorLocation = adUseServer
rstNPANXXStatusList.CursorType = adOpenStatic
rstNPANXXStatusList.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS TO DETERMINE IF NPA ENTERED IS VALID

if rstNPANXXStatusList.RecordCount>0 then

%>
<TABLE align="center" border="0" width="30%">
<TR>
<TD>
<font face="Arial" size="2">
<b>Instructions:</b><br>
- Click on the <image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/> to sort by that column<br>
- Available NXX Codes are <b><font style="color: #009900">Green</font></b><br>
- Unavailable NXX Codes are <font style="color: #663300">Red</font><br>
</font>
</TD>
</TR>
</TABLE>

<br>
<TABLE align="center" border="1">

<TD align="center"><B>NPA</B></TD>
<TD align="center"><B>NXX</B>
<a href="/xcalogon/NPANXXStatusList_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=NXX">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Status</B>
<a href="/xcalogon/NPANXXStatusList_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=Status">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Rate Center</B>
<a href="/xcalogon/NPANXXStatusList_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=RateCenter">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Company</B>
<a href="/xcalogon/NPANXXStatusList_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=Company">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
'<%
' LAUNCH THE SCOCAP PAGE FROM CNAC SITE
'if request.querystring("NPA")="613" or request.querystring("NPA")="819" then
'Response.Write ("<Script Language='JavaScript'>")
'Response.write("window.open ('http://www.cnac.ca/SCOCAP.htm')")
'Response.Write ("</Script>")
'Response.Flush
'end if
'%>

<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstNPANXXStatusList.EOF
%>
<%
' DETERMINE IF THIS IS AN AVAILABLE CODE
if rstNPANXXStatusList("Status")="S" then
%>
<TR>
<TD align="center"><b><font style="color: #009900">&nbsp; <%=rstNPANXXStatusList("NPA")%> &nbsp;</font></b></TD>
<TD align="center"><b><font style="color: #009900">&nbsp; <%=rstNPANXXStatusList("NXX")%> &nbsp;</font></b></TD>
<TD align="center"><b><font style="color: #009900">&nbsp; <%=rstNPANXXStatusList("StatusDesc")%> &nbsp;</font></b></TD>
<TD align="center"><b><font style="color: #009900">&nbsp; <%=rstNPANXXStatusList("RateCenter")%> &nbsp;</font></b></TD>
<TD align="center"><b><font style="color: #009900">&nbsp; <%=rstNPANXXStatusList("Company")%> &nbsp;</font></b></TD>
</TR>
<%
else
' USE THIS RED TO FORMAT THE ROW
%>
<TR>
<TD align="center"><font style="color: #663300">&nbsp; <%=rstNPANXXStatusList("NPA")%> &nbsp;</font></TD>
<TD align="center"><font style="color: #663300">&nbsp; <%=rstNPANXXStatusList("NXX")%> &nbsp;</font></TD>
<TD align="center"><font style="color: #663300">&nbsp; <%=rstNPANXXStatusList("StatusDesc")%> &nbsp;</font></TD>
<TD align="center"><font style="color: #663300">&nbsp; <%=rstNPANXXStatusList("RateCenter")%> &nbsp;</font></TD>
<TD align="center"><font style="color: #663300">&nbsp; <%=rstNPANXXStatusList("Company")%> &nbsp;</font></TD>
</TR>
<%
End If
%>
<%
' GET THE NEXT RECORD IN THE SET
rstNPANXXStatusList.movenext
%>
<%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>The NPA of <%=request.querystring("NPA")%> is Invalid, or a Database Error Occured</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstNPANXXStatusList.Close
objConnection.close
%>
<br>
<%
' THIS IS THE VERSION CONTROL INFORMATION BLOCK
' ---------------------------------------------
'
' Subdued input text box, that when clicked will make an alert with CVS Info
'
%>
<br><br>
<INPUT TYPE="TEXT" 
       STYLE="border: none; background-color: #D7C7A4; font: 7pt Arial; color: gray; width: 200px" 
       ONCLICK="VerInfo()" VALUE="CNAS Version Control Information"
       READONLY>
<SCRIPT language="JavaScript">
function VerInfo()
{
var strAlertText
strAlertText="SAIC Canada - CNAS Version Control Information     \n\n"
+"Version Control Managed by CVSNT & TortoiseCVS Interface     \n\n"
+"$RCSfile: NPANXXStatusList_Result.asp,v $\n"
+"$Revision: 1.3 $\n"
+"$Date: 2015/08/11 11:56:10 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>

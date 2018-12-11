<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA Company CO Codes</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPACompanyCOCodes_Result.asp,v $
'* Commit Date:   $Date: 2004/12/03 17:11:09 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.4 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>Company CO Codes in NPA: <%=request.querystring("NPA")%></BIG></B><BR></P>
<%
' SET UP THE CONNECTION AND RECORDSET

' SET THE FIRST PART OF THE QUERY TEXT

SQLQueryText = "SELECT xca_COCode.OCN,EntityName,NXX,COStatusDescription As Status, " & _
		"SwitchID,RateCenter FROM xca_COCode " & _
		"Left JOIN xca_Entity ON xca_COCode.EntityID=xca_Entity.EntityID " & _
		"Left Join xca_status_codes On xca_COCode.Status=xca_status_codes.COStatus " & _
		"Where (Status='a' or Status='i' or Status='r') " & _
		"and NPA=" & "'" & request.querystring("NPA") & "' "

' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT

Select Case request.querystring("SortOrder")
Case "Company"
SQLQueryText = SQLQueryText & _
			"Order by EntityName, xca_COCode.OCN, NXX, RateCenter, SwitchID, Status"
Case "OCN"
SQLQueryText = SQLQueryText & _
			"Order by xca_COCode.OCN, EntityName, NXX, RateCenter, SwitchID, Status"
Case "Exchange"
SQLQueryText = SQLQueryText & _
			"Order by RateCenter, SwitchID, NXX, EntityName, xca_COCode.OCN, Status"
Case "CLLI"
SQLQueryText = SQLQueryText & _
			"Order by SwitchID, RateCenter, NXX, EntityName, xca_COCode.OCN, Status"
Case "Status"
SQLQueryText = SQLQueryText & _
			"Order by Status, NXX, EntityName, xca_COCode.OCN, RateCenter, SwitchID"
Case "NXX"
SQLQueryText = SQLQueryText & _
			"Order by NXX,  EntityName, xca_COCode.OCN, RateCenter, SwitchID, Status"
End Select

SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPACompanyCOCodes = server.createObject("ADODB.recordset")
rstNPACompanyCOCodes.activeConnection = objConnection
rstNPACompanyCOCodes.CursorLocation = adUseServer
rstNPACompanyCOCodes.CursorType = adOpenStatic
rstNPACompanyCOCodes.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS TO DETERMINE IF NPA ENTERED IS VALID

if rstNPACompanyCOCodes.RecordCount>0 then

%>
<TABLE align="center" border="0" width="60%">
<TR>
<TD>
<font face="Arial" size="2">
<b>Instructions:</b><br>
- Click on the <image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/> to sort by that column<br>
- Click on an <b>OCN</b> to filter the results by that OCN<br>
- Click on a <b>CLLI</b> to go to the SwitchID (CLLI) query results page<br>
- Click on an <b>Exchange</b> to filter the results by that exchange
</font>
</TD>
</TR>
</TABLE>

<br>
<TABLE align="center" border="1">

<TD align="center"><B>&nbsp; OCN</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=OCN">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Company</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=Company">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>&nbsp; NXX</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=NXX">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Status</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=Status">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>CLLI</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=CLLI">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
<TD align="center"><B>Exchange</B>
<a href="/xcalogon/NPACompanyCOCodes_Result.asp?NPA=<%=request.querystring("NPA")%>&SortOrder=Exchange">
<image style="border: 0px solid ;" src="/xcalogon/SortImage.gif"/></a><br></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstNPACompanyCOCodes.EOF
%>
<TR>
<TD align="center"> &nbsp;
<a href="/xcalogon/NPACompanyCOCodes_Filter_OCN.asp?NPA=<%=request.querystring("NPA")%>&OCN=<%=rstNPACompanyCOCodes("OCN")%>">
<%=rstNPACompanyCOCodes("OCN")%></a>
&nbsp; </TD>
<TD align="center">&nbsp; <%=rstNPACompanyCOCodes("EntityName")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPACompanyCOCodes("NXX")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPACompanyCOCodes("Status")%> &nbsp;</TD>
<TD align="center">&nbsp; 
<a href="/xcalogon/CNAS_Switch.asp?Switch=<%=rstNPACompanyCOCodes("SwitchID")%>">
<%=rstNPACompanyCOCodes("SwitchID")%></a> &nbsp;</TD>
<TD align="center"> &nbsp;
<a href="/xcalogon/NPACompanyCOCodes_Filter_RC.asp?NPA=<%=request.querystring("NPA")%>&RC=<%=rstNPACompanyCOCodes("RateCenter")%>">
<%=rstNPACompanyCOCodes("RateCenter")%></a> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstNPACompanyCOCodes.movenext
%><%
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
rstNPACompanyCOCodes.Close
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
+"$RCSfile: NPACompanyCOCodes_Result.asp,v $\n"
+"$Revision: 1.4 $\n"
+"$Date: 2004/12/03 17:11:09 $ (UTC)"
alert(strAlertText)
}
</SCRIPT>
<%
' END BLOCK
' ---------
%>
</BODY>
</HTML>

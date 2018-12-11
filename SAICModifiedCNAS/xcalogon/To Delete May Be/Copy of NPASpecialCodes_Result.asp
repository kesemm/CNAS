<%@ Language=VBScript %><%
Response.Buffer = true
Response.Expires=0
%><% ' VI 6.0 Scripting Object Model Enabled %><!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<HTML>
<HEAD>
<META content="no-cache">
<META name="GENERATOR" content="Microsoft FrontPage 3.0">
<TITLE>CNAS NPA Special Codes</TITLE>
<%
'****************************************************************************************
'* CVS File:      $RCSfile: NPASpecialCodes_Result.asp,v $
'* Commit Date:   $Date: 2004/11/26 18:41:58 $ (UTC)
'* Committed by:  $Author: WalshKel $
'* CVS Revision:  $Revision: 1.2 $
'* Checkout Tag:  $Name:  $ (Version/Build)
'**************************************************************************************** 
%><%
UserEntityType=session("UserEntityType")
%>
</HEAD>
<BODY text="black" bgproperties="fixed" bgcolor="#D7C7A4">
<FORM name="thisForm" method="post" id="thisForm"><!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" --></FORM>
<P align="center"><B><BIG>Special Codes in NPA: <%=request.querystring("NPA")%></BIG></B><BR></P>
<BR>
<%
' SET UP THE CONNECTION AND RECORDSET

SQLQueryText = "SELECT [NPA], [NXX], [COStatusDescription] As [Status], PublicRemarks, CNARemarks " & _
                        "FROM xca_COCode Left JOIN xca_status_codes ON xca_COCode.status=xca_status_codes.COStatus " & _
                        "WHERE NPA='" & request.querystring("NPA") & "' " & _
                        "and (Status='T'or Status='U'or Status='N'or Status='Z'or Status='P') " & _
                        "ORDER BY Status, NXX"
SET objConnection = server.createObject("ADODB.Connection")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstNPASpecialCodes = server.createObject("ADODB.recordset")
rstNPASpecialCodes.activeConnection = objConnection
rstNPASpecialCodes.CursorLocation = adUseServer
rstNPASpecialCodes.CursorType = adOpenStatic
rstNPASpecialCodes.open SQLQueryText, objConnection

' CHECK THAT THERE ARE RECORDS TO DETERMINE IF NPA ENTERED IS VALID

if rstNPASpecialCodes.RecordCount>0 then

%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>NPA</B></TD>
<TD align="center"><B>NXX</B></TD>
<TD align="center"><B>Status</B></TD>
<TD align="center"><B>Public Remarks</B></TD>
<TD align="center"><B>CNA Remarks</B></TD>
</TR>
<%
' START LOOPING THROUGH THE RECORDSET UNTIL THE END
do until rstNPASpecialCodes.EOF
%>
<TR>
<TD align="center">&nbsp; <%=rstNPASpecialCodes("NPA")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPASpecialCodes("NXX")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPASpecialCodes("Status")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPASpecialCodes("PublicRemarks")%> &nbsp;</TD>
<TD align="center">&nbsp; <%=rstNPASpecialCodes("CNARemarks")%> &nbsp;</TD>
</TR>
<%
' GET THE NEXT RECORD IN THE SET
rstNPASpecialCodes.movenext
%><%
' LOOP CALL TO END THE LOOP FOR THE RECORDSET
loop
%></TABLE>
<%
 else
%>
<TABLE align="center" border="1">
<TR>
<TD align="center"><B>No Special Codes in NPA: "<%=request.querystring("NPA")%>" or the NPA is Invalid</B></TD>
</TR>
</TABLE>
<%
End if
%><%
' BE A GOOD BOY AND CLOSE THE RECORDSET AND CONNECTION
rstNPASpecialCodes.Close
objConnection.close
%><BR>
<CENTER><FONT color="gray" face="Arial" size="1"><B>SAIC Canada Version Control Information:</B>&nbsp;&nbsp;&nbsp; $Revision: 1.2 $&nbsp;&nbsp;&nbsp; $Date: 2004/11/26 18:41:58 $ (UTC)</FONT></CENTER>
</BODY>
</HTML>

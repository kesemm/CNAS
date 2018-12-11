<%@ Language=VBScript %>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<html>

<head>
<meta HTTP+EQUIV="Pragma" CONTENT="no-cache">
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title>Processed Table</title>
<script language="JavaScript">
   function changeScreenSize()
   {self.moveTo(10,10)
	self.resizeTo(1200,900)}
</script>
<%


UserEntityType=session("UserEntityType")


 %>
</head>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4">

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
<!-- #Include file="ADOVBS.INC" -->

<%
' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT
sqlQry = "SELECT [Month],IsNull(PartOneNumber,0) as PartOneNumber, IsNull(PartOneUpdates,0) as PartOneUpdates, IsNull(PartOneBulkA,0) as PartOneBulkA, IsNull(PartOneBulkB,0) as PartOneBulkB, IsNull(PartFourNumber,0) as PartFourNumber, IsNull(Recovered,0) as Recovered, IsNull(Misc,0) as Misc, IsNull(ESRDNumberPart1,0) as ESRDNumberPart1, IsNull(ESRDNumberUpdates,0) as ESRDNumberUpdates, IsNull(ESRDNumberPart3,0) as ESRDNumberPart3, IsNull(MBIPart1,0) as MBIPart1, IsNull(MBIRecover,0) as MBIRecover, IsNull(MBIUpdate,0) as MBIUpdate, IsNull(MBIPart3,0) as MBIPart3, IsNull(NonGeo,0) As NonGeo, IsNull(NonGeoRecovered,0) As NonGeoRecovered from xca_Processed  Where [Year]=2018 order by [Year] Desc, Sort"
 %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstQry = server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstQry = objConnection.execute(sqlQry)
%>

</form>
<p align="center"><strong>Number of CO Codes Processed by Month</strong></p>
<p>  <br>
</p>
<p>  <br>
</p>
<table ALIGN="CENTER" BORDER="1" CELLPADING="3" CELLSPACING="3" WIDTH="100%">
<tr ALIGN="CENTER">
<th ALIGN="CENTER">Month</th>
<th ALIGN="CENTER">Part One</th>
<th ALIGN="CENTER">Part One Normal Updates</th>
<th ALIGN="CENTER">Part One Bulk A Updates</th>
<th ALIGN="CENTER">Part One Bulk B Updates</th>
<th ALIGN="CENTER">Part Four</th>
<th ALIGN="CENTER">Recovered</th>
<th ALIGN="CENTER">Other</th>
<th ALIGN="CENTER">ESRD Part1</th>
<th ALIGN="CENTER">ESRD Updates</th>
<th ALIGN="CENTER">ESRD Part3</th>
<th ALIGN="CENTER">MBI Part1</th>
<th ALIGN="CENTER">MBI Recover</th>
<th ALIGN="CENTER">MBI Updates</th>
<th ALIGN="CENTER">MBI Part3</th>
<th ALIGN="CENTER">NonGeo</th>
<th ALIGN="CENTER">NonGeo Recover</th>

</tr>
<tr ALIGN="CENTER">
<th ALIGN="CENTER">&nbsp;</th>
<th ALIGN="CENTER">Min 325</th>
<th ALIGN="CENTER">Min 125</th>
<th ALIGN="CENTER">No Min</th>
<th ALIGN="CENTER">No Min</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">Min 20</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">Min 50</th>
<th ALIGN="CENTER">No Min</th>
<th ALIGN="CENTER">No Min</th>
<th ALIGN="CENTER">No Charge</th>
<th ALIGN="CENTER">No Min</th>
<th ALIGN="CENTER">No Min</th>
</tr>
<tr>
</tr>


<% Do Until rstQry.EOF %>
<tr>
<td ALIGN="CENTER"><b><%=rstQry("Month") %></b></td>
<% if Left(rstQry("Month"),5) = "Total" Then %>
<td ALIGN="CENTER"><b><%=rstQry("PartOneNumber") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("PartOneUpdates") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("PartOneBulkA") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("PartOneBulkB") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("PartFourNumber") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("Recovered") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("Misc") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("ESRDNumberPart1") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("ESRDNumberUpdates") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("ESRDNumberPart3") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("MBIPart1") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("MBIRecover") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("MBIUpdate") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("MBIPart3") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("NonGeo") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("NonGeoRecovered") %></b></td>
<% else %>
<td ALIGN="CENTER"><%=rstQry("PartOneNumber") %></td>
<td ALIGN="CENTER"><%=rstQry("PartOneUpdates") %></td>
<td ALIGN="CENTER"><%=rstQry("PartOneBulkA") %></td>
<td ALIGN="CENTER"><%=rstQry("PartOneBulkB") %></td>
<td ALIGN="CENTER"><%=rstQry("PartFourNumber") %></td>
<td ALIGN="CENTER"><%=rstQry("Recovered") %></td>
<td ALIGN="CENTER"><%=rstQry("Misc") %></td>
<td ALIGN="CENTER"><%=rstQry("ESRDNumberPart1") %></td>
<td ALIGN="CENTER"><%=rstQry("ESRDNumberUpdates") %></td>
<td ALIGN="CENTER"><%=rstQry("ESRDNumberPart3") %></td>
<td ALIGN="CENTER"><%=rstQry("MBIPart1") %></td>
<td ALIGN="CENTER"><%=rstQry("MBIRecover") %></td>
<td ALIGN="CENTER"><%=rstQry("MBIUpdate") %></td>
<td ALIGN="CENTER"><%=rstQry("MBIPart3") %></td>
<td ALIGN="CENTER"><%=rstQry("NonGeo") %></td>
<td ALIGN="CENTER"><%=rstQry("NonGeoRecovered") %></td>
<% end if %>
</tr>
<% rstQry.moveNext
loop %>

</table>
<h5>
<%
// Comment out the following section--------------- <br>
//After Jan 1, 2005<br>
%>
--------------- <br>
Part One month indicates the date Part 3 issued.<br>
Part One Normal Updates includes CLLI updates, Transfers, Effective Date changes etc. Part 3 issued.<br>
Part One Bulk A Updates (One or two items are changed in both CNA and Telcrodia databases. All changes are the identical). Part 3 issued.<br>
Part One Bulk B Updates (More then two items are changed or items are different. Effects both CNA and Telcordia databases). Part 3 issued.<br>
Part Four month indicates In-Service date.<br>
Recovered month indicates the recovered date.  Part 3 issued.<br>
Other requests are items not covered by Part 1, CO Code Updates (including Transfers).  No Part 3 issued.<br>
ESRD Part1 month indicates the date Part 2 issued.<br>
ESRD Part3 month indicates the date Part 3 received.<br>
MBI Part1 month indicates the date Part 2 issued.<br>
MBI Part3 month indicates the date Part 3 received.<br>
<%
//---------------<br>
//Before Jan 1, 2005<br>
//--------------- <br>
//Part One month indicates the application date. Part 3 issued.<br>
//Part One Normal Updates includes CLLI updates, Transfers, Effective Date changes etc. Part 3 issued.<br>
//Part One Bulk A Updates (One or two items are changed in both CNA and Telcrodia databases. All changes are the identical). Part 3 issued.<br>
//Part One Bulk B Updates (More then two items are changed or items are different. Effects both CNA and Telcordia databases). Part 3 issued.<br>
//Part Four month indicates In-Service date.<br>
//Recovered month indicates the recovered date.  Part 3 issued.<br>
//Other requests are items not covered by Part 1, CO Code Updates (including Transfers).  No Part 3 issued.<br>
//ESRD Part1 month indicates the application date. Part 2 issued.<br>
//ESRD Part3 month indicates the date Part 3 received.<br>
%>
<p>  <br>
</body>
</html>

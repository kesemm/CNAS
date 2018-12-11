<%@ Language=VBScript %>
<HTML><HEAD>
<title>Period Data</title>
<%
Response.Buffer = true
Response.Expires=0
%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<meta HTTPEQUIV="Pragma" CONTENT="no-cache"+>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<script language="JavaScript">
   function changeScreenSize()
   {self.moveTo(10,10)
	self.resizeTo(1200,900)}
</script>
<%


UserEntityType=session("UserEntityType")


 %>
</HEAD>

<body text="black" bgProperties="fixed" bgColor="#d7c7a4"  onload="changeScreenSize()">

<%
' SET THE SORT ORDER PORTION OF THE QUERY TEXT BASED ON INPUT
sqlQry = "SELECT  [Period], IsNull(CAST([FiscalYear] as VarChar),'') as [FiscalYear], [CalendarYear], COCodeNormal, COCodeUpdates, IsNull(COCodeBulkA,0) As COCodeBulkA, IsNull(COCodeBulkB,0) As COCodeBulkB, IsNull(COCodeRecovered,0) As COCodeRecovered, IsNull(MBIPart1,0) As MBIPart1, IsNull(MBIRecover,0) As MBIRecover, IsNull(MBIUpdate,0) As MBIUpdate, IsNull (NonGeo,0) As NonGeo, IsNull(NonGeoRecovered,0) As NonGeoRecovered  From Period_Data Where [CalendarYear]=2018 Order by [CalendarYear] Desc, Sort"
 %>



<%
SET objConnection = server.createobject("ADODB.connection")
SET rstQry = server.createobject("ADODB.recordset")
objConnection.open "DSN=cnasadmin;UID=admin;PWD=cnasadmin"
SET rstQry = objConnection.execute(sqlQry)
%>

<form name="thisForm" METHOD="post">
<!--#include file="xca_CNASLib.inc"-->
</form>
<p align="center"><strong>Number of Codes Processed by Period</strong></p>
<p>  <br>
</p>
<p>  <br>
</p>
<table ALIGN="center" BORDER="1" CELLPADING="3" CELLSPACING="3" WIDTH="100%">
<tr ALIGN="center">
<th ALIGN="center">Period</th>
<th ALIGN="center">Fiscal Year</th>
<th ALIGN="center">Calendar Year</th>
<th ALIGN="center">COCode Normal</th>
<th ALIGN="center">COCode Updates</th>
<th ALIGN="center">COCode Bulk A</th>
<th ALIGN="center">COCode Bulk B</th>
<th ALIGN="center">COCode Recovered</th>
<th ALIGN="center">MBI Part1</th>
<th ALIGN="center">MBI Recover</th>
<th ALIGN="center">MBI Updates</th>
<th ALIGN="center">NonGeo</th>
<th ALIGN="CENTER">NonGeo Recover</th>
</tr>
<tr>
<th ALIGN="center">&nbsp; </th>
<th ALIGN="center"> &nbsp;</th>
<th ALIGN="center"> &nbsp;</th>
<th ALIGN="center">Min 325</th>
<th ALIGN="center">Min 125</th>
<th ALIGN="center">No Min</th>
<th ALIGN="center">No Min</th>
<th ALIGN="center">Min 20</th>
<th ALIGN="center">Min 50</th>
<th ALIGN="center">No Min</th>
<th ALIGN="center">No Min</th>
<th ALIGN="center">No Min</th>
<th ALIGN="center">No Min</th>
</tr>
<tr>
</tr>

<% Do Until rstQry.EOF %>
<tr>
<td ALIGN="center" "no wrap"><b><%=rstQry("Period")%></b></td>
<td ALIGN="center"><b><%=rstQry("FiscalYear") %></b></td>
<% if Left(rstQry("Period"),5) = "Total" Then %>
<td ALIGN="center"><b><%=rstQry("CalendarYear") %></b></td>
<td ALIGN="center"><b><%=rstQry("COCodeNormal") %></b></td>
<td ALIGN="center"><b><%=rstQry("COCodeUpdates") %></b></td>
<td ALIGN="center"><b><%=rstQry("COCodeBulkA") %></b></td>
<td ALIGN="center"><b><%=rstQry("COCodeBulkB") %></b></td>
<td ALIGN="center"><b><%=rstQry("COCodeRecovered") %></b></td>
<td ALIGN="center"><b><%=rstQry("MBIPart1") %></b></td>
<td ALIGN="center"><b><%=rstQry("MBIRecover") %></b></td>
<td ALIGN="center"><b><%=rstQry("MBIUpdate") %></b></td>
<td ALIGN="center"><b><%=rstQry("NonGeo") %></b></td>
<td ALIGN="CENTER"><b><%=rstQry("NonGeoRecovered") %></b></td>
<% else %>
<td ALIGN="center"><%=rstQry("CalendarYear") %></td>
<td ALIGN="center"><%=rstQry("COCodeNormal") %></td>
<td ALIGN="center"><%=rstQry("COCodeUpdates") %></td>
<td ALIGN="center"><%=rstQry("COCodeBulkA") %></td>
<td ALIGN="center"><%=rstQry("COCodeBulkB") %></td>
<td ALIGN="center"><%=rstQry("COCodeRecovered") %></td>
<td ALIGN="center"><%=rstQry("MBIPart1") %></td>
<td ALIGN="center"><%=rstQry("MBIRecover") %></td>
<td ALIGN="center"><%=rstQry("MBIUpdate") %></td>
<td ALIGN="center"><b><%=rstQry("NonGeo") %></b></td>
<td ALIGN="CENTER"><%=rstQry("NonGeoRecovered") %></td>
<% end if %>
</b></tr>
<% rstQry.moveNext
loop %>

</table>
<h5>
<p>  <br></p></h5>
</body></HTML>

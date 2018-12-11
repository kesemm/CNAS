<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  This is file is based on, but has undergone extensive modifications:
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  CVS File:      $RCSfile:$
  Commit Date:   $Date:$ (UTC)
  Committed by:  $Author:$
  CVS Revision:  $Revision:$
  Checkout Tag:  $Name(Version/Build)

  Purpose:  This page lists the open problem for the rep or another selected rep.
  -->

  <!-- 	#include file = "public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = Cint(Session("lhd_sid"))

  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "OpenProblems")%></title>
  <link rel="stylesheet" type="text/css" href="default.css">
</head>
<body>

<%
	' Determine if a rep_id has been entered by form or in the
	' URL.  If not, just query on the rep's own id.
	Dim rep_id
	if Len(Request.QueryString("rep_id")) > 0 Then
		rep_id = Cint(Request.QueryString("rep_id"))
	Elseif Len(Request.Form("rep_id")) > 0 Then
		rep_id = Cint(Request.Form("rep_id"))
	Else
		rep_id = sid
	End If

	' Get the rep's name that is being queryed on
	Dim ruid, repRes
  Set repRes = SQLQuery(cnnDB, "SELECT username FROM tblUsers WHERE sid=" & rep_id )
	ruid = repRes("username")

	' Query the database for the problems.
	Dim listStr, cntStr, listRes
	listStr = "SELECT TOP 100 p.id, p.title, p.start_date, r.fname,p.due_date,pri.pname, s.sname,c.cname " & _
	"FROM ((problems AS p " & _
	"INNER JOIN tblUsers AS r ON p.assigned_to = r.sid) " & _
	"INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
	"INNER JOIN status AS s ON p.status = s.status_id " & _
	"INNER JOIN categories AS c ON p.category = c.category_id " & _

	"WHERE"

	cntStr = "SELECT count(*) AS total FROM problems WHERE"
	Dim disp_total

	' If a problem ID is entered, search only for that
	If Len(Request.QueryString("id"))>0 Then
		listStr = listStr & " p.id=" & Request.QueryString("id")
		disp_total = FALSE
	Else
'		listStr = listStr & " p.status<>" & Cfg(cnnDB, "CloseStatus") & " AND p.assigned_to=" & rep_id
'		cntStr = cntStr & " status<>" & Cfg(cnnDB, "CloseStatus") 

		listStr = listStr & " p.status<>" & Cfg(cnnDB, "CloseStatus") & " AND p.assigned_to=" & rep_id
		cntStr = cntStr & " status<>" & Cfg(cnnDB, "CloseStatus") & " AND assigned_to=" & rep_id

		disp_total = TRUE
	End If
  ' Determine Sort Order
  Dim intSort, intOrder, intIDOrder, intTitleOrder, intUIDOrder, intDateOrder, intDueDateOrder,intCatOrder,intPriOrder, intStatusOrder
  intSort = Cint(Request.QueryString("sort"))
  If Len(Request.QueryString("order")) > 0 Then
    intOrder = Cint(Request.QueryString("order"))
  Else
    intOrder = 0
  End If
  Select Case intSort
    Case 1  ' id
      listStr = listStr & " ORDER BY p.id"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intIDOrder = 1
      Else
        listStr = listStr & " ASC"
        intIDOrder = 0
      End If
    Case 2  ' title
      listStr = listStr & " ORDER BY p.title"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intTitleOrder = 1
      Else
        listStr = listStr & " DESC"
        intTitleOrder = 0
      End If
    Case 3  ' username
      listStr = listStr & " ORDER BY r.fname"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intUIDOrder = 1
      Else
        listStr = listStr & " DESC"
        intUIDOrder = 0
      End If
    Case 4  ' start_date
      listStr = listStr & " ORDER BY p.start_date"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intDateOrder = 1
      Else
        listStr = listStr & " ASC"
        intDateOrder = 0
      End If
    Case 5  ' due_date
      listStr = listStr & " ORDER BY p.due_date"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intDueDateOrder = 1
      Else
        listStr = listStr & " DESC"
        intDueDateOrder = 0
      End If
    Case 6  ' category
      listStr = listStr & " ORDER BY c.cname"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intCatOrder = 1
      Else
        listStr = listStr & " ASC"
        intCatOrder = 0
      End If
    Case 7  ' priority
      listStr = listStr & " ORDER BY p.priority"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intPriOrder = 1
      Else
        listStr = listStr & " ASC"
        intPriOrder = 0
      End If
    Case 8  ' status
      listStr = listStr & " ORDER BY p.status"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intStatusOrder = 1
      Else
        listStr = listStr & " ASC"
        intStatusOrder = 0
      End If
    Case Else ' due date
      listStr = listStr & " ORDER BY p.due_date"
      If intOrder = 1 Then
        listStr = listStr & " DESC"
        intIDOrder = 0
      Else
        listStr = listStr & " ASC"
        intIDOrder = 1
      End If
  End Select

  Set listRes = SQLQuery(cnnDB, listStr)

	' Get a total number of problems returned
  Dim cntRes, start
	If disp_total Then
		Set cntRes = SQLQuery(cnnDB, cntStr)
	End If

	' If not empty results, set up the page.  Only display
	' 10 results per page.
	If Not listRes.EOF Then
	Dim Counter, numToDisplay, startNum
	Counter = 1
	If Len(Request.QueryString("num")) > 0 Then
		numToDisplay = CInt(Request.QueryString("num"))
	Else
		numToDisplay = 100
	End if
	If Len(Request.QueryString("start")) > 0 Then
		start = CInt(Request.QueryString("start"))
	Else
		start = 1
	End if

  Dim strColumns, intUseInoutBoard
'  intUseInoutBoard = Cfg(cnnDB, "UseInoutBoard")
'  If intUseInoutBoard = 1 Then
    strColumns = 8
'  Else
    strColumns = 8
'  End If
  
%>
<div align="center">
  <table class="Wide">
  <tr class="Head1">
    <td colspan="<%=strColumns%>">
      <%=lang(cnnDB, "Problemsfor")%>&nbsp;<% = listRes("fname") %>
      <%
        If disp_total Then
          Response.Write("&nbsp;(" & lang(cnnDB, "Total") & ":" & cntRes("total") & ")")
          cntRes.Close
        End If
      %>
    </td>
  </tr>
  <tr align="center" Class="Head2">
    <td nowrap><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=1&order=<% = intIDOrder %>" class="HeadLink"><%=lang(cnnDB, "ID")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=2&order=<% = intTitleOrder %>" class="HeadLink"><%=lang(cnnDB, "Title")%></a></td>
    <td nowrap><%=lang(cnnDB, "AssignedTo")%></td>
    <td nowrap><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=4&order=<% = intDateOrder %>" class="HeadLink"><%=lang(cnnDB, "DateSubmitted")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=5&order=<% = intDueDateOrder %>" class="HeadLink"><%=lang(cnnDB, "Duedate")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=6&order=<% = intCatOrder %>" class="HeadLink"><%=lang(cnnDB, "Category")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=7&order=<% = intPriOrder %>" class="HeadLink"><%=lang(cnnDB, "Priority")%></a></td>
    <td><%=lang(cnnDB, "Status")%></td>
</tr>
  <%
    Do While Not (listRes.EOF) AND (Counter <= (numToDisplay + start - 1))
    If Counter >= start Then
  %>
    <tr align="center" valign="center" class="Body1c">
      <td nowrap><% = listRes("id") %></td>
      <td><A HREF="details.asp?id=<% = listRes("id") %>"><% = listRes("title") %></A></td>
      <td nowrap><% = listRes("fname") %></td>
      <td nowrap><% = MachineFormatDate(listRes("start_date")) %></td>
      <td nowrap><% = MachineFormatDate(listRes("due_date")) %></td>
      <td nowrap><% = listRes("cname") %></td>
      <td nowrap><% = listRes("pname") %></td>
      <td nowrap><% = listRes("sname") %></td>

    </tr>
  <%
    End If
    Counter = Counter + 1
    listRes.MoveNext
    Loop
    Response.Write("</table></center>")

    ' Calculate prev/next page links
    Dim startP, StartN
    startP = start - numToDisplay
    If startP < 1 Then
      startP = 1
    End if
    startN = start + numToDisplay
  %>
    <div align="center">
    <% If start > 1 Then %>
      <A HREF="view.asp?rep_id=<% = rep_id %>&start=<% = startP %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>">Previous</A>&nbsp;
    <% End If
      If Not (listRes.EOF) Then
    %>
      <A HREF="view.asp?rep_id=<% = rep_id %>&start=<% = startN %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>">Next</A>
    <% End If %>
    </div>
  <%

    ' If no results returned:
    Else
  %>
  <div align="center">
  <table border="0" cellspacing="3" cellpadding="5" width="600">
  <tr class="Head1">
    <td colspan="6" valign="center">
      <font size="+2"><b><%=lang(cnnDB, "OpenProblemsfor")%>&nbsp;<% = ruid %></b></font>
    </td>
  </tr>
  <tr align="center" Class="Head2">
    <td nowrap><%=lang(cnnDB, "ID")%></td>
    <td><%=lang(cnnDB, "Title")%></td>
    <td nowrap><%=lang(cnnDB, "AssignedTo")%></td>
    <td nowrap><%=lang(cnnDB, "DateSubmitted")%></td>
    <td><%=lang(cnnDB, "Duedate")%></td>
    <td><%=lang(cnnDB, "Status")%></td>
  </tr>
  <tr align="center" class="Body1">
    <td colspan="6">
      <%=lang(cnnDB, "Noresultsfound")%>.
    </td>
  </tr>
  </table>
  </div>
<%	End If

	' Close results
	listRes.Close
	repRes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>
</html>



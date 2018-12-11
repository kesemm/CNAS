<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Response.Expires = -1
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
 
  Purpose:  This page displays the task details and allows users to modify/update tasks.
 -->

  <!-- 	#include file = "public.asp" -->
  <% 
    Dim cnnDB, sid, logon_fname, To_email,CC_email
    Set cnnDB = CreateCon
    sid = Cint(Session("lhd_sid"))
    logon_fname = session("logon_fname")
  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "EditProblem")%></title>
  <link rel="stylesheet" type="text/css" href="default.css">
<script type="text/javascript" language="JavaScript" src="popcalendar.js"></script>

</head>
<body>

<%
	' Get the problem ID
	Dim id, blnUpdate, strUpdateMessage
      id = Session("id")
      If id = 0 Then
        If Len(Request.Form("id")) = 0 Then
          Call DisplayError(3, "A problem ID number is required.")
        End If
        id = Cint(Request.Form("id"))
      End If
  If Cint(Request.Form("update")) = 1 Then
    blnUpdate = True
  Else 
    blnUpdate = False
  End If

  Dim category, title, description,old_time_spent,entered_by_id,OldDueDate,NewDueDate,quantity
  Dim priority, status, assigned_to, time_spent, solution, notes, startdate, close_date,due_date

  ' ===============================
  ' Update the problem
  If blnUpdate Then

    ' Get the problem data from the form fields.
    id = Request.Form("id")
    category = Cint(Request.Form("category"))
    title = Request.Form("title")
    priority = Cint(Request.Form("priority"))
    status = Cint(Request.Form("status"))
    assigned_to = Cint(Request.Form("assigned_to"))
    time_spent = Cint(Request.Form("time_spent"))
    due_date = SQLDate(Request.Form("duedate"),lhdAddSQLDelim)
    solution = Request.Form("solution")
    notes = Request.Form("notes")
    quantity=Request.Form("quantity")
' Check for required fields
    if IsDate(Request.Form("duedate")) = False Then
    Call DisplayError(3, lang(cnnDB, "InvalidDate"))
    end if
'
    if Len(title)=0 Then
      Call DisplayError(1, Lang(cnnDB, "Title"))
    End if
'
    if (status=Cfg(cnnDB, "CloseStatus")) and (Len(solution)=0) Then
      Call DisplayError(1, Lang(cnnDB, "Solution"))
    End if
' 
    title = Left(Trim(title), 50)
    title = Replace(title, "'", "''")
'
  ' Grab original description
    Dim rstDesc, rstEnteredby,strChangeNotes
    Set rstDesc = SQLQuery(cnnDB, "SELECT entered_by_id,category, assigned_to, status, priority,time_spent,due_date FROM problems WHERE id=" & id)
    If time_spent < rstDesc("time_spent")+ 1 Then
      Call DisplayError(1,Lang(cnnDB,"TimeSpentUpdate"))
    End if
' Insert actions into strChangeNotes
    If (category <> rstDesc("category")) OR _
      (assigned_to <> rstDesc("assigned_to")) OR (status <> rstDesc("status")) OR _
      (priority <> rstDesc("priority")) Then
'
'        If DateDiff("d",FormatDateTime(rstDesc("due_date"),1),FormatDateTime(Request.Form("duedate"),1)) > 0  Then
'          strChangeNotes = strChangeNotes & Lang(cnnDB, "DUEDATE_2") & ": " & FormatDateTime(rstDesc("due_date"),1) & " => " & Request.Form("duedate") & vbNewLine
'        End If
'
        If priority <> rstDesc("priority") Then

          Dim newPri, oldPri
          Set newPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & priority)
          Set oldPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & rstDesc("priority"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "PRIORITY_2") & ": " & oldPri("pname") & " => " & newPri("pname") & vbNewLine
          newPri.Close
          oldPri.Close
        End If
'
        If assigned_to <> rstDesc("assigned_to") Then

          Dim newassigned_to, oldassigned_to
          Set newassigned_to = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE sid=" & assigned_to)
          Set oldassigned_to = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE sid=" & rstDesc("assigned_to"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "TRANSFERREPS") & ": " & oldassigned_to("fname") & " => " & newassigned_to("fname") & vbNewLine
          newassigned_to.Close
          oldassigned_to.Close
        End If
'
        If category <> rstDesc("category") Then

          Dim newCat, oldCat
          Set newCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & category)
          Set oldCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & rstDesc("category"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "CATEGORY_2") & ": " & oldCat("cname") & " => " & newCat("cname") & vbNewLine
          newCat.Close
          oldCat.Close
        End If


        If status <> rstDesc("status") Then

          Dim newStat, oldStat
          Set newStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & status)
          Set oldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & rstDesc("status"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "STATUS_2") & ": " & oldStat("sname") & " => " & newStat("sname") & vbNewLine
          newStat.Close
          oldStat.Close
        End If
    End If

' Update the Notes
    Dim rstUpdateNotes, intPrivate, dtNoteDate, blnSendUpdateMsg
    blnSendUpdateMsg = False

    dtNoteDate = SQLDate(Now, lhdAddSQLDelim)
    If Len(notes)>0 Then
      notes = Replace(notes, "'", "''")
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, updated_by_id) " & _
        "VALUES (" & id & ", '" &  notes & "', " & dtNoteDate & ", " &  sid & ")")
    End If
    If Len(strChangeNotes)> 0 Then
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, updated_by_id) " & _
        "VALUES (" & id & ", '" &  strChangeNotes & "', " & dtNoteDate & ", " & sid & ")")
    End If
  ' Get missing variables to enter problem

    Dim cname, catRes
    Set catRes = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & Request.Form("category"))
    cname = catRes("cname")

    Dim probRes
    Set probRes = SQLQuery(cnnDB, "SELECT start_date FROM problems WHERE id=" & id)
    startdate = probRes("start_date")

  ' Get the old priority
    Dim old_priority, oldPriRes
    Set oldPriRes = SQLQuery(cnnDB, "SELECT priority FROM problems WHERE id=" & id)
    old_priority = oldPriRes("priority")
    oldPriRes.Close

  ' Remove apostrophes
    description = Replace(description, "'", "''")
    solution = Replace(solution, "'", "''")

  ' All data is present
  ' Write problem into database

    Dim probStr
    probStr = "UPDATE problems SET " & _
      "quantity=" & quantity & ", " & _
      "category=" & category & ", " & _
      "title='" & title & "', " & _
      "priority=" & priority & ", " & _
      "status=" & status & ", " & _
      "assigned_to=" & assigned_to & ", " & _
      "time_spent=" & time_spent & ", " & _
      "due_date=" & due_date & ", " & _
      "solution='" & solution & "'"
 ' Add the closed date/time if the problem is closed
    If status = Cfg(cnnDB, "CloseStatus") Then
      probStr = probStr & ", close_date=" & SQLDate(Now, lhdAddSQLDelim)
    End If
    strUpdateMessage = Lang(cnnDB, "Theproblemhasbeensaved") & "."

    probStr = probStr & " WHERE id=" & id

    Set probRes = SQLQuery(cnnDB, probStr)
'
' Send an e-mail to the right individual(s)
'
' The task is closed
'
If status = Cfg(cnnDB, "CloseStatus") Then
'
	If (rstDesc("entered_by_id") <> sid and rstDesc("assigned_to") <> sid and rstDesc("entered_by_id") <>  rstDesc("assigned_to")) Then
'
'  The task was closed by somebody who did not open it or is currently
'  not assigned.
'
		To_email=Usr(cnnDB, assigned_to, "email")
		CC_email=Usr(cnnDB,rstDesc("entered_by_id"),"email") & "," & Usr(cnnDB, sid, "email")
'
	ElseIf (rstDesc("entered_by_id") <> rstDesc("assigned_to") and (rstDesc("entered_by_id")=sid or rstDesc("assigned_to")=sid)) Then
'
'  The task was closed by either the individual who opened it or was the currently assigned.
'
		To_email=Usr(cnnDB, assigned_to, "email")
		CC_email=Usr(cnnDB, rstDesc("entered_by_id"), "email")
	Else
'
'  The task was opened and currently assigned and closed by the same individual.
'
		To_email=Usr(cnnDB, rstDesc("entered_by_id"), "email")
		CC_email=Usr(cnnDB, rstDesc("entered_by_id"), "email")
'
	End If
'
    	Call eMessage(cnnDB, "repclose", id,To_email ,CC_email)
End If

      'Send mail to the appropriate rep for transfered problems
If status <> Cfg(cnnDB, "CloseStatus") Then
      If (assigned_to <> rstDesc("assigned_to")) Then
'
' The task was transferred to a 'new' individual
'
        Call eMessage(cnnDB, "reptransfer", id, Usr(cnnDB, assigned_to, "email"),Usr(cnnDB, sid, "email"))
      Else

'
'  The task was updated but not transferred
'
		If (rstDesc("Entered_by_id") <> sid and rstDesc("assigned_to") <> sid and rstDesc("Entered_by_id") <> rstDesc("assigned_to")) Then
'
'  The task was updated by somebody who did not open it or is currently not assigned.
'
			To_email=Usr(cnnDB, assigned_to, "email")
			'CC_email=Usr(cnnDB,rstEnteredBy("sid"),"email") & "," & Usr(cnnDB, sid, "email")
			CC_email=Usr(cnnDB,rstDesc("Entered_by_id"),"email") & "," & Usr(cnnDB, sid, "email")
		ElseIf (rstDesc("entered_by_id") <> rstDesc("assigned_to") and (rstDesc("entered_by_id")=sid or rstDesc("assigned_to")=sid)) Then
'
'  The task was updated by either the individual who opened it or is currently assigned.
'
			To_email=Usr(cnnDB, assigned_to, "email")
			CC_email=Usr(cnnDB, rstDesc("entered_by_id"), "email")
'
		Else
'
'  The task was updated by the same individual who opened it and is currently assigned.
'
		To_email=Usr(cnnDB, sid, "email")
		CC_email=Usr(cnnDB, sid, "email")
'
		End If
'
      	Call eMessage(cnnDB, "repupdate", id, To_email,CC_email)
      End If
    End If
End If
  ' ===============================
  
  If Cint(Request.QueryString("reopen")) = 1 Then
    Dim strSQLOpen, rstOpenProbUpd, rstOpenNotes, dtOpenNoteDate, strOpenNote
    Dim rstOpenOldStat, rstOpenNewStat, rstProblem
'
' Determine original individual(s) involved with this task.
'
   Set rstProblem = SQLQuery(cnnDB, "SELECT entered_by_id, assigned_to FROM problems WHERE id=" & id)

    strSQLOpen = "UPDATE problems SET " & _
      "assigned_to = " & sid & ", " & _
      "status = " & Cfg(cnnDB, "DefaultStatus") & ", " & _
      "close_date = NULL"

    strSQLOpen = strSQLOpen & " WHERE id = " & id
    Set rstOpenProbUpd = SQLQuery(cnnDB, strSQLOpen)
    Set rstOpenNewStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "DefaultStatus"))
    Set rstOpenOldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "CloseStatus"))
    strOpenNote = Lang(cnnDB, "STATUS_2") & ": " & rstOpenOldStat("sname") & " => " & rstOpenNewStat("sname") & vbNewLine
    dtOpenNoteDate = SQLDate(Now, lhdAddSQLDelim)
    Set rstOpenNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, updated_by_id) " & _
        "VALUES (" & id & ", '" &  strOpenNote & "', " & dtOpenNoteDate & ", " & sid & ")")
'
	If (rstProblem("entered_by_id") <> sid and rstProblem("assigned_to") <> sid and rstProblem("entered_by_id") <> rstProblem("assigned_to")) Then
'
'  The task was re-opened by somebody who did not open it or was not the last assigned.
'
		To_email=Usr(cnnDB, sid, "email")
		CC_email=Usr(cnnDB,rstProblem("entered_by_id"),"email") & "," & Usr(cnnDB, rstProblem("assigned_to"), "email")
'
	Call eMessage(cnnDB, "reopen", id, To_email,CC_email)
'
	ElseIf (rstProblem("entered_by_id") <> rstProblem("assigned_to") and rstProblem("entered_by_id")= sid ) Then
'
'  The task was re-opned by the individual who opened it and they were not the last assigned.
'
		To_email=Usr(cnnDB, sid, "email")
		CC_email=Usr(cnnDB, rstProblem("assigned_to"), "email")
'
	Call eMessage(cnnDB, "reopen", id, To_email,CC_email)
'
	ElseIf (rstProblem("entered_by_id") <> rstProblem("assigned_to") and rstProblem("assigned_to")= sid ) Then
'
'  The task was re-opned by the individual who was the last assigned and not the inidividual who opened it.
'
		To_email=Usr(cnnDB, sid, "email")
		CC_email=Usr(cnnDB, rstProblem("entered_by_id"), "email")
'
	Call eMessage(cnnDB, "reopen", id, To_email,CC_email)
'
	Else
'
'  The task was re-opened and last assigned by the same individual.
'
		To_email=Usr(cnnDB, sid, "email")
		CC_email=Usr(cnnDB, sid, "email")
'
	Call eMessage(cnnDB, "reopen", id, To_email,CC_email)
'
	End If
'
    rstProblem.Close	
    rstOpenNewStat.Close
    rstOpenOldStat.Close

  End If

  ' ===============================

	' Query the database for the problem info
  Dim rstProb, rstSol, rstNotes, strProbQuery
  strProbQuery = "SELECT entered_by_id, time_spent," & _
		"category, status, priority, assigned_to,  start_date, due_date, close_date, title, description,quantity " & _
		"FROM problems WHERE id=" & id

	Set rstProb = SQLQuery(cnnDB, strProbQuery)
  If rstProb.EOF Then
    Call DisplayError(3, "Problem " & id & " could not be found in the database.")
  End If

	' Query for the solution seperately becuase SQL only
	' supports 1 blob per query
	Set rstSol = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)

	time_spent=Cint(rstProb("time_spent"))
	category = Cint(rstProb("category"))
	status = Cint(rstProb("status"))
	priority = Cint(rstProb("priority"))
	startdate = rstProb("start_date")
	close_date = rstProb("close_date")
      due_date = rstProb("due_date")
	title = rstProb("title")
	description = rstProb("description")
	quantity=rstProb("quantity")


  ' Get the Notes for this problem
  Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")

  ' Get the solution and replace characters to make
	' it more readable.
	solution = rstSol("solution")

  Dim strTextDisable, strListDisable
%>


<form action="details.asp" method="POST" id="Form1">
<input type="hidden" name="id" value="<% = id %>">
<input type="hidden" name="oldassigned_to" value="<% = assigned_to %>">
<input type="hidden" name="update" value="1">

<div align="center">
<table class="Normal">
<tr>
	<td colspan="2" align="right">
	  <a href="print.asp?id=<% = id %>" target="printwindow"><%=lang(cnnDB, "PrinterFriendly")%></a>
	</td>
</tr>
<tr class="Head1">
	<td colspan="2">
		<%=lang(cnnDB, "EditProblem")%>&nbsp;<% = id %>
	</td>
</tr>

<% If blnUpdate Then %>
    <tr class="Head2">
      <td colspan="2">
        <div align="center">
          <% = strUpdateMessage %>
        </div>
      </td>
    </tr>
<% End If %>
<tr class="Body1">
	<td valign="top" >

    <table class="Normal">
      <tr>
        <td>
          <b><%=lang(cnnDB, "Category")%>:</b>
        </td>
        <td>
          <SELECT NAME="category" <% = strListDisable %>>
          <%
            Dim rstCat
            Set rstCat = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY category_id ASC")
            If Not rstCat.EOF Then
            Do While Not rstCat.EOF
            If rstCat("category_id") = category Then
            %>
            <OPTION VALUE="<% = rstCat("category_id")%>" SELECTED>
            <% = rstCat("cname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstCat("category_id")%>">
            <% = rstCat("cname") %></OPTION>

          <% 	End If
            rstCat.MoveNext
            Loop
            End If
          %>
          </SELECT>
        </td>
      </tr>
<tr>
        <td>
          <b><%=lang(cnnDB, "Quantity")%>:</b>
        </td>
        <td>
          <input type="text" size="4" name="quantity" value="<% = quantity %>")
        </td>
      </tr>

 <tr>

      <tr>
        <td>
          <b><%=lang(cnnDB, "Status")%>:</b>
        </td>
        <td>
          <SELECT NAME="status" <% = strListDisable %>>
          <%
            Dim rstStat
            Set rstStat = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
            If Not rstStat.EOF Then
            Do While Not rstStat.EOF
            If rstStat("status_id") = status Then
            %>
            <OPTION VALUE="<% = rstStat("status_id")%>" SELECTED>
            <% = rstStat("sname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstStat("status_id")%>">
            <% = rstStat("sname") %></OPTION>

          <% 	End If
            rstStat.MoveNext
            Loop
            End If
          %>
          </SELECT>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Priority")%>:</b>
        </td>
        <td>
          <SELECT NAME="priority" <% = strListDisable %>>
          <%
            Dim rstPri
            Set rstPri = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
            If Not rstPri.EOF Then
            Do While Not rstPri.EOF
            If rstPri("priority_id") = priority Then
            %>
            <OPTION VALUE="<% = rstPri("priority_id")%>" SELECTED>
            <% = rstPri("pname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstPri("priority_id")%>">
            <% = rstPri("pname") %></OPTION>

          <% 	End If
            rstPri.MoveNext
            Loop
            End If
          %>
          </SELECT>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "AssignTo")%>:</b>
        </td>
        <td>
          <SELECT NAME="assigned_to" <% = strListDisable %>>
          <%
            Dim rstRes
            Set rstRes = SQLQuery(cnnDB, "SELECT * From tblUsers ORDER BY username ASC")
            If Not rstRes.EOF Then
            Do While Not rstRes.EOF
            If rstRes("sid") = rstProb("assigned_to") Then
            %>

            <OPTION VALUE="<% = rstRes("sid")%>" SELECTED>
            <% = rstRes("fname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstRes("sid")%>">
            <% = rstRes("fname") %></OPTION>

          <% 	End If
            rstRes.MoveNext
            Loop
            End If
          %>
          </SELECT>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "TimeSpent")%>:</b>
        </td>
        <td>
          <input type="text" size="4" name="time_spent" value="<% = time_spent %>" <% = strTextDisable %>>(<%=lang(cnnDB, "minutes")%>)
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "StartDate")%>:</b>
        </td>
        <td>
          <% = FormatDateTime(startdate,1) %>
        </td>
      </tr>
			<tr>
                    <td>
                      <b><%=lang(cnnDB, "Duedate")%>:</b>
                    </td>
<td>
<!-- New Calendar -->
<input type=text name='duedate' size=20 maxlength=20 value='<%= FormatDateTime(due_date,1)%>'>

<a href="javascript:showCalendar(Form1.duedate, Form1.duedate, 'mmm dd, yyyy',null,1,null,null)"><img src="calendaropen.gif"</img></a>
</td>
</tr>

<!-- End New Calendar -->

      <tr>
        <td>
          <b><%=lang(cnnDB, "CloseDate")%>:</b>
        </td>
        <td>
<% If Len(Trim(close_date)) > 0 Then %>
<% =FormatDateTime(close_date,1) %>
<% Else %>
<% =DisplayDate(close_date,0) %>
<% end if %>
        </td>
      </tr>

    </table>
    <% If status = Cfg(cnnDB, "CloseStatus") Then %>
      <div align="center">
        <b><a href="details.asp?id=<% = id %>&reopen=1"><%=lang(cnnDB, "ReopenProblem")%></a></b>
      </div>
    <% End If %>
	</td>
</tr>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "ProblemInformation")%>:
  </td>
</tr>
<tr class="Body1">
	<td colspan="2" >
		<b><%=lang(cnnDB, "Title")%>:</b><br />
    <input type="text" name="title" size="50" value="<% = title %>"  <% = strTextDisable %>>
		
    <p>
		<b><%=lang(cnnDB, "Description")%>:</b><br />
		<textarea readonly rows="8" cols="80" name="disp_description"><% = description %></textarea>
	</td>
</tr>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "Notes")%>:
  </td>
</tr>
<tr class="Body1">
  <td colspan="2" >
    <% If rstNotes.EOF Then %>
       <%=lang(cnnDB, "NoAvailableNotes")%>
     <% 
      Else
      Do While Not rstNotes.EOF
      
        Response.Write("<b>[")
        Response.Write(DisplayDate(rstNotes("addDate"), lhdDateTime) & " - " & Usr(cnnDB,rstNotes("updated_by_id"),"fname") & "]")
        Response.Write("</b><br />" & vbNewLine)
        Response.Write(Replace(rstNotes("note"), vbNewLine, "<br />"))
        Response.Write("<p>" & vbNewLine)

        rstNotes.MoveNext
      Loop
      End If %>
  </td>
</tr>
<% If status <> Cfg(cnnDB, "CloseStatus") Then %>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "EnterAdditionalNotes")%>:
  </td>
</tr>
<tr class="Body1">
	<td colspan="2" >
		<textarea rows="8" cols="80" name="notes"  <% = strTextDisable %>></textarea><br />
	</td>
</tr>
<% End If %>
<tr class="Head2">
	<td colspan="2" >
		<%=lang(cnnDB, "Solution")%>:
  </td>
</tr>
<tr class="Body1">
  <td colspan="2">
    <textarea rows="8" cols="80" name="solution"  <% = strTextDisable %>><% = solution %></textarea>
  </td>
</tr>
</table>
<% If status <> Cfg(cnnDB, "CloseStatus")Then
     Response.Write("<tr class=""Head2"" align=""center""><td colspan=""2"">") %>
      <input type="submit" value="<%=lang(cnnDB, "SaveProblem")%>" name="B1">
      </td></tr>
    <% End If %>
</div>
</form>
<%
	' close record sets
	rstCat.Close
	rstStat.Close
	rstPri.Close
	rstRes.Close

  rstProb.Close
  rstSol.Close
  rstNotes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>

</html>
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

  CVS File:      $RCSfile: new.asp,v $
  Commit Date:   $Date: 2004/10/20 14:58:39 $ (UTC)
  Committed by:  $Author: WalshKel $
  CVS Revision:  $Revision: 1.1 $
  Checkout Tag:  $Name(Version/Build)

  Purpose:  A form for users to enter new problems.
  -->

  <!-- 	#include file = "public.asp" -->

  <% 
    Dim cnnDB, sid, logon_fname,entered_by_id,quantity
    Set cnnDB = CreateCon
    quantity=1	
    sid = Cint(Session("lhd_sid"))
    entered_by_id=sid
    logon_fname = session("logon_fname")

    Session.timeout = 120
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "NewProblem")%></title>
    <link rel="stylesheet" type="text/css" href="default.css">
<script type="text/javascript" language="JavaScript" src="popcalendar.js"></script>

  </head>
  <body>

    <%
      ' Determine the username
'      dim username
'      username = Usr(cnnDB, sid, "username")
    %>

    <form action="postnew.asp" method="POST" id="Form1">
      <input type="hidden" name="entered_by_id" value="<% = entered_by_id %>">
      <div align="center">
        <table class="Wide">
          <tr>
    
          </tr>
          <tr class="Head1">
            <td colspan="1">
              <%=lang(cnnDB, "SubmitANewProblem")%>
            </td>
          </tr>
          <tr Class="Body1">
            <td valign="top">
              <div align="left">
                <table class="narrow" border="0">
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "InitiatedBy")%>:</b>
                    </td>
                    <td>
                      <% = logon_fname %>
                    </td>
                  </tr>

 <tr>
                    <td>
                      <b><%=lang(cnnDB, "Category")%>:</b>
                    </td>
                    <td>
                      <SELECT NAME="category">
                      <OPTION VALUE="0" SELECTED><%=lang(cnnDB, "SelectCategory")%></OPTION>
                      <%
                        ' Get list of categories to display
                        Dim rstCatList
                        Set rstCatList = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                        If Not rstCatList.EOF Then
                        Do While Not rstCatList.EOF
                      %>
                      <OPTION VALUE="<% = rstCatList("category_id")%>">
                      <% = rstCatList("cname") %></OPTION>

                      <% 		rstCatList.MoveNext
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
                    <td>
                      <b><%=lang(cnnDB, "Priority")%>:</b>
                    </td>
                    <td>
                      <SELECT NAME="priority">
                      <OPTION VALUE="0" SELECTED><%=lang(cnnDB, "SelectPriority")%></OPTION>

                      <%
                        ' Get list of categories to display
                        Dim rstPriList
                        Set rstPriList = SQLQuery(cnnDB, "SELECT * From priority ORDER BY pname ASC")
                        If Not rstPriList.EOF Then
                        Do While Not rstPriList.EOF
                      %>
                      <OPTION VALUE="<% = rstPriList("priority_id")%>">
                      <% = rstPriList("pname") %></OPTION>

                      <% 		rstPriList.MoveNext
                        Loop
                        End If
                      %>
                      </SELECT>
                    </td>
                  </tr>

			<tr>
                    <td>
                      <b><%=lang(cnnDB, "AssignedTo")%>:</b>
                    </td>
                    <td>
                      <SELECT NAME="assignedto">
                       <%
                        ' Get list of categories to display
                        Dim rstAssignList
                        Set rstAssignList = SQLQuery(cnnDB, "SELECT * From tblUsers ORDER BY fname ASC")
                        If Not rstAssignList.EOF Then
                        Do While Not rstAssignList.EOF
				If rstAssignList("fname") = logon_fname Then
                      %>
                      <OPTION VALUE="<% = rstAssignList("sid")%>" SELECTED>
                      <% = rstAssignList("fname") %></OPTION>
                      <% Else %>
                      <OPTION VALUE="<% = rstAssignList("sid")%>" >
                      <% = rstAssignList("fname") %></OPTION>
                      <% End If %>

                      <% 		rstAssignList.MoveNext
                        Loop
                        End If %>


                    </td>
                  </tr>
<tr>
<td>
<b><%=lang(cnnDB, "StartDate")%>:</b>
</td>
<td>
<!-- New Calendar -->
<input type=text name='startdate' size=20 maxlength=20 value='<%= MachineFormatDate(now())%>'>

<a href="javascript:showCalendar(Form1.startdate, Form1.startdate, 'yyyy-mm-dd',null,1,null,null)"><img src="calendaropen.gif"</img></a>
</td>
</tr>

<tr>
<td>
<b><%=lang(cnnDB, "DueDate")%>:</b>
</td>
<td>
<input type=text name='duedate' size=20 maxlength=20 value='<%= MachineFormatDate(now())%>'>

<a href="javascript:showCalendar(Form1.duedate, Form1.duedate, 'yyyy-mm-dd',null,1,null,null)"><img src="calendaropen.gif"</img></a>
</td>
</tr>

<!-- End New Calendar -->

                </table>
              </div>
            </td>
          <tr class="Head2">
            <td colspan="2">
              <%=lang(cnnDB, "ProblemInformation")%>:
            </td>
          </tr>
          <tr class="Body1">
            <td colspan="2">
              <b><%=lang(cnnDB, "Title")%>:</b><br>
              <input type="text" name="title" size="70">
              <p>
              <b><%=lang(cnnDB, "Description")%>:</b><br />
              <textarea rows="12" cols="80" name="description"></textarea>
            </td>
          </tr>
          <tr class="Head2">
            <td colspan="2">
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "SubmitProblem")%>" name="B1">&nbsp;<input type="reset" value="<%=lang(cnnDB, "ClearForm")%>" name="B2">
              </div>
            </td>
          </tr>
        </table>
      </div>
    </form>

    <%
      ' close record sets
      rstCatList.Close
   

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
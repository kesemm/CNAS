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
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: moduser.asp
  Date:     $Date: 2002/08/28 15:30:07 $
  Version:  $Revision: 1.52.4.2 $
  Purpose:  Form to modify user account info.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ModifyUser")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      Call CheckAdmin

      Dim modSid
      modSid = Cint(Request.Form("usersid"))

      Dim success
      success = FALSE

      If Request.Form("save") = 1 Then
        Dim uid, email, fname, phone, location, department, pager
        Dim firstname, lastname, ListOnInoutboard, phone_home, phone_mobile
        Dim jobfunction, userresume, statuscode, statustext, statusdate
        Dim usrLanguage, RepAccess, InoutAdmin
        
        uid = Left(Lcase(Trim(Request.Form("uid"))), 50)
        uid = Replace(uid, "'", "''")

        email = Left(Lcase(Trim(Request.Form("email"))), 50)
        email = Replace(email, "'", "''")
        firstname = Left(Trim(Replace(Request.Form("firstname"), "'", "''")), 25)
        lastname = Left(Trim(Replace(Request.Form("lastname"), "'", "''")), 24)
        fname = firstname & " " & lastname
        statusdate = SQLDate(Now, lhdAddSQLDelim)

        
        If Len(email) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Emailaddress") & lang(cnnDB, "is a required field") & ".")
        End If
        If Len(firstname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "FirstName") & lang(cnnDB, "is a required field") & ".")
        End IF
        If Len(lastname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "LastName") & lang(cnnDB, "is a required field") & ".")
        End IF

        Dim blnRepProbs
        blnRepProbs = False
        
        Dim sqlString, updRes

        Dim newpassword
        newpassword = Left(Trim(Request.Form("newpassword")), 50)
        newpassword = Replace(newpassword, "'", "''")
        If Len(newpassword) > 0 Then
          sqlString = "UPDATE tblUsers SET " & _
            "email = '" & email & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "statusdate = " & statusdate & ", " & _
            "[password] = '" & newpassword & "'"
        Else
          sqlString = "UPDATE tblUsers SET " & _
            "email = '" & email & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "statusdate = " & statusdate & " "
        End If

        sqlString = sqlString & " WHERE sid=" & modSid
        Set updRes = SQLQuery(cnnDB, sqlString)
        success = True
      End If

      If Request.Form("delete") = 1 Then
        Dim delRes, rstProblemUpd1, rstProblemUpd2, strUserId
        strUserId = usr(cnnDB, modSid, "uid")
        If Usr(cnnDB, modSid, "IsRep") = 0 Then
          Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
          Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
          Set rstProblemUpd2 = SQLQuery(cnnDB, "UPDATE problems SET rep=0 WHERE rep = " & modSid)
          success = True
        Else
          Dim rstRepCats
          Set rstRepCats = SQLQuery(cnnDB, "SELECT category_id FROM categories WHERE rep_id=" & modSid)
          If Not rstRepCats.EOF Then
            Call DisplayError(3, "Please reassign categories to a different support rep.")
          End If
          rstRepCats.Close
          Set rstRepProbs = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & modSid & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          If Not rstRepProbs.EOF Then
            blnRepProbs = True
          Else
            Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
            Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
            success = True
          End If
          rstRepProbs.Close
        End If
        'Code to delete user image file if In/Out Board is activated
        If cfg(cnnDB, "UseInoutBoard") = 1 Then
          If success = True Then
            Dim objFSO, strUserImage
            strUserImage = Server.MapPath("..\image\" & strUserId & ".jpg")
            Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
            If objFSO.FileExists(strUserImage) Then
              objFSO.DeleteFile strUserImage, False
            End If
            Set objFSO = Nothing
          End If
        End If
      Else
        Dim frm_email, frm_fname, frm_phone, frm_location, frm_department, frm_pager
        Dim frm_firstname, frm_lastname, frm_phone_home, frm_phone_mobile
        Dim frm_jobfunction, frm_userresume, frm_listoninoutboard, frm_usrLanguage
        Dim frm_statuscode, frm_statustext, frm_IsRep, frm_RepAccess, frm_InoutAdmin
        frm_email = Usr(cnnDB, modSid, "email")
        frm_firstname = Usr(cnnDB, modSid, "firstname")
        frm_lastname = Usr(cnnDB, modSid, "lastname")
      End If

    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "UpdateInformation")%>
          </td>
        </tr>
        <% If blnRepProbs Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "ErrorUpdatingAccount")%>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <%=lang(cnnDB, "ErrorUpdatingAccountText")%>.
            </td>
          </tr>
        <% ElseIf Request.Form("delete") = 1 Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "AccountDeleted")%>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <%=lang(cnnDB, "Theaccounthasbeenremoved")%>. 
              </div>
            </td>
          </tr>
        <% Else ' user form %>
          <% If success Then %>
            <tr class="Head2">
              <td>
                <div align="center">
                  <%=lang(cnnDB, "AccountUpdated")%>
                </div>
              </td>
            </tr>
          <% End If %>
          <tr class="Body1">
            <td>
              <form name="upduser" action="moduser.asp" method="POST">
                <input type="hidden" name="usersid" value="<% = modSid %>">
                <input type="hidden" name="save" value="1">
                <p>
                <table class="Normal">
                
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "Username")%>: </b>
                    </td>
                    <td>
                      <b><% = Usr(cnnDB, modSid, "uid") %></b>
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "FirstName")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="firstname" size="30" value="<% = frm_firstname %>">
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "LastName")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="lastname" size="30" value="<% = frm_lastname %>">
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "EmailAddress")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="email" size="30" value="<% = frm_email %>">
                    </td>
                  </tr>
                 <tr class="Head2">
                    <td colspan="2">
                      <%=lang(cnnDB, "Password")%>:
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "NewPassword")%>: </b>
                    </td>
                    <td>
                      <input type="password" name="newpassword" size="30">
                    </td>
                  </tr>
                </table>
                <p>
                <div align="center">
                  <p>
                  <input type="submit" value="<%=lang(cnnDB, "Save")%>">
                </div>
              </form>
              <hr width="80%">
              <p>
              <div align="center">
                <form name="deluser" action="moduser.asp" method="POST">
                  <input type="hidden" name="usersid" value="<% = modSid %>">
                  <input type="hidden" name="delete" value="1">
                  <input type="submit" value="<%=lang(cnnDB, "DeleteAccount")%>">
                </form>
              </div>
            </td>
          </tr>
        <% End If %>
      </table>
      <p>
      <a href="viewusers.asp"><%=lang(cnnDB, "ManageUsers")%></a><br>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>

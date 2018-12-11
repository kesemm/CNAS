<SCRIPT LANGUAGE=VBScript RUNAT=SERVER>

'  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
'  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.

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
  -->
'  Purpose:  This file holds the public procs and functions used
'            throughout the site.  Each page has an include for
'            this file.

' ################################
' PAGE START CODE
' ################################
' Declare Constants
CONST adOpenStatic = 3
CONST adOpenForwardOnly = 0

CONST lhdDateOnly = 0
CONST lhdDateTime = 1

CONST lhdAddSQLDelim = 0
CONST lhdNoSQLDelim = 1
' 
' Session Time Out Section
'
If Session("lhd_ext_uid") = ""   then
		
        Response.Redirect("Timeout.asp")
end if 


'#################################

' Cfg:
' Returns a configuration setting from tblConfig
Function Cfg(cnnDB, strSetting)
  Err.Clear
  Dim confRes
  Set confRes = SQLQuery(cnnDB, "SELECT " & strSetting & " From tblConfig")
  If confRes.EOF Then
    Cfg = vbNull
    Call DisplayError(3, strSetting & " " & lang(cnnDB, "isaninvalidsetting") & ".")
  Else
    Cfg = confRes(strSetting)
  End If
  confRes.Close
End Function


' Usr:
' Returns selected user information
Function Usr(cnnDB, sid, strColumn)
  Dim usrRes
  Set usrRes = SQLQuery(cnnDB, "SELECT " & strColumn & " FROM tblUsers WHERE sid=" & sid)
  If usrRes.EOF Then
    Call DisplayError(3, lang(cnnDB, "Usernotfound") & ".")
  Else
    Usr = UsrRes(0)
  End If
  UsrRes.Close
End Function



' Lang:
' Returns select Language string
Function Lang(cnnDB, variable)
  Dim intLangSetting, sid
  sid = GetSid
  If sid = 0 Then
    intLangSetting = Cfg(cnnDB, "DefaultLanguage")
  Else
    intLangSetting = Session("lhd_LanguageID")
    If Cint(intLangSetting) < 1 Then
      intLangSetting = Usr(cnnDB, sid, "[language]")
      If IsNull(intLangSetting) Or IsEmpty(intLangSetting) Or intLangSetting = 0 Then
        intLangSetting = Cfg(cnnDB, "DefaultLanguage")
      End If
      Session("lhd_LanguageID") = intLangSetting
    End If
  End If

  Dim varLangCache, intLangCacheCount, strAppVar
  strAppVar = "lhd_LangCache_" & intLangSetting
  varLangCache = Application(strAppVar)
  intLangCacheCount = Application(strAppVar & "_Count")
  If IsEmpty(varLangCache) Then
    Dim rstLangStr
    Set rstLangStr = SQLQuery(cnnDB, "SELECT variable, langText FROM tblLangStrings WHERE id=" & intLangSetting & " ORDER BY variable ASC")
    If Not rstLangStr.EOF Then
      varLangCache = rstLangStr.GetRows
    Else
      rstLangStr.Close
      Call DisplayError(3, "The language strings are missing.  Please run setup.asp to install them. (" & intLangSetting & ")")
    End If
    rstLangStr.Close
    Application(strAppVar) = varLangCache
    Application(strappVar & "_Count") = Cdbl(UBound(varLangCache, 2))
    intLangCacheCount = Cdbl(UBound(varLangCache, 2))
  End If
  Lang = variable


  Dim strValue
  strValue = ArrayFind(varLangCache, variable, 0, intLangCacheCount) ' array, search term, lower bound, upper bound
  If strValue = "_0" Then
    Lang = "@" & variable & "@"
  Else
    If Len(strValue) < 1 Then
      Lang = "!" & variable & "!"
    Else
      Lang = strValue
    End If
  End If
End Function


' ArrayFind:
' Finds a value in a two demensional, sorted array
' Returns "_0" if the key is not found
Function ArrayFind(varArray, strKey, intLBound, intUBound)
  Dim intMid
  ArrayFind = "_0"
  Do While intLBound <= intUBound
    intMid = (intUBound + intLBound) \ 2
    Select Case StrComp(UCase(strKey), UCase(varArray(0, intMid)), 1)
      Case 0  ' Key Found
        ArrayFind = varArray(1, intMid)
        Exit Function
      Case -1 ' Key less than mid
        intUBound = intMid - 1
      Case 1  ' Key greater than mid
        intLBound = intMid + 1
    End Select
  Loop
End Function


' ClearLangCache:
' Removes all language strings from the lang cache application variables
' The cache application variables acre called lhd_LangCache_<lang id>
Sub ClearLangCache(cnnDB)
On Error Goto 0
  Dim strAppVar, rstLangID
  Set rstLangID = SQLQuery(cnnDB, "SELECT id FROM tblLanguage")
  While Not rstLangID.EOF 
    strAppVar = "lhd_LangCache_" & rstLangID("id")
    Application(strAppVar) = Empty
    rstLangID.MoveNext
  Wend
  rstLangID.Close
  Set rstLangID = Nothing
End Sub


' CreateCon:
' Returns a ADO Connection object
Function CreateCon
	Dim strConn, cnnDB
	' Check for usage of SQL securing or integrated security and
	' use the correct connection string

	' Connection strings with DRIVER are ODBC,
	' those with PROVIDER are OLE DB connections
	Select Case Application("DBType")
		Case 1	' SQL Sec
			strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
			  ";Initial Catalog=" & Application("SQLDBase") & _
			  ";uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")

		Case 2	' SQL Integrated Sec
			strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
			  ";Initial Catalog=" & Application("SQLDBase") & ";Integrated Security=SSPI"

		Case 3	' Access
			strConn = "Provider=Microsoft.Jet.OLEDB.4.0" & _
			  ";Data Source=" & Application("AccessPath")
		Case 4	' DSN
			strConn = "DSN=" & Application("DSN_Name")

	End Select

	' Keep Errors from occuring
	If Not Application("Debug") Then
		On Error Resume Next
	End If

	' Create and open the database connection and save it as a
	' session variable
	Set cnnDB = Server.CreateObject("ADODB.Connection")
	cnnDB.Open(strConn)

	Set CreateCon = cnnDB

	If Err.Number <> 0 Then
		Call TrapError(Err.Number, Err.Description, Err.Source)
	End If
End Function


' GetSid:
' Returns the user's sid or 0
Function GetSid
	If Session("lhd_sid") > 0 Then
		GetSid = Session("lhd_sid")
	Else
		GetSid = 0
	End If
End Function


' DisplayError:
' Procedure for creating error pages.
Sub DisplayError(eType, component)

	' Create the web page
  If Not Application("Debug") Then
    Response.Clear
    Response.Write("<html><head><title>ERROR</title></head><body>")
  End If
  Response.Write "<p><center><table width=""200""><tr><td bgcolor=""red"" align=""center"">" & _
	  "<b>ERROR</b></tr></td><tr><td bgcolor=""#eeeeee"" align=""center"">"

'	Error Types:
'	1: Missing required field
'	2: SQL error
'	3: Generic Error, just display full component string

	Select Case eType
	Case 1
		Response.Write "<b>" & component & "</b> " & _
		  "&nbsp;" & lang(cnnDB, "isarequiredfield") & ".<p>" & _
		  "<i>" & lang(cnnDB, "PleasepresstheBACKbutton") & "</i></p>"
	Case 2
		Response.Write(lang(cnnDB, "ASQLqueryhasfailed") & ". ")
	Case 3
		Response.Write(component)
	End Select


	' Finish off the table and page
	Response.Write("</tr></td></table></center><p>&nbsp;</p></body></html>")

	' Stop processing the .asp file
	Response.End()
End Sub


' TrapError:
' Gracefully trap errors and print message
' ** Do not translate this routine
Sub TrapError (intNum, strDescription, strSource)

	Dim strInformation, strHexNum

	' Format the error number in Hex, 8 characters long
	strHexNum = Right(String(8, "0") & Hex(intNum), 8)

	Response.Write("<p><center><table width=""300"">")
	Response.Write("<tr><td bgcolor=""red"" align=""center"">")
	Response.Write("<B>Application Error</b></td></tr>")
	Response.Write("<tr><td bgcolor=""#EEEEEE"" align=""left"">")

	Response.Write("<b>Number: </b>" & intNum & " (0x" & strHexNum & ")<br />")
	Response.Write("<b>Source: </b>" & strSource & "<br />")
	Response.Write("<b>Description: </b>" & strDescription & "<hr />")

	' Print extra information
	Select Case Err.Number
		Case 3709	' no db connection
			strInformation = "The database connection could not be opened.  Please check your" & _
				" configuration and make sure the database is accessible."
		Case -2147217865 ' sql: bad table reference
			strInformation = "There is an error in the SQL query string probably referencing a" & _
				" table.  Read the description above and check the code."
		Case -2147217900 ' sql: bad syntax
			strInformation = "The SQL query string is using bad syntax.  Read the description above" & _
				" and check the code."
		Case -2147217904 ' sql: invalid column
			strInformation = "The SQL query references a invalid column or object.  Read the description above" & _
				" and check the code."
		Case Else
			strInformation = "No more information is available."
	End Select

	strInformation = strInformation & vbNewLine & _
		"<p>Contact your administrator or visit the Liberum Help Desk " & _
		"<a href=""http://www.liberum.org"">website</a>."
	Response.Write(strInformation)

	Response.End
End Sub



' DisplayHeader:
' Procedure to write HTML for the header.
Sub DisplayHeader(cnnDB, sid)
	Response.Write "<center><table width=""500""><tr><td valign=""top"" align=""left"">" & _
	  "<font size=""-1"">"

	'Display user's information
		Dim hdrNumProblems, hdrNumStr, hdrNumRes
		hdrNumStr = "SELECT count(*) AS total FROM problems WHERE assigned_to=" & sid & _
			"AND status<>" & Cfg(cnnDB, "CloseStatus")
		Set hdrNumRes = SQLQuery(cnnDB, hdrNumStr)
		hdrNumProblems = hdrNumRes("total")

		Response.Write "<b>" & lang(cnnDB, "UserName") & ":</b> " & Usr(cnnDB, sid, "fname") & _
		  "<br /><b>" & lang(cnnDB, "Problems") & ":</b> <a href=""view.asp"">" & hdrNumProblems & "</a>"
		Dim hdrProbStr, hdrProbRes
		hdrProbStr = "SELECT TOP 1 id, title FROM problems WHERE assigned_to='" & _
			Usr(cnnDB, sid, "sid") & "' ORDER BY id DESC"
		Set hdrProbRes = SQLQuery(cnnDB, hdrProbStr)
		If Not hdrProbRes.EOF Then
			Response.Write "<br /><b>" & lang(cnnDB, "MostRecent") & ":</b> <a href=""details.asp?id=" & hdrProbRes("id") & _
				""">" & hdrProbRes("title") & "</a>"
		End If
	Response.Write("</font></td><td valign=""top"" align=""right"">")

	' Display extra information (login type, admin URL)
	Response.Write("<font size=""-1"">")
	
	Response.Write("</font></td></tr></table></center>")
End Sub


' DisplayFooter:
' Procedure to write the HTML for a footer.  Use at the
' bottom of all pages.
Sub DisplayFooter(cnnDB, sid)
  Dim userChkRes
	Set userChkRes = SQLQuery (cnnDB, "SELECT username FROM tblUsers WHERE sid=" & sid)
	If (Not userChkRes.EOF) AND (sid <> 0) Then
    Response.Write("<p><div align=""center"">")
Response.Write "<a href=""default.asp"">" & "Home Page" & "</a> | <a href=""logoff.asp"">" & "LogOff" & "</a>"

'          Response.Write "<a href=""" & Cfg(cnnDB, "BaseURL") & "/rep"">" & lang(cnnDB, "RepMenu") & "</a> | " & _
'        "<a href=""" & Cfg(cnnDB, "BaseURL") & "/logoff.asp"">" & lang(cnnDB, "LogOff") & "</a>" & _
'		  "</div></p>"
	userChkRes.Close
End If
End Sub


' SQLQuery:
' Takes an input string which is the SQL query statment and
' returns the results of the query as a recordset if any
' results are returned.
Function SQLQuery(cnnDB, queryStr)

	If Application("Debug") Then
		If InStr(Lcase(queryStr), "config")=0 Then
			Response.Write("<p><b>SQL Query: </b>" & queryStr & "</p><p>")
		End If
	Else
		On Error Resume Next
	End If
  
  Dim adoRec
	Set adoRec = Server.CreateObject("ADODB.Recordset")

	adoRec.ActiveConnection = cnnDB

	adoRec.Open(queryStr)

	Set SQLQuery = adoRec

	If Err.Number <> 0 Then
		Call TrapError(Err.Number, Err.Description, Err.Source)
	End If

End Function


' GetUnique:
' Finds a unique key for a database.  The database name
' is given (dbname) and the lookup done in a table called
' db_keys.  This proc should be atomic, but isn't.  May cause
' some problems on a busy site.
Function GetUnique(cnnDB, dbname)

	Dim queryStr, key, adoRec

	queryStr = _
	"SELECT " & dbname & " FROM db_keys"

	Set adoRec = SQLQuery(cnnDB, queryStr)

	' Get the key from the results and return it
	key = Cint(adoRec(dbname))

	adoRec.Close

	GetUnique = key

	' Increment the key and update the database for next time.
	queryStr = _
	"UPDATE db_keys SET " & dbname & "=" & (key+1)

	Set adoRec = SQLQuery(cnnDB, queryStr)

End Function


' CheckUser:
' Checks to see if user is logged on
Sub CheckUser(cnnDB, sid)
  Dim userchkRes
	Set userchkRes = SQLQuery(cnnDB, "SELECT username FROM tblUsers WHERE sid=" & sid)
	If (userchkRes.EOF) OR (sid = 0) Then
		Dim reAddr

		reAddr = Cfg(cnnDB, "BaseURL") & "/logon.asp?URL=" & _
			Request.ServerVariables("PATH_INFO")
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			reAddr = reAddr & _
				"?" & Request.ServerVariables("QUERY_STRING")
		End If
		userchkRes.close
		cnnDB.Close
		Response.Clear
		Response.Redirect reAddr
	End If
	userchkRes.close
End Sub


'
' CheckRep:
' Checks to see if IsRep is TRUE, If not, returns
' an error.  Used for permission check on pages.
Sub CheckRep(cnnDB, sid)
  Dim userchkRes
	Set userchkRes = SQLQuery(cnnDB, "SELECT username FROM tblUsers WHERE sid=" & sid)
	If (userchkRes.EOF) OR (sid = 0) Then
		Dim reAddr

		reAddr = Cfg(cnnDB, "BaseURL") & "/logon.asp?URL=" & _
			Request.ServerVariables("PATH_INFO")
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			reAddr = reAddr & _
				"?" & Request.ServerVariables("QUERY_STRING")
		End If
		userchkRes.Close
		cnnDB.Close
		Response.Clear
		Response.Redirect reAddr
	End If
	userchkRes.Close

	If Usr(cnnDB, sid, "IsRep") <> 1 Then
		Call DisplayError(3, "Access denied.  You do not have permission to view this page.")
	End If
End Sub

' CheckKB:
' Checks to see if user has permissions to the KB
' EnableKB 0-Disable, 1-Rep, 2-User/Rep, 3-AnyOne
Sub CheckKB(cnnDB, sid)
  Dim blnOK
  blnOK = False
  Select Case Cfg(cnnDB, "EnableKB")
    Case 0
      blnOK = False
    Case 1
      Dim rstRep
      Set rstRep = SQLQuery(cnnDB, "SELECT IsRep FROM tblUsers WHERE sid=" & sid)
      If Cint(rstRep("IsRep")) = 1 Then
        blnOK = True
      Else
        blnOK = False
      End If
      rstRep.Close
    Case 2
      If sid > 0 Then
        blnOK = True
      Else
        blnOK = False
      End If
    Case 3
      blnOK = True
    Case Else
      blnOK = False
  End Select
	If Not blnOK Then
		cnnDB.Close
    Call DisplayError(3, lang(cnnDB, "Accessdenied") & " " & lang(cnnDB, "NoPermission")  & ".")
	End If
End Sub


' CheckAdmin:
' Same as CheckRep, except looks for admin privs
Sub CheckAdmin
	If Not Session("lhd_IsAdmin") Then
    Call DisplayError(3, lang(cnnDB, "Accessdenied") & " " & lang(cnnDB, "NoPermission") & ".")
	End If
End Sub


' SendMail:
' Sends mail using the supported system set in global.asa
Sub SendMail (strToAddr, strCCAddr,strSubject, strBody, cnnDB)

Dim Mail

If Not Application("Debug") Then
	On Error Resume Next
End If

Select Case Cfg(cnnDB, "EmailType")
	Case 0	'No Email

	Case 1	'CDONTS
		Set Mail = Server.CreateObject("CDONTS.NewMail")
		Mail.BodyFormat = 1	' Text Only, 0 for HTML
		Mail.Subject = strSubject
		Mail.To = strToAddr
		Mail.Body = strBody
		Mail.Send(Cfg(cnnDB, "HDName") & "<" & Cfg(cnnDB, "HDReply") & ">")
		Set Mail = Nothing

	Case 2	'Jmail
		On Error Resume Next ' Use Jmail logging
		Set Mail = Server.CreateObject("Jmail.Message")
		Mail.Logging = True
		Mail.From = Cfg(cnnDB, "HDReply")
		Mail.FromName = Cfg(cnnDB, "HDName")
		Mail.AddRecipient strToAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		If Not Mail.Send(Cfg(cnnDB, "SMTPServer")) Then
			If Application("Debug") Then
				Call DisplayError(3, Mail.Log)
			End If
		End If
		Set Mail = Nothing

	Case 3	'ASPEmail
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.Host = Cfg(cnnDB, "SMTPServer")
		Mail.From = Cfg(cnnDB, "HDReply")
		Mail.FromName = Cfg(cnnDB, "HDName")
		Mail.AddAddress strToAddr
If strToAddr<>strCCAddr Then Mail.AddCC strCCAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		Mail.Send
		Set Mail = Nothing

  Case 4  ' ASPMail
    Set Mail = Server.CreateObject("SMTPsvg.Mailer")
    Mail.FromName = Cfg(cnnDB, "HDName")
    Mail.FromAddress = Cfg(cnnDB, "HDReply")
    Mail.RemoteHost = Cfg(cnnDB, "SMTPServer")
    Mail.AddRecipient "Help Desk User", strToAddr
    Mail.Subject = strSubject
    Mail.BodyText = strBody
    Mail.SendMail
   
End Select
If Err.Number <> 0 Then
	Call TrapError(Err.Number, Err.Description, Err.Source)
End If

End Sub


' Message:
' Parses and sends an email
Sub eMessage (cnnDB, eType, id, strToAddr,strCCAddr)
  Dim emailRes
	Set emailRes = SQLQuery(cnnDB, "Select subject, body FROM tblEmailMsg WHERE type='" & eType & "'")
	If emailRes.EOF Then
		Call DisplayError(3, lang(cnnDB, "Nomessageoftype") & " " & eType & " " & lang(cnnDB, "wasfoundinthedatabase") & ".")
	End If
	Dim subject, body
	subject = emailRes("subject")
	body = emailRes("body")

	emailRes.Close
	Dim queryStr

	queryStr = _
		"SELECT p.id, p.entered_by_id, p.due_date,p.start_date, p.status, s.sname, " & _
		"p.close_date, pri.pname, c.cname, p.assigned_to, p.title, p.solution, p.description " & _
		"FROM ((problems AS p " & _
		"INNER JOIN status AS s ON p.status = s.status_id) " & _
		"INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
		"INNER JOIN categories AS c ON p.category = c.category_id " & _
		"WHERE p.id=" & id

  Dim probRes, repRes, userRes, notesRes, notes,rstAssignedTo,rstUpdatedBy
	Set probRes = SQLQuery(cnnDB, queryStr)
'	Set repRes = SQLQuery(cnnDB, "SELECT uid, email, fname FROM tblUsers WHERE sid=" & probRes("assigned_to"))
'	Set userRes = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE uid='" & probRes("uid") & "'")
	Set rstAssignedTo = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE sid=" & probRes("assigned_to"))
	Set rstUpdatedBy = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE sid=" & sid)

  Set notesRes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")

  If probRes.EOF Then
		cnnDB.Close
		Call DisplayError(3, lang(cnnDB, "Problem") & " " & id & " " & lang(cnnDB, "doesnotexist") & ". " & lang(cnnDB, "Cannotsendmail") & ".")
	End If

  If Not notesRes.EOF Then
    Do While Not notesRes.EOF
      If Len(notes) > 0 Then
        notes = notes & vbNewLine
      End If
      notes = notes & "[" & notesRes("addDate") & " - " & Usr(cnnDB,notesRes("updated_by_id"),"fname") & "]"
      notes = notes & vbNewLine
      notes = notes & notesRes("note") & vbNewLine
      notesRes.MoveNext
    Loop
  Else
    notes = " "
  End If
  notesRes.Close

  On Error Resume Next
  body = Replace(body, "[problemid]", probRes("id") & vbCRLF)
	body = Replace(body, "[description]", probRes("description")& vbCRLF)
	body = Replace(body, "[status]", probRes("sname")& vbCRLF)
	body = Replace(body, "[priority]", probRes("pname")& vbCRLF)
	body = Replace(body, "[startdate]", FormatDateTime(probRes("start_date"), 1) & vbCRLF)
	body = Replace(body, "[close_date]", FormatDateTime(probRes("close_date"), 1)& vbCRLF)
	body = Replace(body, "[due_date]", FormatDateTime(probRes("due_date"), 1)& vbCRLF)
	body = Replace(body, "[category]", probRes("cname")& vbCRLF)
	body = Replace(body, "[solution]", probRes("solution")& vbCRLF)
'	body = Replace(body, "[uid]", probRes("uid")& vbCRLF)
	body = Replace(body, "[ufname]", rstUpdatedBy("fname"))
'	body = Replace(body, "[rid]", repRes("uid")& vbCRLF)
'	body = Replace(body, "[rfname]", repRes("fname")& vbCRLF)
'	body = Replace(body, "[remail]", repRes("email")& vbCRLF)
	body = Replace(body, "[assigned_to]", rstAssignedTo("fname"))
      body = Replace(body, "[notes]", notes& vbCRLF)

	subject = Replace(subject, "[problemid]", probRes("id"))
	subject = Replace(subject, "[title]", probRes("title"))
'	subject = Replace(subject, "[description]", probRes("description"))
'	subject = Replace(subject, "[status]", probRes("sname"))
'	subject = Replace(subject, "[priority]", probRes("pname"))
'	subject = Replace(subject, "[startdate]", FormatDateTime(probRes("start_date"), 1)& " at " & FormatDateTime(probRes("start_date"), 4)& " hrs")
'	subject = Replace(subject, "[close_date]", FormatDateTime(probRes("close_date"), 1))
'	subject = Replace(subject, "[category]", probRes("cname"))
'	subject = Replace(subject, "[solution]", probRes("solution"))
'	subject = Replace(subject, "[baseurl]", Cfg(cnnDB, "BaseURL"))
'	subject = Replace(subject, "[uid]", probRes("uid"))
'	subject = Replace(subject, "[ufname]", userRes("fname"))
'	subject = Replace(subject, "[uemail]", probRes("uemail"))
'	subject = Replace(subject, "[rid]", repRes("uid"))
'	subject = Replace(subject, "[rfname]", repbRes("fname"))
'	subject = Replace(subject, "[remail]", repRes("email1"))
'	subject = Replace(subject, "[uurl]", Cfg(cnnDB, "BaseURL") & "/user/view.asp?id=" & id)
'	subject = Replace(subject, "[rurl]", Cfg(cnnDB, "BaseURL") & "/rep/view.asp?id=" & id)

  Err.Clear
  On Error GoTo 0
	Call SendMail (strToAddr, strCCAddr,Subject, Body, cnnDB)
End Sub

' FixDay:
' Returns a day not greater than the last day of the month
Function FixDay (intMonth, intDay, intYear)
  FixDay = intDay
  If (intMonth=4) Or (intMonth=6) Or (intMonth=9) Or (intMonth=11) Then
    If intDay > 30 Then
      FixDay = 30
    End If
  End If
  If (intMonth=2) and (intDay>28) Then
    If (intYear Mod 4 = 0) Then
      FixDay=29
    Else
      FixDay=28
    End If
  End If
End Function
         

' Function to making it easy to make dropdown lists from database
' cnnDB = a open ADO Connection object
' dropdownlistname = text name of the form field
' keyfieldname = text fieldname with value returned when the form is executed
' selectedkey = value related to keyfieldname used to show initial selected value
' valuefieldname = text name of the field shown in the dropdownlist
' tablename = text name of the table or view containing the fields
' criteria = text with the criteria for selecting data ex. 'id > 1 ' (can be empty)
' sortorder = text to enable sorting data ex. ' id ASC ' can be empty

function dropdownlist(cnnDB, dropdownlistname, keyfieldname, selectedkey, _
														 valuefieldname, tablename, criteria, sortorder)

	dim tempstr, dropdownRes, sqlStr

	' build and execute SQL query
	
	sqlStr = "SELECT " & keyfieldname & ", " & valuefieldname & " FROM " & tablename
	if criteria <> "" then
		sqlStr = sqlStr & " WHERE " & criteria
	end if
	if  sortorder <> "" then
		sqlStr = sqlStr & " ORDER BY " & sortorder
	end if
  Set dropdownRes = SQLQuery(cnnDB, sqlStr)

	tempstr = "<SELECT NAME=""" & dropdownlistname & """>"

  If Not dropdownRes.EOF Then
    Do While Not dropdownRes.EOF
      tempstr = tempstr & "<OPTION VALUE=""" & dropdownRes(keyfieldname) & """"
      if dropdownRes(keyfieldname)=selectedkey then
	      tempstr = tempstr & " SELECTED "
	    End If  
      tempstr = tempstr & ">" & dropdownRes(valuefieldname) & "</OPTION>"
      dropdownRes.MoveNext
    Loop
  End If

	tempstr = tempstr & "</SELECT>"
  
	dropdownlist = tempstr

End Function

         
' SQLDate
' Returns a correctly formated date for a SQL statement including
' the correct delimitations.
' dtDate = Date format recognized by CDate
Function SQLDate (dtDate, intUseDelim)
  If Len(Trim(dtDate)) > 0 Then
    'dtDate = CDate(dtDate)
    SQLDate = Year(dtDate) & "-" & Month(dtDate) & "-" & Day(dtDate)
    SQLDate = SQLDate & " " & Hour(dtDate) & ":" & Minute(dtDate) & ":" & Second(dtDate)
    Dim strDeLim
    If Application("DBType") = 1 Or Application("DBType") = 2 Then
      strDeLim = "'"
    Else
      strDeLim = "#"
    End If
    If intUseDelim = lhdAddSQLDelim Then
      SQLDate = strDelim & SQLDate & strDeLim
    End If
  Else
    SQLDate = ""
  End If
End Function
' DisplayDate
' Formats the date and time for the locale
' dtDate = Date format recognized by CDate
' intFormat = 0 - Date Only
'           = 1 - Date and Time
Function DisplayDate(dtDate, intFormat)
  If Len(Trim(dtDate)) > 0 Then
    dtDate = CDate(dtDate)

    Dim strLocalDate, strDefaultFormat
    strDefaultFormat = "mm/dd/yyyy"


    ' ###### CHANGE THIS STRING TO MATCH LOCAL DATE FORMAT ######
    strLocalDate = "mm/dd/yyyy"
    ' ###########################################################


    If Len(strLocalDate) = 0 Then
      strLocalDate = strDefaultFormat
    End If
    strLocalDate = LCase(strLocalDate)
    strLocalDate = Replace(strLocalDate, "yyyy", Year(dtDate))
    strLocalDate = Replace(strLocalDate, "yy", Right(Year(dtDate), 2))
    strLocalDate = Replace(strLocalDate, "mm", Month(dtDate))
    strLocalDate = Replace(strLocalDate, "dd", Day(dtDate))
    If intFormat = lhdDateTime Then
      strLocalDate = strLocalDate & " " & FormatDateTime(dtDate, vbLongTime)
    End If
    DisplayDate = strLocalDate
  Else
    DisplayDate = ""
  End If
End Function

Function MachineFormatDate(dtDate)
    dtDate = CDate(dtDate)

    Dim strLocalDate, strDefaultFormat
    strDefaultFormat = "yyyy-mm-dd"


    ' ###### CHANGE THIS STRING TO MATCH LOCAL DATE FORMAT ######
    strLocalDate = "yyyy-mm-dd"
    ' ###########################################################


    If Len(strLocalDate) = 0 Then
      strLocalDate = strDefaultFormat
    End If
    strLocalDate = LCase(strLocalDate)
	strLocalDate = Replace(strLocalDate, "yyyy", Year(dtDate))
    strLocalDate = Replace(strLocalDate, "yy", Right(Year(dtDate), 2))
    strLocalDate = Replace(strLocalDate, "mm", Right("0" & Month(dtDate),2))
    strLocalDate = Replace(strLocalDate, "dd", Right("0" & Day(dtDate),2))
    
    MachineFormatDate = strLocalDate
End Function

</SCRIPT>
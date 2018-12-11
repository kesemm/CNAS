<SCRIPT LANGUAGE = "VBScript" RUNAT="Server">

'  This is file is based on, but has undergone extensive modifications:
'  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
'  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.
'
'  CVS File:      $RCSfile:$
'  Commit Date:   $Date:$ (UTC)
'  Committed by:  $Author:$
'  CVS Revision:  $Revision:$
'  Checkout Tag:  $Name(Version/Build)
' --------
' SETTINGS.ASP
'
' Loads the appplication variables.
' --------
' SetAppVariables:
' The procedure runs when the application is started or the file is changed
' Primary ojbectives are to set variables/constants used throughout the
' application.
Sub SetAppVariables


	'========================================
	' Database Information

		' Database Type
		' 1 - SQL Server with SQL security (set SQLUser/SQLPass)
		' 2 - SQL Server with integrated security
		' 3 - Access Database (set AccessPath)
		' 4 - DSN (An ODBC DataSource) (set DSN_Name)

	Application("DBType") = 1
	'========================================

	'============ SQL SETTINGS ==============
	Application("SQLServer") = "10.10.10.10"	' Server name (don't put the leading \\)
	Application("SQLDBase") = "CNATracking"	' Database name
	Application("SQLUser") = "CNATracking"			' Account to log into the SQL server with
	Application("SQLPass") = "ticket"	' Password for account
	' =======================================

	'=========== ACCESS SETTINGS ============
	'Physical path to database file
	Application("AccessPath") = "C:\Inetpub\Databases\helpdesk2000.mdb"
	'========================================

	'============= DSN SETTINGS =============
	Application("DSN_Name") = "HelpDeskDSN"
	'========================================

	' Enable Debugging:
	' Set to true to view full MS errors and other debug information
	' printed.  (This will disable most On Error Resume Next statements.)
	Application("Debug") = False


End Sub

</SCRIPT>
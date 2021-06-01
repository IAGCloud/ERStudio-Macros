'TITLE:  Data Model Quality Checker
'DESCRIPTION: This macro checks that particular table and column properties have been set for all selected table objects
'             in the current model.
'             Properties checked are:
'					Tables have a valid tablename prefix
'					Column names follow BI naming standards
'					Table and column descriptions exist
'					Tables have a PK and PI defined
'					Views don't have a PK or PI defined
'					PK, PI, SI names meet ICW standards
'                   Complex views have some SQL defined
'					Database name set for a table or view
'					Data Area diagram name set for a table or view
'					At least one Security Role set for a table or view
'					Data type not set to default CHAR(999) - implies the data type hasn't actually been set
'					Column ICW_TX_PD exists in tables
'			  Errors and warnings are written to a log file
'AUTHOR: Jeremy James
'DATE:  30/11/2016
'
'CHANGE HISTORY
'Version  Date      Author     Description
'1.0      06/12/17  Jeremy J   Baseline
'1.1      09/12/17  Jeremy J   Added check for Reference Table attachment for REF tables
'                              Added check For User Role And/Or Job View Access
'1.2      13/12/17  Jeremy J   Fixed problem where valid work tablename prefix failing naming check
'1.3      14/12/17  Jeremy J   Relaxed checking for job role on views, and ICW_TX_PD on ICW GUI ref tables
'1.4      04/01/18  Jeremy J   Check for ";" at end of complex view postSQL
'                              Save report file location in clipboard
'                              Check ICW_TX_PD datatype correct
'							   Check ICW_TX_PD not on ICW GUI maintained table
'                              Check LUV_ALL, LUV_MAIN not in Additional Views attachment
'                              Check for extraneous Value field content for user Security Roles
'                              Check for use of a View modelling object
'1.5      11/01/18  Jeremy J   Check for database value with Generate SRC_ from LF_ attachment
'                              Check Duplicate Row Control is SET or MULTISET
'1.6      19/01/18  Jeremy J   Added SRC_ to list of work tables to ignore
'1.7      06/02/18  Jeremy J   Check database names in Additional ViewDB start with LUV or LJV
'                              Check column FORMAT for DATE, TIME, and TIMESTAMP datatypes in LF_ table
'1.8      08/02/18  Jeremy J   Get Naming Standards file path location from RefFilePath.txt, if found
'1.9      19/02/18  Jeremy J   Add warning if a PUBLIC security role is being used
'							   Fix problem with key/index errors being reported when option was de-selected
'1.10     22/02/18  Jeremy J   Check for empty Value fields in Attachments that require a value
'                              Fix problem with extra spaces in REPLACE VIEW statement in PostSQL field
'1.11	  28/02/18  Jeremy J   Add table and view name checking
'1.12     12/04/18  Jeremy J   Fix bug with prefix checking when tablename does not have a prefix
'                              Allow for REPLACE RECURSIVE VIEW in complex view SQL checking
'1.13	  18/04/18  Jeremy J   Add check for spaces and uppercase in table/view and column names
'							   Check column names in multi-column statistics attachment
'							   Select entities as they are processed to show progress in Explorer pane
'1.14     04/06/18  Jeremy J   Remove check on Value field in Security Properties
'1.15     14/06/18  Jeremy J   Exclude PARTITION from multi-column stats check
'							   Add check for Statistics on tables
'                              Downgrade no PK on table to a Warning
'							   Add (disabled) check for Data Governance Area
'1.16	  21/06/18  Jeremy J   Enable check for Data Governance Area
'1.17     06/07/18  Jeremy J   Check that FORMAT length is consistent with column length for CHAR and VARCHAR columns
'1.18     24/07/18  Jeremy J   Add exceptions handling for Naming Standards suffix datatype checking
'1.19     31/07/18  Jeremy J   Cater for Global Temporary (GT_) tables
'1.20     06/08/18  Jeremy J   Check for table and complex view with same name
'1.21     07/08/18  Jeremy J   Check for User role and Support role
'                              Check Data Governance Area attachment for LF_ tables
'1.22     09/10/18  Jeremy J   Change table/complex view with same name from Warning to Error
'1.23     29/11/18  Jeremy J   Check column Security Roles
'                              Check table and column name length <= than max length
'							   Check GUI reference table has a RefDataAdmin security role
'1.24     20/12/18  Jeremy J   Allow full word for first part of double-suffix column name
'1.25     09/01/19  Jeremy J   Fixed problem with Ref table with a RefAdmin role resulting in 'No User Role' warning
'1.26     31/01/19  Jeremy J   Reduced column mask datatype conflict to Warning
'                              Fixed spurious error generated for 'Hide' column masking value
'1.27     28/02/19  Jeremy J   Skip LF_ checks if 'Validated LF_' attachment used
'1.28     26/04/19  Jeremy J   No column security check on LAST_UPD_NM in REF_ tables
'1.29     15/07/19  Jeremy J   Fix problem detecting missing ";" at end of complex view DDL
'1.30     02/01/20  Jeremy J   Check data area abbreviation valid for B_ and A_ tables
'                              Check for Redemption Period on tables
'							   Add QuickSort
'1.31     11/01/20  Jeremy J   Add _TSZ and _PDZ suffixes to Naming Standards check
'1.32     23/01/20  Jeremy J   Add GT_ to list of work tables prefixes
'1.33     26/06/20  Jeremy J   Check NoUserView attachment and user security role consistency
'1.34     24/08/20  Jeremy J   Add SH_ to list of work table prefixes
'1.35     24/09/20  Jeremy J   Allow *Z_ as valid view prefix
'1.36     02/10/20  Jeremy J   Check for Archive attachment on view
'1.36a    14/02/21  Jeremy J   Modify Naming Standards to allow IDENTIFIER as a name part

'#Language "WWB-COM"

'Dim ER/Studio variables
Dim diag As Diagram
Dim mdl As Model
Dim dict As Dictionary
Dim submdl As SubModel
Dim selObjects As SelectedObjects
Dim so As SelectedObject
Dim entNames As Variant
Dim entNum As Variant
Dim entList(1024)
Dim entCount As Integer
Dim ent As Entity
Dim otherEnt As Entity
Dim attr As AttributeObj
Dim binding As BoundAttachment
Dim boundSecProp As BoundSecurityProperty
Dim viewNames As StringObjects
Dim viewName As StringObject

Dim EntType As String
Dim IsTable As Boolean
Dim IsView As Boolean
Dim IsWorkTable As Boolean
Dim IsRefTable As Boolean
Dim PKfound As Boolean
Dim PIfound As Boolean
Dim hasNoPI As Boolean
Dim ICWTXPDfound As Boolean
Dim hasDataArea As Boolean
Dim hasDataGovArea As Boolean
Dim RetPolID As Integer
Dim hasStats As Boolean
Dim hasRefMaint As Boolean
Dim IsICWGUI As Boolean
Dim guiCTS As Boolean
Dim guiUTS As Boolean
Dim guiUNM As Boolean
Dim hasUserRole As Boolean
Dim hasSupportRole As Boolean
Dim hasRefDataRole As Boolean
Dim hasJobAccess As Boolean
Dim viewDBlist As Variant
Dim viewDB As String
Dim ErrorCount As Integer
Dim WarningCount As Integer
Dim NamingCount As Integer
Dim CRLF As String
Dim optionsStr As String
Dim searchTxt As String
Dim wrkStr As String
Dim IsGood As Boolean
Dim selectAll As Integer
Dim RefFilePath As String
Dim attValue As Variant
Dim statsCols As Variant
Dim colName As String
Dim runTime As Date
Dim version As String
Dim i As Integer
Dim f As Decimal

Dim NSWords As String
Dim NSWordsLT As String
Dim NSPhrases As String
Dim NSSuffixFormats As String
Dim NSSuffixTypes As String
Dim NSSuffixLT As String
Dim NSTest As String
Dim alphaNum As String
Dim NSTblMax As Integer
Dim NSColMax As Integer

Type NameValuePair
	nm As String
	vlu As Variant
End Type

'Dim dataArea As NameValuePair
Dim dataAreas(1000) As NameValuePair
Dim dataAreaCount As Integer

Dim suffixException As NameValuePair
Dim NSSuffixExceptions(1000)
Dim NSExceptionsCount As Integer

Sub Main

	version = "v1.36a"

	Set diag = DiagramManager.ActiveDiagram
	Set mdl = diag.ActiveModel
	Set submdl = mdl.ActiveSubModel
	Set dict = diag.EnterpriseDataDictionaries.Item("ICW_EDD")

	TableCount = 0
	ViewCount = 0
	TotalError = 0
	TotalWarning = 0
	TotalNaming = 0
	NSTblMax = 50						'Table name max length
	NSColMax = 50						'Column name max length
	CRLF = Chr(13) & Chr(10)
	alphaNum = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	RetPolID = dict.AttachmentTypes.Item("Table metadata").Attachments.Item("Retention Policy").ID

	Debug.Clear

	currUser = DiagramManager.CurrentUser
	activeDir = DiagramManager.RepoActiveDirectory
	userID = Mid(currUser, InStr(currUser, "\") + 1)

	If currUser = "" Then
		MsgBox("You must be logged in to the Repository to use this macro", vbExclamation, "Not logged in")
		Exit Sub
	End If

	If CInt(Format(Now,"hh")) < 12 Then
		greeting = "Good morning " & userID & "."
	ElseIf CInt(Format(Now,"hh")) > 11 And CInt(Format(Now,"hh")) < 18 Then
		greeting = "Good afternoon " & userID & "."
	Else
		greeting = "Good evening " & userID & ". Working late!"
	End If

	'Get checking options
	Begin Dialog optionsDialog 615,345,"Data Model Quality Check " & version
			Text 20,15,560,15, greeting
			Text 20,35,560,15,"This macro will check all selected tables and views in the current sub-model for things"
			Text 20,50,560,15,"that are easily forgotten or incorrect, many of which can be easily fixed by using the"
			Text 20,65,560,15,"'Data Model Fixer-Upper' macro. If you aren't ready to check some aspects of your model"
			Text 20,80,560,15,"you can de-select them from the list below."
			Text 20,105,560,15,"Please make your selections below, and click 'OK' when ready."
			CheckBox 50,130,560,15, "Naming standards", .chkName
			CheckBox 50,150,560,15, "Descriptions", .chkDesc
			CheckBox 50,170,560,15, "PK and Indexes", .chkKeys
			CheckBox 50,190,560,15, "Complex View SQL", .chkView
			CheckBox 50,210,560,15, "Data Area name", .chkArea
			CheckBox 50,230,560,15, "Database name", .chkDB
			CheckBox 50,250,560,15, "Security role", .chkSec
			CheckBox 50,275,560,15, "Select all tables/views in sub-model ",.selectAll
			OKButton 150,305,120,21,.OK
			CancelButton 310,305,120,21,.Cancel
	End Dialog
	Dim options As optionsDialog

	'Use any saved options as default
	optionsFile = DiagramManager.RepoActiveDirectory & "\" & "DMQCoptions.txt"
	Open optionsFile For Binary As #1
	optionsStr = "00000000"
	Get #1, 1, optionsStr
	If Left(optionsStr,1) <> "0" And Left(optionsStr,1) <> "1" Then
		optionsStr = "11111111"
	End If
	options.chkName = Val(Mid(optionsStr,1,1))
	options.chkDesc = Val(Mid(optionsStr,2,1))
	options.chkKeys = Val(Mid(optionsStr,3,1))
	options.chkView = Val(Mid(optionsStr,4,1))
	options.chkArea = Val(Mid(optionsStr,5,1))
	options.chkDB = Val(Mid(optionsStr,6,1))
	options.chkSec = Val(Mid(optionsStr,7,1))
	options.selectAll = Val(Mid(optionsStr,8,1))

	If Dialog(options) = 0 Then
		Close #1
		Exit Sub
	End If

	Set selObjects = submdl.SelectedObjects
	If options.selectAll = 0 And selObjects.Count = 0 Then
		MsgBox("You need to select some tables or views first. Try again.", vbExclamation, "Nothing selected")
		Exit Sub
	End If

	'Build a sorted list of tables and views to check
	submdl.EntityNames(entNames, entNum)
	If options.selectAll = 1 Then						'Add all entities in the sub-model to the list
		For entCount = 0 To entNum-1
			entList(entCount) = mdl.Entities.Item(entNames(entCount)).TableName
		Next entCount
	Else												'Add only selected entities in the sub-model to the list
		entCount = 0
		For Each so In selObjects
			If so.Type = 1 Then
				entList(entCount) = mdl.Entities.Item(so.ID).TableName
				entCount += 1
			End If
		Next so
	End If
	dhQuickSort(entList,0,entCount-1)					'Sort the list

	runTime = Now
	DefaultName = "Quality Check Report " & Left(diag.FileName,InStr(diag.FileName,".dm1")-1) & " " & submdl.Name & " " & Format(runTime,"yyyymmddhhmmss") & ".txt"
	ReportFileName = GetFilePath(DefaultName,"Text file|*.txt",,"Report file name", 7)

	If ReportFileName = "" Then
		Exit Sub
	End If

	'Save the options they selected for next time
	Close #1
	Kill(optionsFile)

	Open optionsFile For Binary As #1
	optionsStr = ""
	optionsStr &= Trim(Str(options.chkName))
	optionsStr &= Trim(Str(options.chkDesc))
	optionsStr &= Trim(Str(options.chkKeys))
	optionsStr &= Trim(Str(options.chkView))
	optionsStr &= Trim(Str(options.chkArea))
	optionsStr &= Trim(Str(options.chkDB))
	optionsStr &= Trim(Str(options.chkSec))
	optionsStr &= Trim(Str(options.selectAll))
	Put #1, 1, optionsStr
	Close #1

	'Get any override path to Naming Standards file
	On Error GoTo Open_ReportFile
	RefFilePath = "\\usergroup1\im-warehouse\BI Design\ERStudio\macros\"
	optionsFile = DiagramManager.RepoActiveDirectory & "\" & "RefFilePath.txt"
	Open optionsFile For Input As #1
	Line Input #1, RefFilePath
	Close #1

	Open_ReportFile:
	On Error GoTo Error_unknown
	Open ReportFileName For Output As #1
	Print #1, "-- Data Model Quality Report " & version & " from sub-model '" & submdl.Name & "' in diagram '" & diag.FileName & "'"
	Print #1, "-- Created by " & userID & " at " & runTime & CRLF

	notChecked = "-- Not checking: "
	indent = "-- Not checking: "
	If options.chkName = 0 Then
		notChecked &= "Naming Standards" & CRLF & indent
	End If
	If options.chkDesc = 0 Then
		notChecked &= "Table and column descriptions" & CRLF & indent
	End If
	If options.chkKeys = 0 Then
		notChecked &= "Primary Keys, Primary Indexes, and Secondary Indexes" & CRLF & indent
	End If
	If options.chkView = 0 Then
		notChecked &= "SQL defined for complex Views" & CRLF & indent
	End If
	If options.chkArea = 0 Then
		notChecked &= "Data Area defined" & CRLF & indent
	End If
	If options.chkDB = 0 Then
		notChecked &= "Database defined" & CRLF & indent
	End If
	If options.chkSec = 0 Then
		notChecked &= "Security Roles defined" & CRLF & indent
	End If

	If notChecked = "-- Not checking: " Then
		Print #1, "-- Checking everything"
	Else
		Print #1, Left(notChecked, Len(notChecked)-Len(indent)-2)
	End If

	If entCount = entNum Then
		Print #1, "-- Checking all tables/views in sub-model"
	Else
		Print #1, "-- Only checking " & entCount & " out of " & entNum & " tables/views in sub-model"
	End If

	If options.chkName = 1 Then
		LoadNamingStandards()
	End If

	'If selObjects.Count > 0 Then
	'	For Each so In selObjects
	'		selObjects.Remove(so.Type,so.ID)
	'	Next
	'End If

	For entNum = 0 To entCount-1
			IsView = False
			IsTable = False
			IsWorkTable = False
			IsRefTable = False
			PKfound = False
			PIfound = False
			hasNoPI = False
			hasDataArea = False
			hasDataGovArea = False
			hasStats = False
			ICWTXPDfound = False
			hasUserRole = False
			hasSupportRole = False
			hasRefDataRole = False
			hasJobAccess = False
			hasRefMaint = False
			IsICWGUI = False
			guiCTS = False
			guiUTS = False
			guiUNM = False
			IsGood = True
			ErrorCount = 0
			WarningCount = 0
			NamingCount = 0

			'Set ent = mdl.Entities.Item(so.ID)
			Set ent = mdl.Entities.Item(entList(entNum))
			Debug.Print "Starting " & ent.TableName
			'selObjects.Add(1,ent.ID)

			'See if we have a View or a Table
			If Left(ent.TableName, 1) = "*" Then
				IsView = True
				EntType = "View "
				Set ViewCount = ViewCount + 1
			Else
				IsTable = True
				EntType = "Table "
				Set TableCount = TableCount + 1
			End If

			'See if we have a work table
			If IsTable Then
				p = InStr(ent.TableName, "_")
				If p > 1 Then
					If InStr("*H*N*LF*SRC*SW*SF*SX*SH*AW*AF*WRK*LND*GT*", "*" & Left(UCase(ent.TableName),p-1) & "*") > 0 Then
						IsWorkTable = True
					End If
				End If
			End If

			'Check table or view name length
			If Len(ent.TableName) > NSTblMax Then
				LogError("E", EntType, ent.TableName, "name exceeds " & NSTblMax & " character maximum length")
			End If

			'See if we have a table or view name conflict
			If IsTable Then
				Set otherEnt = mdl.Entities.Item("*" & ent.TableName)
				If otherEnt IsNot Nothing Then
					LogError("E", EntType, ent.TableName, "name same as complex view name " & otherEnt.TableName)
				End If
			Else
				Set otherEnt = mdl.Entities.Item(Mid(ent.TableName,2))
				If otherEnt IsNot Nothing Then
					LogError("E", EntType, ent.TableName, "name same as table name " & otherEnt.TableName)
				End If
			End If

			'Do Naming Standards checks for table name
			If options.chkName = 1 Then
				NSTest = ""
				If IsView Then									'Check for case, space, and invalid characters
					NSTest = checkCharCase(Mid(ent.TableName,2))
				Else
					NSTest = checkCharCase(ent.TableName)
				End If
				If NSTest <> "" Then LogError("N","",ent.TableName,NSTest)

				If IsWorkTable = False Then
					NSTest = checkEntityPrefix(ent.TableName)	'Check we have a valid table/view name prefix
					If NSTest <> "" Then LogError("N","",ent.TableName,NSTest)

					If InStr("A_B_",Left(ent.TableName,2)) > 0 Then		'Check data area abbreviation in table name
						NSTest = checkDataArea(ent.TableName)
						If NSTest <> "" Then LogError("N","",ent.TableName,NSTest)
					End If

					NSTest = checkEntityName(ent.TableName)		'Check we have a valid table/view name parts
					If NSTest <> "" Then LogError("N","",ent.TableName,NSTest)
				End If
			End If

			'Check we have a table/view definition
			If options.chkDesc = 1 And IsWorkTable = False Then
				If ent.Definition = "" Then LogError("E", EntType, ent.TableName, "has no description")
				If (Len(ent.Definition) > 0 And Len(ent.Definition) < 10) Or InStr(ent.Definition,"Created by RevEng Fixer-Upper") <> 0 Then
					LogError("E", EntType, ent.TableName, "has inadequate description")
				End If
			End If

			'Check we have a database for this table/view defined
			If options.chkDB = 1 Then
				If ent.Owner = "" Then
					LogError("E", EntType, ent.TableName, "has no database defined (in Owner field)")
				ElseIf IsTable And Left(ent.TableName,3) <> "GT_" And Left(ent.Owner, 4) <> "LDB_" Then
					LogError("W", EntType, ent.TableName, "database name does not start 'LDB_'. Please check")
				ElseIf IsView And ent.Owner <> "LUV_ALL" Then
					LogError("W", EntType, ent.TableName, "database name is not 'LUV_ALL'. Please check")
				End If
			End If

			'Check we have Duplicate Row Control defined
			If IsTable Then
				If ent.TeradataDuplicateRowControl <> "SET" And ent.TeradataDuplicateRowControl <> "MULTISET" Then
					LogError("E", EntType, ent.TableName, "Duplicate Row Control not defined")
				End If
			End If

			'Check the Index definitions
			If options.chkKeys = 1 Then
				CheckIndexes
				If  IsTable Then
					If Not PKfound And IsWorkTable = False Then
						LogError("W", EntType, ent.TableName, "has no Primary Key defined")
					End If
					If Not PIfound Then
						LogError("E", EntType, ent.TableName, "has no Primary Index")
					End If
				End If

				If IsView Then
					If Not PKfound Then
						LogError("W", EntType, ent.TableName, "has no Primary Key defined")
					End If
					If PIfound Then
						LogError("E", EntType, ent.TableName, "has a Primary Index defined!")
					End If
				End If
			End If

			'Check column names in multi-column statistics attachment
			For Each binding In ent.BoundAttachments
				If binding.Attachment.Name = "Multi-column statistics" Then
					attValue = Split(Trim(binding.ValueCurrent),Chr(13) & Chr(10))
					If Trim(binding.ValueCurrent) = "" Then
						LogError("W", EntType, ent.TableName, "multi-column statistics has no columns defined")
					Else
						hasStats = True
					End If
					For i = 0 To UBound(attValue)
						StatsText = Trim(attValue(i))
						If Left(StatsText,1) <> "(" Or Right(StatsText,1) <> ")" Then
							LogError("E", EntType, ent.TableName, "multi-column statistics has missing '(' or ')'")
							GoTo nextValue
						End If
						StatsText = Mid(StatsText,2,Len(StatsText)-2)
						statsCols = Split(StatsText,",")
						For x = 0 To UBound(statsCols)
							colName = Trim(statsCols(x))
							Set attr = ent.Attributes.Item(colName)
							If attr Is Nothing And colName <> "PARTITION" Then
								LogError("E", EntType, ent.TableName, "multi-column statistics " & colName & " not found")
							End If
						Next x
					nextValue:
					Next i
				End If
			Next binding

			'Check we have a Data Area and Retention Policy defined
			For Each binding In ent.BoundAttachments
				If binding.Attachment.Name = "Data Area" And binding.ValueCurrent <> "-- None --" Then
					hasDataArea = True
				End If
				If binding.Attachment.Name = "Data Governance Area" And binding.ValueCurrent <> "-- None --" Then
					hasDataGovArea = True
				End If
			Next
			If options.chkArea = 1 Then
				If IsWorkTable = False Or Left(ent.TableName,3) = "LF_" Then
					If hasDataArea = False Then LogError("E", EntType, ent.TableName, "has no Data Area set")
					If hasDataGovArea = False Then LogError("E", EntType, ent.TableName, "has no Data Governance Area set")
				End If
				If IsTable = True And IsWorkTable = False Then
					If ent.BoundAttachments.Item(RetPolID) Is Nothing Then LogError("E", EntType, ent.TableName, "has no Retention Policy set")
				End If
			End If

			'If it's a View, check we have some SQL defined for it
			If IsView And options.chkView = 1 Then
				If ent.PostSQL = "" Then
					LogError("E", EntType, ent.TableName, "has no SQL defined")
				Else
					searchTxt = "REPLACEVIEW" & ent.Owner & "." & Mid(ent.TableName,2)
					If InStr(Replace(ent.PostSQL," ",""), searchTxt) = 0 Then
						searchTxt = "REPLACERECURSIVEVIEW" & ent.Owner & "." & Mid(ent.TableName,2)
						If InStr(Replace(ent.PostSQL," ",""), searchTxt) = 0 Then
							LogError("E", EntType, ent.TableName, "SQL is inconsistent with view and/or database name")
						End If
					End If
					If InStr(ent.PostSQL, " LOCKING ") > 0 Then
						LogError("W", EntType, ent.TableName, "SQL should not contain 'LOCKING' modifier")
					End If
					If InStr(ent.PostSQL, Chr(10) & "LOCKING ") > 0 Then
						LogError("W", EntType, ent.TableName, "SQL should not contain 'LOCKING' modifier")
					End If
					ddlTxt = Trim(Replace(ent.PostSQL,CRLF," "))
					If Mid(ddlTxt,Len(ddltxt),1) <> ";" Then
						LogError("E", EntType, ent.TableName, "has no terminating ';' character")
					End If
				End If
			End If

			Set binding = ent.BoundAttachments.Item("Archive")
			If binding IsNot Nothing And IsView Then
				LogError("E", EntType, ent.TableName, "has Archive attachment. Please remove")
			End If

			If Mid(UCase(ent.TableName),2,4) = "_REF" Or Left(UCase(ent.TableName),4) = "REF_" Then
				IsRefTable = True
			End If

			If ent.BoundAttachments.Count > 0 Then
				For Each binding In ent.BoundAttachments
					If binding.Attachment.Name = "Reference table" Then
						If binding.ValueCurrent <> "-- None --" Then
							hasRefMaint = True
							If Left(binding.ValueCurrent,7) = "ICW GUI" Then
								IsICWGUI = True
							End If
						End If
					End If
					If binding.Attachment.Name = "Additional ViewDB" Then
						If InStr(binding.ValueCurrent,"LUV_ALL") > 0 Or InStr(binding.ValueCurrent,"LUV_MAIN") > 0 Then
							LogError("E", EntType, ent.TableName, "LUV_ALL or LUV_MAIN in Additional ViewDB list")
						End If
						If Trim(binding.ValueCurrent) = "" Then
							LogError("E", EntType, ent.TableName, "has no database defined in Additional ViewDB list")
						End If
						viewDBlist = Split(binding.ValueCurrent,",")
						For Each viewDB In viewDBlist
							viewDB = UCase(Trim(viewDB))
							If Left(viewDB,4) <> "LJV_" And Left(viewDB,4) <> "LUV_" Then
								LogError("W", EntType, ent.TableName, "Database not LUV_ or LJV_ in Additional ViewDB list")
							End If
						Next
					End If
					If binding.Attachment.Name = "Generate SRC_ from LF_" Then
						If Trim(binding.ValueCurrent) = "" Then
							LogError("E", EntType, ent.TableName, "has no database defined for SRC_ table")
						End If
					End If
					If binding.Attachment.Name = "View Sequence" Then
						If Trim(binding.ValueCurrent) = "" Then
							LogError("E", EntType, ent.TableName, "has no Value defined for View Sequence")
						End If
					End If
				Next
			End If
			If IsRefTable = True And hasRefMaint = False Then
				LogError("E", EntType, ent.TableName, "has no update method defined")
			End If

			'Check we have some Security Roles defined
			If options.chkSec = 1 Then
				If ent.BoundSecurityProperties.Count > 0 Then
					For Each boundSecProp In ent.BoundSecurityProperties
						Select Case boundSecProp.SecurityProperty.SecurityType.Name
						Case "Job View Access"
							hasJobAccess = True
						Case "BizinfoImSupport Roles"
							hasSupportRole = True
						Case "RefDataAdmin Roles"
							hasRefDataRole = True
						Case "Column Security Roles"
							LogError("E", "", "", "Security Property " & boundSecProp.SecurityProperty.Name & " invalid for a table or view")
						Case Else
							hasUserRole = True
						End Select
						If boundSecProp.SecurityProperty.Name = "PUBLIC" Then
							LogError("W", EntType, ent.TableName, "Security Property " & boundSecProp.SecurityProperty.Name & " has been discontinued")
						End If
						'If boundSecProp.SecurityProperty.Name <> "PUBLIC" And boundSecProp.ValueCurrent <> "" Then
						'	LogError("E", EntType, ent.TableName, "Security Property " & boundSecProp.SecurityProperty.Name & " has unexpected Value field")
						'End If
					Next boundSecProp
				End If
				If Left(ent.TableName,3) = "GT_" Then
					If hasJobAccess = True Or hasUserRole = True Or hasSupportRole = True Then
						LogError("W", EntType, ent.TableName, "has Security Roles defined. Please check.")
					End If
				Else
					If hasJobAccess = False And IsTable = True Then
						LogError("W", EntType, ent.TableName, "has no Job View Access defined")
					End If
					If hasSupportRole = False Then
						LogError("E", EntType, ent.TableName, "has no Support Role defined")
					End If
					If IsICWGUI = True And hasRefDataRole = False Then
						LogError("E", EntType, ent.TableName, "has no RefDataAdmin Role defined")
					End If
				End If

				'Check for NoUserView/user security role inconsistency
				If ent.BoundAttachments.Item("No User View") IsNot Nothing Then
					If hasUserRole = True Or hasRefDataRole = True Then
						LogError("E", EntType, ent.TableName, "has User Role defined with NoUserView attachment")
					End If
				Else
					If hasUserRole = False And hasRefDataRole = False And IsWorkTable = False Then
						LogError("E", EntType, ent.TableName, "has no User Role defined")
					End If
				End If

			End If

			'Start checking all the columns in this table/view
			For Each attr In ent.Attributes

				'Check column name length
				If Len(attr.ColumnName) > NSColMax Then
					LogError("E", "", attr.ColumnName, "name exceeds " & NSColMax & " character maximum length")
				End If

				'Check for case, space, and invalid characters
				If options.chkName = 1 Then
					NSTest = checkCharCase(attr.ColumnName)
					If NSTest <> "" Then LogError("N","",attr.ColumnName,NSTest)
				End If

				If options.chkName = 1 And IsWorkTable = False Then
					'Check column name parts
					NSTest = IsStandardName(attr.ColumnName)
					If NSTest <> "" Then
						LogError("N", "", attr.ColumnName, "has invalid abbreviation: "+NSTest)
					End If

					'Check suffix
					NSTest = checkSuffix(attr.ColumnName, attr.Datatype)
					If NSTest <> "" Then
						LogError("N", "", attr.ColumnName, NSTest)
					End If

					'Check column name is not reserved
					If IsReservedWord(attr.ColumnName) = -1 Then
						LogError("W", "column ", attr.ColumnName, "is a reserved word")
					End If

				End If

				'Check we have a column description
				If options.chkDesc = 1 And IsWorkTable = False And attr.Definition = "" Then
					LogError("E", "column ", attr.ColumnName, "has no description")
				End If

				'Check if it looks like we're just using the default column data type
				If attr.Datatype = "CHAR" And attr.DataLength = 999 Then
					LogError("W", "column ", attr.ColumnName, "has default datatype CHAR(999). Please check")
				End If

				'See if we've got ICW_TX_PD
				If attr.ColumnName = "ICW_TX_PD" Then
					ICWTXPDfound = True
					goodDefn = "TIMESTAMPNOT NULL2TruePERIOD(CURRENT_TIMESTAMP(2), UNTIL_CHANGED)"
					attrDefn = attr.Datatype & attr.NullOption & attr.DataLength & attr.TeradataPeriod & attr.DeclaredDefault
					If attrDefn <> goodDefn Then
						LogError("E", "column ", attr.ColumnName, "datatype definition incorrect or incomplete")
					End If
				End If

				'Check ICW GUI audit columns
				If IsICWGUI And attr.ColumnName = "CREATE_TS" Then
					guiCTS = True
					goodDefn = "TIMESTAMP0False"
					attrDefn = attr.Datatype & attr.DataLength & attr.TeradataPeriod
					If attrDefn <> goodDefn Then
						LogError("E", "column ", attr.ColumnName, "ICW GUI audit column definition incorrect or incomplete")
					End If
				End If
				If IsICWGUI And attr.ColumnName = "LAST_UPD_TS" Then
					guiUTS = True
					goodDefn = "TIMESTAMP0False"
					attrDefn = attr.Datatype & attr.DataLength & attr.TeradataPeriod
					If attrDefn <> goodDefn Then
						LogError("E", "column ", attr.ColumnName, "ICW GUI audit column definition incorrect or incomplete")
					End If
				End If
				If IsICWGUI And attr.ColumnName = "LAST_UPD_NM" Then
					guiUNM = True
					goodDefn = "CHAR10"
					attrDefn = attr.Datatype & attr.DataLength
					If attrDefn <> goodDefn Then
						LogError("E", "column ", attr.ColumnName, "ICW GUI audit column definition incorrect or incomplete")
					End If
				End If

				'Check we have FORMAT for DATE, TIME, TIMESTAMP datatypes in an LF_ table
				If Left(ent.TableName, 3) = "LF_" And ent.BoundAttachments.Item("Validated LF_") Is Nothing Then
					If InStr("*DATE*TIME*TIMESTAMP*",attr.Datatype) > 0 And attr.Format = "" Then
						LogError("E", "column ", attr.ColumnName, "has no FORMAT specifed")
					End If
				End If

				'Check we have a value specified for an Initial Value attachment
				If attr.BoundAttachments.Count> 0 Then
					For Each binding In attr.BoundAttachments
						If binding.Attachment.Name = "Initial value" Then
							If Trim(binding.ValueCurrent) = "" Then
								LogError("E", EntType, attr.ColumnName, "has no value defined for Initial value")
							End If
						End If
					Next
				End If

				If attr.TeradataCollectStatistics = True Then hasStats = True

				'Check any FORMAT is consistent with column width
				If attr.Datatype = "CHAR" Or attr.Datatype = "VARCHAR" Then
					x = InStr(attr.Format,"X(")
					If x > 0 Then
						y = InStr(x+2,attr.Format,")")
						flen = Val(Mid(attr.Format,x+2,y-x-2))
						'If attr.CharacterSet = "UNICODE" Then flen = flen*2
						If flen <> attr.DataLength Then
								LogError("E", "column ", attr.ColumnName, "FORMAT length does not match column length")
						End If
					End If
				End If

				'Check column security
				If options.chkSec = 1 Then
					If IsTable = True And attr.BoundAttachments.Item("Personal Data") IsNot Nothing And attr.BoundSecurityProperties.Count = 0 Then
						If attr.ColumnName <> "LAST_UPD_NM" And (Left(ent.TableName,4) <> "REF_" Or Left(ent.TableName,6) <> "B_REF_") Then
							LogError("W", "column ", attr.ColumnName, "is Personal Data but has no column security role")
						End If
					End If
					If attr.BoundSecurityProperties.Count > 0 Then
						For Each boundSecProp In attr.BoundSecurityProperties
							If boundSecProp.SecurityProperty.SecurityType.Name <> "Column Security Roles" Then
								LogError("E", "column ", attr.ColumnName, "Security Property " & boundSecProp.SecurityProperty.Name & " invalid for a column")
								GoTo nextSecProp
							End If
							If boundSecProp.ValueCurrent = "Hide" Then GoTo nextSecProp
							If attr.Datatype = "DATE" And Left(boundSecProp.ValueCurrent,6) = "DATE '" And Right(boundSecProp.ValueCurrent,1) = "'" Then
								If IsDate(Mid(boundSecProp.ValueCurrent,7,Len(boundSecProp.ValueCurrent)-7)) Then GoTo nextSecProp
							End If
							If (attr.Datatype = "CHAR" Or attr.Datatype = "VARCHAR") Then
								If boundSecProp.ValueCurrent <> "Stars" Then
									LogError("W", "column ", attr.ColumnName, "Security Property " & boundSecProp.SecurityProperty.Name & " has conflicting mask value '" & boundSecProp.ValueCurrent & "' for column datatype")
								End If
								GoTo nextSecProp
							End If
							If IsNumeric(boundSecProp.ValueCurrent) Then GoTo nextSecProp
							LogError("W", "column ", attr.ColumnName, "Security Property " & boundSecProp.SecurityProperty.Name & " has conflicting mask value '" & boundSecProp.ValueCurrent & "' for column datatype")
						nextSecProp:
						Next boundSecProp
					End If
				End If


			Next attr

			'Check we have an ICW_TX_PD column if this is a Base layer table
			If IsTable And Not IsICWGUI And Left(ent.TableName,2) = "B_" And Not ICWTXPDfound Then
				LogError("W", EntType, ent.TableName, "does not have column ICW_TX_PD. Please check")
			End If

			'Check we have right columns for an ICW GUI maintained table
			If IsICWGUI Then
				If ICWTXPDfound Then LogError("E", EntType, ent.TableName, "ICW GUI maintained table should not have column ICW_TX_PD")
				If guiCTS = False Then LogError("E", EntType, ent.TableName, "ICW GUI maintained table missing CREATE_TS")
				If guiUTS = False Then LogError("E", EntType, ent.TableName, "ICW GUI maintained table missing LAST_UPD_TS")
				If guiUNM = False Then LogError("E", EntType, ent.TableName, "ICW GUI maintained table missing LAST_UPD_NM")
			End If

			'Check we have Collect Stats for a table
			If IsTable And options.chkKeys = 1 And Not hasNoPI And Not hasStats Then
				LogError("W", EntType, ent.TableName, "has no statistics columns defined. Please check")
			End If

			TotalError += ErrorCount
			TotalWarning += WarningCount
			TotalNaming += NamingCount

			If IsGood = True Then
				Print #1, CRLF & EntType & ent.TableName & " passed checks"
			End If

	Next entNum

	'Check for use of View model objects
	Print #1, ""
	If options.selectAll = 1 Then
		Set viewNames = submdl.ViewNames
		If viewNames.Count > 0 Then
			For Each viewName In viewNames
				vw = viewName.StringValue
				If Left(vw,3) <> "FI_" And Left(vw,3) <> "FO_" Then
					Print #1, "  Error  : Unexpected View object " & vw & " found."
					ErrorCount += 1
				End If
			Next
		End If
	Else
		For Each so In selObjects
			If so.Type = 16 Then
				vw = mdl.Views.Item(so.ID).Name
				If Left(vw,3) <> "FI_" And Left(vw,3) <> "FO_" Then
					Print #1, "  Error  : Unexpected View object " & vw & " found."
					ErrorCount += 1
				End If
			End If
		Next
	End If

	Print #1, " "
	Print #1, TableCount & " tables found"
	Print #1, ViewCount & " views found"
	Print #1, TotalError & " errors"
	Print #1, TotalWarning & " warnings"
	Print #1, TotalNaming & " naming standards violations"
	Print #1, " "
	Print #1, "Explanation of all problem messages can be found at:"
	'Print #1, " "
	Print #1, "  https://baplc.sharepoint.com/sites/BusinessIntelligence/Team Wiki/ER Studio - Quality Checker errors and warnings.aspx"
	Print #1, " "
	Print #1, "-- End of Quality Check Report for " & submdl.Name & "."

	Close #1

	'Show results summary
	Begin Dialog ResultsDialog 360,180,"Data Model Quality Check Results"
			Text 20,15,180,15,TableCount & " tables checked"
			Text 20,35,180,15,ViewCount & " views checked"
			Text 20,55,180,15,TotalError+TotalNaming & " errors"
			Text 20,75,180,15,TotalWarning & " warnings"
			Text 20,110,180,15, "See logfile for further details"
			OKButton 130,140,90,21,.OK
	End Dialog
	Dim ShowResults As ResultsDialog

	Dialog ShowResults

	'If options.selectAll = 0 Then
	'	For Each so In selObjects
	'		selObjects.Remove(so.Type,so.ID)
	'	Next
	'End If

	Clipboard(ReportFileName)
	Exit Sub

	Error_unknown:
	MsgBox(Err.Description,vbCritical,"Error!")
	Exit Sub

End Sub


Sub CheckIndexes

	Dim ind As Index

	For Each ind In ent.Indexes
		If ind.IsPK Then
			PKfound = True
			If ind.Name <> "PK" Then
				LogError("E", EntType, ent.TableName, "has a bad Primary Key name '" & ind.Name & "'")
			End If
		End If
		If ind.TeradataPrimary Then
			If ind.IsPK Then
				LogError("E", EntType, ent.TableName, "must use separate Indexes for PK and PI")
			Else
				PIfound = True
				If ind.Name <> "PI" And IsWorkTable = False Then
					LogError("E", EntType, ent.TableName, "has a bad Primary Index name '" & ind.Name & "'")
				End If
			End If
		End If
		If Not ind.IsPK And Not ind.TeradataPrimary Then
			If Left(ind.Name, 3) <> "SI_" Then
				LogError("E", EntType, ent.TableName, "has a bad Secondary Index name '" & ind.Name & "'")
			End If
		End If
	Next

	If ent.BoundAttachments.Count > 0 Then
		For Each binding In ent.BoundAttachments
			If binding.Attachment.Name = "No Primary Index" Then
				hasNoPI = True
				PIfound = True
			End If
		Next
	End If


End Sub

Sub LogError(severity As String, objType As String, objName As String, msgText As String)

	oType = objType
	oName = objName

	If objType = "Table " Or objType = "View " Then
		oName = ""
		oType = LCase(Trim(objType))
	End If

	If objType = "column " Then
		oType = ""
		oName = objName
	End If

	If IsGood = True Then
		Print #1, CRLF & EntType & ent.TableName & " failed checks"
		IsGood = False
	End If

	If severity = "E" Then
		Print #1, "  Error  : " & Trim(oType & oName & " " & msgText)
		ErrorCount += 1
	End If

	If severity = "W" Then
		Print #1, "  Warning: " & Trim(oType & oName & " " & msgText)
		WarningCount += 1
	End If

	If severity = "N" Then
		Print #1, "  Naming : " & Trim(oType & oName & " " & msgText)
		NamingCount += 1
	End If

	If severity <> "E" And severity <> "W" And severity <> "N" Then
		Print #1, "  Uh?!?  : " & objType & " " & objName & " " & msgText
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	Case 6 ' Function key
	End Select

End Function

Function LoadNamingStandards() As Integer
	Dim i As Integer
	Dim fileName As String
	Dim inputRec As String
	Dim standardRec As Variant
	Dim abbrvTxt As String
	Dim longTxt As String
	Dim standardTyp As String
	Dim formatTyp As String
	Dim sep As String
	Dim suffixName As String
	Dim suffixFormat As String
	Dim suffix As String
	Dim dataArea As NameValuePair

	'fileName = "\\usergroup1\im-warehouse\BI Design\ERStudio\macros\NamingStandards.txt"
	fileName = RefFilePath & "NamingStandards.txt"
	Debug.Print fileName
	Open fileName For Input As #2

	sep = "*"

	NSWords = sep
	NSWordsLT = sep
	NSPhrases = sep
	NSSuffixFormats = sep
	NSSuffixTypes = sep

	While Not EOF(2)
	'While abbrvTxt <> "LAST"
		Line Input #2, inputRec
		standardRec = Split(inputRec, "|")
		abbrvTxt = standardRec(0)
		standardTyp = standardRec(1)
		longTxt = standardRec(2)
		formatTyp = standardRec(4)
		'Debug.Print standardTyp & " " & abbrvTxt & " " & longTxt & " " & formatTyp

		Select Case UCase(Trim(standardTyp))
		Case "WORD"
			NSWords = NSWords & Trim(abbrvTxt) & sep
			NSWordsLT = NSWordsLT & UCase(Trim(longTxt)) & sep
		Case "PHRASE"
			NSPhrases = NSPhrases & Trim(abbrvTxt) & sep
		Case "SUFFIX"
			suffixName = Trim(abbrvTxt)
			suffixFormat = Trim(formatTyp)
			suffix = suffixName & "=" & suffixFormat
			' If we haven't already seen this suffix and format add it to NSSuffixFormats
			' NSSuffixes will be similar to *CD=CHAR*CD=INTEGER*CD=SMALLINT*DESC=VARCHAR...
			If InStr(NSSuffixFormats,suffix) = 0 Then
				NSSuffixFormats = NSSuffixFormats & suffix & sep
			End If
			' If we haven't already seen this suffix and format add it to NSSuffixTypes
			' NSSuffixes will be similar to *CD=CHAR*CD=INTEGER*CD=SMALLINT*DESC=VARCHAR...
			If InStr(NSSuffixTypes,suffixName) = 0 Then
				NSSuffixTypes = NSSuffixTypes & suffixName & sep
			End If
		End Select
	Wend

	NSSuffixTypes &= "TSZ*PDZ*"

	Close #2

	'Load Naming Standards suffix exceptions
	fileName = RefFilePath & "NamingSuffixExceptions.txt"
	Debug.Print fileName
	Open fileName For Input As #2
	Line Input #2, inputRec
	NSExceptionsCount = 0

	While Not EOF(2)
		Line Input #2, inputRec
		standardRec = Split(inputRec, "|")
		suffixException.nm = standardRec(0)
		suffixException.vlu = standardRec(1)
		NSSuffixExceptions(NSExceptionsCount) = suffixException
		NSExceptionsCount += 1
	Wend

	Close #2

	'Load Naming Standards suffix words
	fileName = RefFilePath & "NamingSuffixLongText.txt"
	Debug.Print fileName
	Open fileName For Input As #2
	Line Input #2, inputRec
	NSSuffixLT = "*"

	While Not EOF(2)
		Line Input #2, inputRec
		NSSuffixLT &= inputRec & "*"
	Wend

	Close #2

	'Load Data Area abbreviations
	fileName = RefFilePath & "DataAreaAbbreviations.csv"
	Debug.Print fileName
	Open fileName For Input As #2
	Line Input #2, inputRec
	dataAreaCount = 0

	While Not EOF(2)
		Line Input #2, inputRec
		standardRec = Split(inputRec, ",")
		dataArea.nm = Trim(Replace(standardRec(1),"_",""))
		dataArea.vlu = Left(standardRec(2),1)
		dataAreas(dataAreaCount) = dataArea
		dataAreaCount += 1
	Wend

	Close #2

	Return 0
End Function

Function checkCharCase(strName As String)

	If InStr(strName," ") > 0 Then Return "contains space character(s)"
	If StrComp(UCase(strName),strName,0) <> 0 Then Return "is not in UPPERCASE"
	If InStr(alphaNum,Left(strName,1)) = 0 Then Return "starts with invalid character"
	If InStr(alphaNum,Right(strName,1)) = 0 Then Return "ends with invalid character"

	Return ""

End Function

Function IsStandardName(strName As String) As String

	'Split the column name into individual parts and convert to uppercase
	parts = Split(UCase(strName),"_")
	suffix = parts(UBound(parts))
	For i = 0 To (UBound(parts)-1)
		starPart = "*"+parts(i)+"*"
		If InStr(NSWords,starPart) = 0 Then
			If InStr(NSWordsLT,starPart) = 0 Or starPart = "*IDENTIFIER*" Then
				Return ""
			End If
			If InStr(NSPhrases,starPart) = 0 And (i <> UBound(parts)-1 Or InStr(NSSuffixLT,starpart) = 0) Then
				Return parts(i)
			End If
		End If
	Next i

	Return ""

End Function

Function checkEntityPrefix(strName As String)

	'Skip further checks for an archive view
	If Left(strName,3) = "*Z_" Then Return ""

	'Check for single character prefix for view name
	If Left(strName, 1) = "*" Then
		If InStr(strName,"_") = 3 Then
			Return "has invalid prefix for a view"
		Else
			Return ""
		End If
	End If

	'Check for valid prefix for table name
	p = InStr(strName, "_")
	If p >= 2 And p <= 4 Then
		If InStr("*A*B*AH*H*N*LF*SRC*SW*SF*SX*SH*AW*AF*WRK*LND*GT*", "*" & Left(strName,p-1) & "*") > 0 Then
			Return ""
		Else
			Return "has invalid tablename prefix"
		End If
	End If

	'Table name has no prefix
	Return ""

End Function

Function checkEntityName(strName As String)

Dim theName As String
Dim parts As Variant
Dim badName As Boolean
Dim msgTxt As String

	If Left(strName,1) = "*" Then
		theName = Mid(strName,2)
	Else
		theName = strName
	End If

	'NSTest = checkCharCase(theName)
	'If NSTest <> "" Then Return(NSTest)

	badParts = "*CATEGORY*DETAIL*"

	parts = Split(UCase(theName),"_")
	msgTxt = "has invalid name part "
	badName = False
	For i = 0 To UBound(parts)
		If InStr(badparts,"*" & parts(i) & "*") Then
			If badName = True Then
				msgTxt = Replace(msgTxt,"part ","parts ")
				msgTxt &= ", " & parts(i)
			Else
				msgTxt &= parts(i)
				badName = True
			End If
		End If
	Next i

	If badName = True Then
		Return msgTxt
	Else
		Return ""
	End If

End Function

Function checkDataArea(strName As String)

	Dim parts As Variant
	Dim msgTxt As String
	Dim i As Integer

	parts = Split(UCase(strName),"_")
	msgTxt = "data area abbreviation " & parts(1)

	For i = 0 To dataAreaCount - 1
		If dataAreas(i).nm = parts(1) Then
			If dataAreas(i).vlu = "A" Then Return ""
			Return msgTxt & " is deprecated"
		End If
	Next i

	Return msgTxt & " is invalid"

End Function


Function checkSuffix(strName As String, strType As String)

	'NSTest = checkCharCase(strName)
	'If NSTest <> "" Then Return(NSTest)

	parts = Split(UCase(strName),"_")
	i = UBound(parts)
	suffix = parts(i)
	starSuffix = "*"+suffix+"="+strType+"*"

	'Special case for ICW_LOAD_TM
	If UCase(strName) = "ICW_LOAD_TM" Then
		If strType = "TIME" Then
			Return ""
		ElseIf strType = "TIMESTAMP" Then
			Return ""
		End If
		Return "suffix is invalid for datatype " & strType
	End If

	' Special case for _PD and _PDZ suffix
	If UCase(suffix) = "PD" Or UCase(suffix) = "PDZ" Then
		If InStr("*TIMESTAMP*TIME*DATE*", strType) = 0 Then
			Return "suffix is invalid for datatype " & strType
		End If
		If attr.TeradataPeriod = False Then
			Return "suffix invalid as datatype not defined as PERIOD"
		End If
		If suffix = "PDZ" And attr.IsWithTimeZone = False Then
			Return "suffix is invalid for datatype " & strType
		End If
		If suffix = "PD" And attr.IsWithTimeZone = True Then
			Return "suffix is invalid for datatype " & strType & " WITH TIME ZONE"
		End If
		Return ""
	End If

	' If no suffix -> error
	If InStr(strName,"_") = 0 Then
		Return "has no suffix"
	End If

	' If suffix does not exist -> error
	If InStr(NSSuffixTypes,"*" & suffix & "*" ) = 0 Then
		Return "suffix is invalid"
	End If

	'See if we have a valid exception to the suffix-datatype rules
	If NSExceptionsCount > 0 Then
		e = False
		For n = 0 To NSExceptionsCount-1
			suffixException = NSSuffixExceptions(n)
			x = Len(suffixException.nm)
			If Right(UCase(strName),x) = suffixException.nm Then
				e = True
				If strType = suffixException.vlu Then
					Select Case suffixException.nm
					Case "_TSZ", "_PDZ"
						If attr.IsWithTimeZone = True Then Return ""
					Case Else
						Return ""
					End Select
				End If
			End If
		Next n
		If e = True Then Return "suffix is invalid for datatype " & strType
	End If

	' If suffix wrong format -> error
	If InStr(NSSuffixFormats,starSuffix) = 0 Then
		Return "suffix is invalid for datatype " & strType
	End If

	' Check WITH TIME ZONE not set
	If suffix <> "TSZ" And suffix <> "PDZ" Then
		If attr.IsWithTimeZone = True Then
			Return "suffix is invalid for " & strType & " WITH TIME ZONE"
		End If
	End If

	' If double-suffix -> error
	dbl_sfx = parts(i-1) & "_" & parts(i)
	If dbl_sfx = "TYP_CD" Or dbl_sfx = "TYP_NM" Or dbl_sfx = "TYP_DESC" Or dbl_sfx = "ID_TYP" Then
		If i > 1 Then
			If InStr(NSSuffixTypes,"*" & parts(i-2) & "*") > 0 Then
				Return "double-suffix " & parts(i-2) & "_" & dbl_sfx & " is invalid"
			End If
		Else
			Return "column name is invalid"
		End If
	Else
		If InStr(NSSuffixTypes,"*" & parts(i-1) & "*") > 0 Then
			Return "double-suffix " & parts(i-1) & "_" & parts(i) & " is invalid"
		End If
	End If

	Return ""

End Function


Function IsReservedWord(strWord As String) As Integer
	Dim keyWords As String
	strWord = "*"+UCase(strWord)+"*"
	keyWords = "*ABORT*ABORTSESSION*ABS*ACCESS_LOCK*ACCOUNT*ACOS*ACOSH*ADD*ADD_MONTHS*ADMIN*AFTER*AGGREGATE*ALL*ALTER*AMP*AND*ANSIDATE*ANY*ARRAY*AS"
	keyWords = keyWords+"*ASC*ASIN*ASINH*AT*ATAN*ATAN2*ATANH*ATOMIC*AUTHORIZATION*AVE*AVERAGE*AVG*BEFORE*BEGIN*BETWEEN*BINARY*BLOB*BOTH*BT*BUT*BY"
	keyWords = keyWords+"*BYTE*BYTEINT*BYTES*CALL*CALLED*CARDINALITY*CASE*CASE_N*CASESPECIFIC*CAST*CD*CHAR*CHAR_LENGTH*CHAR2HEXINT*CHARACTER"
	keyWords = keyWords+"*CHARACTER_LENGTH*CHARACTERS*CHARS*CHECK*CHECKPOINT*CLASS*CLOB*CLOSE*CLUSTER*CM*COALESCE*COLLATION*COLLECT*COLUMN*COMMENT"
	keyWords = keyWords+"*COMMIT*COMPRESS*CONDITION*CONNECT*CONSTRAINT*CONSTRUCTOR*CONTAINS*CONTINUE*CONVERT_TABLE_HEADER*CORR*COS*COSH*COUNT"
	keyWords = keyWords+"*COVAR_POP*COVAR_SAMP*CREATE*CROSS*CS*CSUM*CT*CUBE*CURRENT*CURRENT_DATE*CURRENT_ROLE*CURRENT_TIME*CURRENT_TIMESTAMP*CURRENT_USER"
	keyWords = keyWords+"*CURSOR*CV*CYCLE*DATA*DATABASE*DATABLOCKSIZE*DATE*DATEFORM*DAY*DEC*DECIMAL*DECLARE*DEFAULT*DEFERRED*DEGREES*DEL*DELETE*DESC"
	keyWords = keyWords+"*DESCRIBE*DETERMINISTIC*DIAGNOSTIC*DIAGNOSTICS*DISABLED*DISTINCT*DO*DOMAIN*DOUBLE*DROP*DUAL*DUMP*DYNAMIC*EACH*ECHO*ELSE*ELSEIF"
	keyWords = keyWords+"*ENABLED*END*END-EXEC*EQ*EQUALS*ERROR*ERRORFILES*ERRORTABLES*ESCAPE*ET*EXCEPT*EXCEPTION*EXEC*EXECUTE*EXISTS*Exit*Exp*EXPLAIN"
	keyWords = keyWords+"*EXTERNAL*EXTRACT*FALLBACK*FASTEXPORT*FETCH*FIRST*FLOAT*For*FOREIGN*Format*FOUND*FREESPACE*FROM*FULL*Function*GE*GENERATED*Get"
	keyWords = keyWords+"*GIVE*Global*GoTo*GRANT*GRAPHIC*GROUP*GROUPING*GT*HANDLER*HASH*HASHAMP*HASHBAKAMP*HASHBUCKET*HASHROW*HAVING*HELP*Hour*IDENTITY"
	keyWords = keyWords+"*If*IMMEDIATE*In*INCONSISTENT*Index*INITIATE*INNER*INOUT*Input*INS*INSERT*INSTEAD*Int*Integer*INTEGERDATE*INTERSECT*INTERVAL*INTO"
	keyWords = keyWords+"*Is*ISOLATION*ITERATE*Join*JOURNAL*KEY*KURTOSIS*LANGUAGE*LARGE*LAST*LE*LEADING*LEAVE*Left*LEVEL*Like*LIMIT*LN*LOADING*LOCAL"
	keyWords = keyWords+"*LOCATOR*Lock*LOCKING*Log*LOGGING*LOGON*Long*Loop*LOWER*LT*MACRO*MAP*MAVG*MAX*MAXIMUM*MCHARACTERS*MDIFF*MEMBER*MERGE*METHOD*MIN"
	keyWords = keyWords+"*MINDEX*MINIMUM*INUS*Minute*MLINREG*MLOAD*Mod*MODE*MODIFIES*MODIFY*MONITOR*MONRESOURCE*MONSESSION*Month*MSUBSTR*MSUM*MULTISET"
	keyWords = keyWords+"*NAMED*NATURAL*NE*New*NEW_TABLE*Next*NO*NONE*NORMALIZE*Not*NOWAIT*Null*NULLIF*NULLIFZERO*NUMERIC*Object*OBJECTS*OCTET_LENGTH*OF"
	keyWords = keyWords+"*Off*OLD*OLD_TABLE*On*ONLY*Open*Option*Or*ORDER*ORDINALITY*OUT*OUTER*OVER*OVERLAPS*OVERRIDE*PARAMETER*PARTITION*PASSWORD*PERCENT"
	keyWords = keyWords+"*PERCENT_RANK*PERM*PERMANENT*POSITION*PRECISION*PREPARE*Preserve*PRIMARY*Print*PRIOR*PRIVILEGES*Procedure*PROFILE*PROTECTION"
	keyWords = keyWords+"*Public*QUALIFIED*QUALIFY*QUANTILE*RADIANS*Random*RANGE*RANGE_N*RANK*Read*READS*REAL*RECURSIVE*REFERENCES*REFERENCING*REGR_AVGX"
	keyWords = keyWords+"*REGR_AVGY*REGR_COUNT*REGR_INTERCEPT*REGR_R2*REGR_SLOPE*REGR_SXX*REGR_SXY*REGR_SYY*RELATIVE*RELEASE*RENAME*REPEAT*Replace"
	keyWords = keyWords+"*REPLICATION*REQUEST*RESTART*RESTORE*RESTRICT*RESULT*Resume*RET*RETRIEVE*Return*RETURNS*REVALIDATE*REVOKE*Right*RIGHTS*Role"
	keyWords = keyWords+"*ROLLBACK*ROLLFORWARD*ROLLUP*ROW*ROW_NUMBER*ROWID*ROWS*SAMPLE*SAMPLEID*SCROLL*Second*SEL*Select*SESSION*Set*SETRESRATE*SETS"
	keyWords = keyWords+"*SETSESSRATE*SHOW*Sin*SINH*SIZE*SKEW*SMALLINT*SOME*SOUNDEX*SPECIFIC*SPOOL*SQL*SQLCODE*SQLERROR*SQLEXCEPTION*SQLSTATE*SQLTEXT"
	keyWords = keyWords+"*SQLWARNING*SQRT*SS*START*STARTUP*STATEMENT*STATISTICS*STDDEV_POP*STDDEV_SAMP*STEPINFO*STRING_CS*SUBSCRIBER*SUBSTR*SUBSTRING*SUM"
	keyWords = keyWords+"*SUMMARY*SUSPEND*SYSTEM*TABLE*Tan*TANH*TBL_CS*TEMPORARY*TERMINATE*Then*THRESHOLD*Time*TIMESTAMP*TIMEZONE_HOUR*TIMEZONE_MINUTE"
	keyWords = keyWords+"*TITLE*To*TOP*TRACE*TRAILING*TRANSACTION*TRANSLATE*TRANSLATE_CHK*Trigger*Trim*Type*UC*UESCAPE*UNDEFINED*UNDO*UNION*UNIQUE*UNKNOWN"
	keyWords = keyWords+"*UNNEST*Until*UPD*UPDATE*UPPER*UPPERCASE*USE*User*USING*VALUE*Values*VAR_POP*VAR_SAMP*VARBYTE*VARCHAR*VARGRAPHIC*VARYING*View"
	keyWords = keyWords+"*VOLATILE*WHEN*WHERE*While*WIDTH_BUCKET*With*WORK*Write*Year*ZEROIFNULL*ZONE*"

	keyWords = UCase(keyWords)

	If InStr(keyWords,strWord) = 0 Then
		Return 0
	Else
		Return -1
	End If
End Function

' Quicksort for simple data types.

' Indicate that a parameter is missing.
Const dhcMissing = -2

Sub dhQuickSort(varArray As Variant, _
 Optional intLeft As Integer = dhcMissing, _
 Optional intRight As Integer = dhcMissing)

    ' From "VBA Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 1997; Sybex, Inc. All rights reserved.

    ' Entry point for sorting the array.

    ' This technique uses the recursive Quicksort
    ' algorithm to perform its sort.

    ' In:
    '   varArray:
    '       A variant pointing to an array to be sorted.
    '       This had better actually be an array, or the
    '       code will fail, miserably. You could add
    '       a test for this:
    '       If Not IsArray(varArray) Then Exit Sub
    '       but hey, that would slow this down, and it's
    '       only YOU calling this procedure.
    '       Make sure it's an array. It's your problem.
    '   intLeft:
    '   intRight:
    '       Lower and upper bounds of the array to be sorted.
    '       If you don't supply these values (and normally, you won't)
    '       the code uses the LBound and UBound functions
    '       to get the information. In recursive calls
    '       to the sort, the caller will pass this information in.
    '       To allow for passing integers around (instead of
    '       larger, slower variants), the code uses -2 to indicate
    '       that you've not passed a value. This means that you won't
    '       be able to use this mechanism to sort arrays with negative
    '       indexes, unless you modify this code.
    ' Out:
    '       The data in varArray will be sorted.

    Dim i As Integer
    Dim j As Integer
    Dim varTestVal As Variant
    Dim intMid As Integer

    If intLeft = dhcMissing Then intLeft = LBound(varArray)
    If intRight = dhcMissing Then intRight = UBound(varArray)

    If intLeft < intRight Then
        intMid = (intLeft + intRight) \ 2
        varTestVal = UCase(varArray(intMid))
        i = intLeft
        j = intRight
        Do
            Do While UCase(varArray(i)) < varTestVal
                i = i + 1
            Loop
            Do While UCase(varArray(j)) > varTestVal
                j = j - 1
            Loop
            If i <= j Then
                SwapElements varArray, i, j
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= intMid Then
            Call dhQuickSort(varArray, intLeft, j)
            Call dhQuickSort(varArray, i, intRight)
        Else
            Call dhQuickSort(varArray, i, intRight)
            Call dhQuickSort(varArray, intLeft, j)
        End If
    End If


End Sub

Private Sub SwapElements(varItems As Variant, intItem1 As Integer, intItem2 As Integer)
    Dim varTemp As Variant

    varTemp = varItems(intItem2)
    varItems(intItem2) = varItems(intItem1)
    varItems(intItem1) = varTemp
End Sub

' DivorceMate Database Export and Consolidation Script
' Version: 1.0.0
' Last Modified: 2024-08-28
' Author: Divorcepath Corp.

Option Explicit

' Declare all variables at the beginning
Dim dbPath, outputDir, logFile, fso, conn, rs, rs2, ts, table, sql, fileName
Dim i, connected, tableCount, totalRows, exportedTables, startTime, endTime, duration
Dim connectionStrings(2), jsonFile, jsonContent, files, parties, children, lawyers, provinces, courts
Dim fileData, party, lawyer, child, jsonTs, totalRecords, processedRecords

' Create FileSystemObject once at the beginning
Set fso = CreateObject("Scripting.FileSystemObject")

' Function to safely create a folder
Sub CreateFolder(folderPath)
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
    End If
    If Err.Number <> 0 Then
        WriteLog logFile, "Error creating folder " & folderPath & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Function to write to log file
Sub WriteLog(logFile, message)
    On Error Resume Next
    Dim ts
    Set ts = fso.OpenTextFile(logFile, 8, True)
    ts.WriteLine Now & " - " & message
    ts.Close
    Set ts = Nothing
    If Err.Number <> 0 Then
        WScript.Echo "Error writing to log: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    WScript.Echo message
End Sub

' Function to parse date
Function ParseDate(dateString)
    If IsNull(dateString) Or Trim(dateString) = "" Then
        ParseDate = ""
    Else
        On Error Resume Next
        ParseDate = FormatDateTime(CDate(dateString), 2)
        If Err.Number <> 0 Then
            WriteLog logFile, "Warning: Unable to parse date '" & dateString & "'. Using original string."
            ParseDate = dateString
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

' Function to escape JSON string
Function EscapeJson(str)
    If IsNull(str) Then
        EscapeJson = "null"
    Else
        Dim result
        result = Replace(Replace(Replace(str, "\", "\\"), """", "\"""), vbCrLf, "\n")
        result = Replace(Replace(Replace(result, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
        EscapeJson = """" & result & """"
    End If
End Function

' Function to update and display progress
Sub UpdateProgress(current, total)
    Dim percent
    percent = Int((current / total) * 100)
    WScript.StdOut.Write vbCr & "Progress: " & percent & "% (" & current & "/" & total & ")"
End Sub

' Function to validate data
Function ValidateData(tableName, data)
    Dim warnings
    warnings = ""
    
    Select Case LCase(tableName)
        Case "tblfiles"
            If IsNull(data("FileID")) Or Trim(data("FileID")) = "" Then
                warnings = warnings & "Warning: File ID is missing. "
            End If
        Case "tblparties"
            If IsNull(data("PartyID")) Or Trim(data("PartyID")) = "" Then
                warnings = warnings & "Warning: Party ID is missing. "
            End If
            If IsNull(data("FullLegalName")) Or Trim(data("FullLegalName")) = "" Then
                warnings = warnings & "Warning: Full Legal Name is missing. "
            End If
        Case "tblchildren"
            If IsNull(data("ChildID")) Or Trim(data("ChildID")) = "" Then
                warnings = warnings & "Warning: Child ID is missing. "
            End If
        Case "tbllawyers"
            If IsNull(data("LawyerID")) Or Trim(data("LawyerID")) = "" Then
                warnings = warnings & "Warning: Lawyer ID is missing. "
            End If
    End Select
    
    ValidateData = warnings
End Function

' Main script
On Error Resume Next

' Use command line arguments if provided, otherwise use defaults
If WScript.Arguments.Count > 0 Then
    dbPath = WScript.Arguments(0)
Else
    WScript.Echo "Error: Please provide the path to the DivorceMate database."
    WScript.Quit 1
End If

If WScript.Arguments.Count > 1 Then
    outputDir = WScript.Arguments(1)
Else
    outputDir = ".\DivorceMateExport"
End If

CreateFolder outputDir
logFile = outputDir & "\export_log.txt"
jsonFile = outputDir & "\consolidated_data.json"

WriteLog logFile, "Starting export and consolidation process..."
WriteLog logFile, "Database path: " & dbPath
WriteLog logFile, "Output directory: " & outputDir

' Check if the database file exists
If Not fso.FileExists(dbPath) Then
    WriteLog logFile, "Error: Database file not found at " & dbPath
    WScript.Quit 2
End If
WriteLog logFile, "Database file found at " & dbPath

Set conn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

' Try different connection strings
connectionStrings(0) = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Persist Security Info=False;"
connectionStrings(1) = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Persist Security Info=False;"
connectionStrings(2) = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & dbPath & ";"

connected = False

For i = 0 To UBound(connectionStrings)
    conn.Open connectionStrings(i)
    If Err.Number = 0 Then
        connected = True
        WriteLog logFile, "Successfully connected using connection string: " & connectionStrings(i)
        Exit For
    Else
        WriteLog logFile, "Connection attempt failed with error: " & Err.Description & " (Error " & Err.Number & ")"
        Err.Clear
    End If
Next

If Not connected Then
    WriteLog logFile, "Error: Unable to connect to the database. Please ensure you have the correct version of Microsoft Access Database Engine installed."
    WScript.Quit 3
End If

' Get list of tables
Set rs = conn.OpenSchema(20) ' adSchemaTables
If Err.Number <> 0 Then
    WriteLog logFile, "Error getting table list: " & Err.Description & " (Error " & Err.Number & ")"
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    WScript.Quit 4
End If

On Error GoTo 0

' Initialize data structures
Set files = CreateObject("Scripting.Dictionary")
Set parties = CreateObject("Scripting.Dictionary")
Set children = CreateObject("Scripting.Dictionary")
Set lawyers = CreateObject("Scripting.Dictionary")
Set provinces = CreateObject("Scripting.Dictionary")
Set courts = CreateObject("Scripting.Dictionary")

' Export tables and load data
rs.MoveFirst
startTime = Timer
totalRecords = 0
processedRecords = 0

' Count total records
Do Until rs.EOF
    If rs("TABLE_TYPE") = "TABLE" Then
        sql = "SELECT COUNT(*) FROM [" & rs("TABLE_NAME") & "]"
        Set rs2 = conn.Execute(sql)
        totalRecords = totalRecords + rs2(0)
        Set rs2 = Nothing
    End If
    rs.MoveNext
Loop

rs.MoveFirst

Do Until rs.EOF
    If rs("TABLE_TYPE") = "TABLE" Then
        table = rs("TABLE_NAME")
        sql = "SELECT * FROM [" & table & "]"

        On Error Resume Next
        Set rs2 = CreateObject("ADODB.Recordset")
        rs2.Open sql, conn, 3, 3 ' 3, 3 for read-only, static cursor

        If Err.Number = 0 Then
            WriteLog logFile, "Exporting and loading table: " & table

            ' Load data into appropriate structure
            Select Case LCase(table)
                Case "tblfiles", "tblparties", "tblchildren", "tbllawyers", "tblprovinces", "tblcourts"
                    Do Until rs2.EOF
                        Dim dataRow, warnings
                        Set dataRow = CreateObject("Scripting.Dictionary")
                        For i = 0 To rs2.Fields.Count - 1
                            dataRow.Add rs2.Fields(i).Name, rs2.Fields(i).Value
                        Next
                        
                        warnings = ValidateData(table, dataRow)
                        If warnings <> "" Then
                            WriteLog logFile, warnings & " in " & table & " for ID: " & dataRow(rs2.Fields(0).Name)
                        End If

                        Select Case LCase(table)
                            Case "tblfiles"
                                files.Add dataRow("FileID"), dataRow
                            Case "tblparties"
                                If Not parties.Exists(dataRow("FileID")) Then
                                    Set parties(dataRow("FileID")) = CreateObject("Scripting.Dictionary")
                                End If
                                parties(dataRow("FileID")).Add dataRow("PartyID"), dataRow
                            Case "tblchildren"
                                If Not children.Exists(dataRow("FileID")) Then
                                    Set children(dataRow("FileID")) = CreateObject("Scripting.Dictionary")
                                End If
                                children(dataRow("FileID")).Add dataRow("ChildID"), dataRow
                            Case "tbllawyers"
                                lawyers.Add dataRow("LawyerID"), dataRow
                            Case "tblprovinces"
                                provinces.Add dataRow("ProvinceID"), dataRow("Name")
                            Case "tblcourts"
                                courts.Add dataRow("CourtID"), dataRow("CourtName")
                        End Select

                        processedRecords = processedRecords + 1
                        UpdateProgress processedRecords, totalRecords

                        rs2.MoveNext
                    Loop
            End Select
        Else
            WriteLog logFile, "Error exporting " & table & ": " & Err.Description & " (Error " & Err.Number & ")"
        End If
        If Not rs2 Is Nothing Then
            If rs2.State = 1 Then rs2.Close
            Set rs2 = Nothing
        End If
        On Error GoTo 0
    End If
    rs.MoveNext
Loop

WScript.StdOut.WriteLine

endTime = Timer
duration = endTime - startTime

WriteLog logFile, "Export and data loading completed in " & FormatNumber(duration, 2) & " seconds."

' Generate JSON
WriteLog logFile, "Generating consolidated JSON..."
startTime = Timer

jsonContent = "["

For Each fileID In files.Keys
    Set fileData = files(fileID)

    jsonContent = jsonContent & "{" & _
        """fileID"":" & EscapeJson(fileData("FileID")) & "," & _
        """name"":" & EscapeJson(fileData("FileName")) & "," & _
        """dateOfMarriage"":" & EscapeJson(ParseDate(fileData("DateOfMarriage"))) & "," & _
        """dateOfSeparation"":" & EscapeJson(ParseDate(fileData("DateOfSeparation"))) & "," & _
        """dateStartedLivingTogether"":" & EscapeJson(ParseDate(fileData("DateStartedLivingTogether"))) & "," & _
        """placeOfMarriage"":" & EscapeJson(fileData("PlaceOfMarriage")) & "," & _
        """neverLivedTogether"":" & LCase(fileData("NeverLivedTogether")) & "," & _
        """stillLivingTogether"":" & LCase(fileData("StillLivingTogether")) & "," & _
        """courtFileNumber"":" & EscapeJson(fileData("CourtFileNumber")) & "," & _
        """courtInfo"":{" & _
            """courtID"":" & EscapeJson(fileData("CourtID")) & "," & _
            """courtName"":" & EscapeJson(courts(fileData("CourtID"))) & "," & _
            """courtAddress"":" & EscapeJson(fileData("CourtAddress")) & "," & _
            """registryName"":" & EscapeJson(fileData("RegistryName")) & _
        "},"

    ' Add parties
    If parties.Exists(fileID) Then
        Dim partyIndex: partyIndex = 0
        For Each partyID In parties(fileID).Keys
            Set party = parties(fileID)(partyID)
            jsonContent = jsonContent & """party" & Chr(65 + partyIndex) & """:{" & _
                """partyID"":" & EscapeJson(party("PartyID")) & "," & _
                """name"":" & EscapeJson(party("FullLegalName")) & "," & _
                """address1"":" & EscapeJson(party("AddressLine2")) & "," & _
                """address2"":" & EscapeJson(party("AddressLine3")) & "," & _
                """address3"":" & EscapeJson(party("AddressLine4")) & "," & _
                """address4"":" & EscapeJson(party("AddressLine5")) & "," & _
                """address5"":" & EscapeJson(party("AddressLine6")) & "," & _
                """address6"":" & EscapeJson(party("AddressLine7")) & "," & _
                """address7"":" & EscapeJson(party("AddressLine8")) & "," & _
                """address8"":""""," & _
                """cityProv"":" & EscapeJson(party("CityMunicipality")) & "," & _
                """postalCode"":""""," & _
                """telephone"":""""," & _
                """fax"":""""," & _
                """emailOther"":""""," & _
                """birthDate"":" & EscapeJson(ParseDate(party("DateOfBirth"))) & "," & _
                """municipality"":" & EscapeJson(party("CityMunicipality")) & "," & _
                """provinceID"":" & EscapeJson(party("Province")) & "," & _
                """provinceName"":" & EscapeJson(provinces(party("Province"))) & "," & _
                """status"":" & EscapeJson(party("PersonType")) & "," & _
                """sin"":" & EscapeJson(party("SIN")) & "," & _
                """isApplicant"":" & LCase(party("IsApplicant")) & "," & _
                """isPrimary"":" & LCase(party("IsPrimary")) & "," & _
                """personType"":" & EscapeJson(party("PersonType"))

            ' Add lawyer info
            If lawyers.Exists(party("LawyerID")) Then
                Set lawyer = lawyers(party("LawyerID"))
                jsonContent = jsonContent & ",""lawyer"":{" & _
                    """lawyerID"":" & EscapeJson(lawyer("LawyerID")) & "," & _
                    """title"":" & EscapeJson(lawyer("Title")) & "," & _
                    """firstName"":" & EscapeJson(lawyer("FirstName")) & "," & _
                    """lastName"":" & EscapeJson(lawyer("LastName")) & "," & _
                    """firm"":" & EscapeJson(lawyer("Firm")) & "," & _
                    """addressLine1"":" & EscapeJson(lawyer("AddressLine1")) & "," & _
                    """addressLine2"":" & EscapeJson(lawyer("AddressLine2")) & "," & _
                    """city"":" & EscapeJson(lawyer("City")) & "," & _
                    """province"":" & EscapeJson(lawyer("Province")) & "," & _
                    """postalCode"":" & EscapeJson(lawyer("PostalCode")) & "," & _
                    """phone"":" & EscapeJson(lawyer("Phone")) & "," & _
                    """fax"":" & EscapeJson(lawyer("Fax")) & "," & _
                    """emailAddress"":" & EscapeJson(lawyer("EmailAddress")) & "," & _
                    """swearingCity"":" & EscapeJson(lawyer("SwearingCity")) & "," & _
                    """swearingProvince"":" & EscapeJson(lawyer("SwearingProvince")) & "," & _
                    """lsucNumber"":" & EscapeJson(lawyer("LSUCNumber")) & "," & _
                    """firmLawyer"":" & LCase(lawyer("FirmLawyer")) & _
                "}"
            Else
                jsonContent = jsonContent & ",""lawyer"":null"
            End If

            jsonContent = jsonContent & "},"
            partyIndex = partyIndex + 1
            If partyIndex = 2 Then Exit For ' Only process up to 2 parties
        Next
    End If

    ' Add children
    jsonContent = jsonContent & """children"":["
    If children.Exists(fileID) Then
        For Each childID In children(fileID).Keys
            Set child = children(fileID)(childID)
            jsonContent = jsonContent & "{" & _
                """childID"":" & EscapeJson(child("ChildID")) & "," & _
                """name"":" & EscapeJson(child("FullLegalName")) & "," & _
                """birthDate"":" & EscapeJson(ParseDate(child("DateOfBirth"))) & "," & _
                """sex"":" & EscapeJson(child("Sex")) & "," & _
                """residentIn"":" & EscapeJson(child("CityMunicipality")) & "," & _
                """gradeYearAndSchool"":" & EscapeJson(child("SchoolInfo")) & "," & _
                """residesWith"":" & EscapeJson(child("ResidesWith")) & "," & _
                """lawyerID"":" & EscapeJson(child("LawyerID")) & _
            "},"
        Next
        ' Remove trailing comma if children exist
        If children(fileID).Count > 0 Then
            jsonContent = Left(jsonContent, Len(jsonContent) - 1)
        End If
    End If
    jsonContent = jsonContent & "],"

    ' Add remaining file fields
    jsonContent = jsonContent & _
        """ppuRegistered"":" & LCase(fileData("PPURegistered")) & "," & _
        """unregisteredUses"":" & fileData("UnregisteredUses") & "," & _
        """revisionNumber"":" & fileData("RevisionNumber") & "," & _
        """archived"":" & LCase(fileData("Archived")) & "," & _
        """ourFileNumber"":" & EscapeJson(fileData("OurFileNumber")) & "," & _
        """lastUsed"":" & EscapeJson(fileData("LastUsed")) & "," & _
        """otherLawyer"":" & EscapeJson(fileData("OtherLawyer")) & "," & _
        """initialFilingDate"":" & EscapeJson(ParseDate(fileData("InitialFilingDate"))) & "," & _
        """daysFromInitialFiling"":" & fileData("DaysFromInitialFiling") & "," & _
        """finStatFromDate"":" & EscapeJson(ParseDate(fileData("FinStatFromDate"))) & "," & _
        """finStatToDate"":" & EscapeJson(ParseDate(fileData("FinStatToDate"))) & "," & _
        """selectedChild"":" & fileData("SelectedChild") & _
    "},"

    ' Update progress
    UpdateProgress files.Keys.Count, files.Count
Next

WScript.StdOut.WriteLine

' Remove trailing comma and close JSON array
If files.Count > 0 Then
    jsonContent = Left(jsonContent, Len(jsonContent) - 1)
End If
jsonContent = jsonContent & "]"

' Write JSON to file
On Error Resume Next
Set jsonTs = fso.CreateTextFile(jsonFile, True)
If Err.Number <> 0 Then
    WriteLog logFile, "Error creating JSON file: " & Err.Description
    WScript.Quit 5
End If
On Error GoTo 0

jsonTs.Write jsonContent
jsonTs.Close
Set jsonTs = Nothing

endTime = Timer
duration = endTime - startTime

WriteLog logFile, "JSON generation completed in " & FormatNumber(duration, 2) & " seconds."
WriteLog logFile, "Consolidated JSON file created: " & jsonFile

' Clean up
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If

WriteLog logFile, "Export and consolidation process finished."
WriteLog logFile, "IMPORTANT: Please verify the exported data for accuracy and completeness."
WriteLog logFile, "Handle the exported files in accordance with data protection and privacy regulations."

' Clean up objects
Set fso = Nothing
Set files = Nothing
Set parties = Nothing
Set children = Nothing
Set lawyers = Nothing
Set provinces = Nothing
Set courts = Nothing

WScript.Echo "Process completed. Please check the log file for details: " & logFile

' Final error check
If Err.Number <> 0 Then
    WriteLog logFile, "An unexpected error occurred: " & Err.Description
    WScript.Quit 6
End If
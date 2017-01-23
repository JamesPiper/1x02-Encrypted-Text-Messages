Attribute VB_Name = "modDataDumpToFile"
Option Explicit

Public Sub subDataDumpToFile()

    ' <ctrl>+<alt>D was pressed.
    ' Execute code to save the datadump to a text file then run notepad with the data in it.
    '

    ' Step 1 - Declare local variables.
    '-------------------------------------------------------------------------------------------------------------------------
    ' Declare filename variable.
    Dim strFileName As String
    ' Declare variable for API call.
    Dim lngAttributes As Long
    ' Declare return value variable.
    Dim lngReturnValue As Long
    ' File system handles.
    Dim objFSO As Object
    Dim objFile As Object
    ' Strong data typing doesn't want to work.
    Dim aryDataDump
    ' Loop counter.
    Dim x As Integer

    ' Step 2 - Create filename to store datadump.
    '-------------------------------------------------------------------------------------------------------------------------
    ' SubStep 1 - Declare filename variable.
    Dim strFileName As String
    
    ' SubStep 2 - Create filename to store datadump.
    
    ' Path and year.
    strFileName = App.Path & "\" & "FbTC" & Year(Date) & "."
    
    ' Add month to filename.
    If (Month(Date) < 10) Then
        strFileName = strFileName & "0" & Month(Date) & "."
    Else
        strFileName = strFileName & Month(Date) & "."
    End If ' If (Month(Date) < 10) Then
    
    ' Add day to filename.
    If (Day(Date) < 10) Then
        strFileName = strFileName & "0" & Day(Date) & "."
    Else
        strFileName = strFileName & Day(Date) & "."
    End If ' If (Day(Date) < 10) Then
    
    ' Add hour to filename.
    If (Hour(Time) < 10) Then
        strFileName = strFileName & "0" & Hour(Time) & "."
    Else
        strFileName = strFileName & Hour(Time) & "."
    End If ' If (Hour(Time) < 10) Then
    
    ' Add minute to filename.
    If (Minute(Time) < 10) Then
        strFileName = strFileName & "0" & Minute(Time) & "."
    Else
        strFileName = strFileName & Minute(Time) & "."
    End If ' If (Minute(Time) < 10) Then
    
    ' Add second to filename.
    If (Minute(Time) < 10) Then
        strFileName = strFileName & "0" & Second(Time)
    Else
        strFileName = strFileName & Second(Time)
    End If ' If (Second(Time) < 10) Then
    
    strFileName = strFileName & ".log"
    
    
    ' Step 3 - Create object as FSO.
    '-------------------------------------------------------------------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Step 4 - Create the file.
    '-------------------------------------------------------------------------------------------------------------------------
    ' True - overwrite any existing file.
    ' False - coding is ASCII.
    Set objFile = objFSO.CreateTextFile(strFileName, True, False)
    
    ' Step 5 - Write header info to the file.
    '-------------------------------------------------------------------------------------------------------------------------
    objFile.Write (strFileName & Chr(13) & Chr(10))
    objFile.Write (Date & "  " & Time & Chr(13) & Chr(10))
    objFile.Write (objDatePlusXDays.CalculatorName & " : " & objDatePlusXDays.FileName(constCurrentOutput) & Chr(13) & Chr(10))
    
    ' Step 6  - Get the data.
    '-------------------------------------------------------------------------------------------------------------------------
    aryDataDump = objDatePlusXDays.DataDump()

    ' Step 7 - Loop through the data and write it to the file.
    '-------------------------------------------------------------------------------------------------------------------------
    For x = 1 To UBound(aryDataDump)
        objFile.Write (aryDataDump(x) & Chr(13) & Chr(10))
    Next

    ' Step 8 - Close the file.
    '-------------------------------------------------------------------------------------------------------------------------
    objFile.Close
    
    ' Step 9 - Set objects to nothing.
    '-------------------------------------------------------------------------------------------------------------------------
    Set objFile = Nothing
    Set objFSO = Nothing
    
    ' Step 10 - Execute shell to run notepad.exe with this data file.
    '-------------------------------------------------------------------------------------------------------------------------
    ' Options: vbNormalFocus, vbMaximizedFocus, vbMinimizedFocus and so on.
    lngReturnValue = Shell("notepad.exe " & strFileName, vbMaximizedFocus)

    ' Step 11 - Set file attributes for the datadump file.
    '-------------------------------------------------------------------------------------------------------------------------
    lngAttributes = READONLY + HIDDEN
    ' Make API call to set the file attributes.
    lngReturnValue = SetFileAttributes(strFileName, lngAttributes)

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit     ' Remove in ASP
' ***** Comment out on the web.   *******
'<%                 ' Add in ASP
'
' clsFormatData for TaxCalcs.com
'
' Name: clsFormatData
' Type: Class Module
'
' Version 2006.01
'
' Date Started:     2004.03.31
' Date Completed:   2004.04.25
' Last Updated:     2007.02.17
'
' Author:   James Piper, BBA, CA
'
' (c) 2004 to 2007 by AarhusPiper, Inc.  *** All rights reserved.  Unauthorized use prohibited.  ***
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' A. Summary
' B. Purpose
' C. Dependencies
' D. Last Update & Explanation of the Changes
' E. Constants
' F. Variables
'    01. Private intIsUsedOnWeb
'
' G. Class_Initialize()
' H. Properties
'    101. Public Property Let IsUsedOnWeb(ByVal vNewValue)
'
'    801. Public Property Get ObjectName()
'    802. Public Property Get ReleaseVersion()
'    803. Public Property Get ReleaseDate()
'    804. Public Property Get ReleaseNotes()
'
' I. Procedures - Methods
'    101. Public Sub GreaterThanZero(ByVal IsValid, ByRef InputNum,
'                   ByRef OutputNum, ByVal xDecimal)
'    102. Public Sub GreaterThanZeroWithNegative(ByVal IsValid,
'                   ByRef InputNum, ByRef OutputNum, ByRef OutputNum_Negative, ByVal xDecimal)
'    103. Public Sub GreaterThanZeroBlankOk(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
'    104. Public Sub GreaterThanZeroBlankOkWithNegative(ByRef InputNum, ByRef OutputNum,
'                   ByRef OutputNum_Negative, ByVal xDecimal)
'    105. Public Sub PosOrNegOrBlankOk(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
'    106. Public Sub RealNumber(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
'    107. Public Sub TaxRate(ByVal InputNum, ByRef OutputNum, ByVal xDecimal)
'    108. Public Function DayOfWeekLongDate(ByVal InputDate)
'    109. Public Function ShortDate(ByVal InputDate)
'    110. Public Function LongDate(ByVal InputDate)
'    111. Public Function AbbreviatedLongDate(ByVal InputDate)
'
' J. Class_Terminate()
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' B. Purpose
'
' 01. Formats numbers to strings for use in outputing data.
' 02. The calc will be completely encapsulated in this code.
' 03. The object then can be used on all three platforms without changing the object:
'       a. VBA - Excel - for code development
'          (Copy the class module from Excel to VP & ASP file.)
'       b. VBScript - ASP - the web server for TaxCalcs.com
'       c. VB - Visual Basic - the stand alone program
' 04. This object is used by the main object of the calc.
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' C. Dependencies
'
' This object does not require any other object.
' It is called by main objects.
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' D. Last Update & Explanation of the Changes
'
'   Version 1.01
'   1. Added Public Sub TaxRate(ByVal InputNum, ByRef OutputNum, ByVal xDecimal)
'
'   Version 1.02
'   1. Added Public Sub LongDate(ByVal InputDate, ByRef OutputDate)
'      This formats a date value into a string version
'
'   Version 1.03
'   1. Changed GreaterThanZeroBlankOkWithNegative to remove ByRef OutputNum argument.
'
'   Version 1.04
'   1. Changing use of 'OutputNum = ""' in certain subs.
'
'   Version 1.05
'   1. Changing values for Outputnum to add trailing spaces.  This way neg- nums line up with pos+ nums.
'
'   Version 0.05 - 2004.08.22
'   1. Changed number system for versions.  A leading zero implies development.
'
'   Version 0.06 - 2004.08.28
'   1. Changing the way trailing spaces are handled for output data.  Based on IsUsedOnWeb.
'
'   Version 0.07 - 2004.09.02
'   1. Adding ReleaseDate and ReleaseVersion properties.
'
'   Version 0.08 - 2004.09.02
'   1. Fixing bug where CurrentInput etc. weren't defined.
'
'   Version 0.09 - 2004.10.17
'   1. Adding '&nbsp;' for web output.
'
'   Version 0.10 - 2004.10.20
'   1. Changing '&nbsp;' to be in a concantonated string.
'
'   Version 0.11 - 2004.10.21
'   1. Removing trailing spaces for TaxRate sub.
'
'   Version 0.12 - 2004.10.27
'   1. Adding code in TaxRate to show leading digits.
'
'   Version 0.13 - 2004.11.06
'   1. Using vbTrue instead of -1 as arguments.
'   2. Adding parenthesis for negative tax rate.
'
'   Version 0.14 - 2004.11.08
'   1. For Negative output and where value is zero, there's no brackets. Need to add trailing spaces.
'    Changes made in GreaterThanZeroWithNegative.
'
'   Version 0.15 - 2004.11.12
'   1. Changes in TaxRate to vary output with number of decimals when the input is not valid.
'
'   Version 0.16 - 2005.04.30
'   1. Changed TaxRate to exclude brackets for negative numbers.
'      (n.nn%) is now -n.nn%.
'      The brackets were causing run-time errors in VB.
'
'   Version 0.17 - 2005.05.06
'   1. Adding 'Public Sub ShortDate' so that 2005/4/4 becomes 2005/04/04.
'
'   Version 0.18 - 2005.07.22
'   1. Changing ShortDate to used IsUsedOnWeb to add trailing space.
'   2. Adding LongDate function to create Month DD, YYYY from a date.
'   3. Cleaning up the code to follow standard commenting.
'   4. Existing LongDate renamed DayOfWeekLongDate.
'   5. Adding AbbreviatedLongDate function to create Mth. DD, YYYY from a date.
'
'   Version 0.19 - 2005.08.22
'   1. In AbbreviatedLongDate, removing code that added trailing space.
'   2. In AbbreviatedLongDate, removing code that added trailing space.
'
'   Version 0.20 - 2005.09.16
'   1. Adding IsUsedOnWeb property.
'   2. Removing IsUsedOnWeb as argument for methods in this cls mod.
'
'   Version 2006.01 - [modMain Version 2006.04] - 2007.02.17
'   1. Overhaul before going to Version 2007.01.
'
'
'
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' E. Constants
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
Private CurrentInput    ' ***** Comment out on the web.   *******
Private CurrentOutput   ' ***** Comment out on the web.   *******


'*****************************************************************************************************************************
'
' Class clsFormatData            '(For use in VBScript - ASP - Web page.)
'
'*****************************************************************************************************************************

'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' F. Variables
'
' 01. Establish the data variables.
' 02. All variables and properties will not be typed - all implied variants.
' 03. All the variables will be private variables.
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

' 01.
Private bolIsUsedOnWeb


'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' G. Class_Initialize()
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()

    ' Set psuedo constants.
    CurrentInput = 4        ' ***** Comment out on the web.   *******
    CurrentOutput = 5       ' ***** Comment out on the web.   *******

End Sub

'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' H. Properties
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

'101. IsUsedOnWeb - LET
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Let IsUsedOnWeb(ByVal vNewValue)

    ' Store the value for being used on web site.
    bolIsUsedOnWeb = vNewValue
     
End Property ' 101. Public Property Let IsUsedOnWeb(ByVal vNewValue)
' 801. ObjectName - GET
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get ObjectName()

    ' This is a read-only property with the name of the object.
    '
    
    ObjectName = "Format Data"

End Property ' 801. Public Property Get ObjectName()
' 802. ReleaseVersion - GET
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get ReleaseVersion()
    
    ' This read-only property tells the user what the current version is.
    '
    
    ReleaseVersion = "2006.01"
    
End Property ' 802. Public Property Get ReleaseVersion()
' 803. ReleaseDate - GET
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get ReleaseDate(ByVal Index)
    
    ' This read-only property tells the user when the current object code was released for use.
    '
        
    ' Pass the appropriate value based on Index.
    Select Case Index
    Case CurrentInput
        ReleaseDate = DateSerial(2007, 2, 17)
    Case CurrentOutput
        ReleaseDate = "February 17, 2007"
    End Select
    
End Property ' 803. Public Property Get ReleaseDate()
' 804. ReleaseNotes - GET
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get ReleaseNotes()
    
    ' This read-only property tells users the history of releases for this object.
    '
    
    ' Step 1 - Declare local variables.
    '-------------------------------------------------------------------------------------------------------------------------
    ' Array index counter.
    Dim i
    ' Array to hold release notes.
    Dim aryReleaseNotes(5)
    
    ' Step 2 - Start with zero.
    '-------------------------------------------------------------------------------------------------------------------------
    i = 0
    
    ' Step 3 - Populate the array.
    '-------------------------------------------------------------------------------------------------------------------------
    ' Define the release notes to be displayed to users when using this calc.
    aryReleaseNotes(i) = ObjectName                                         ' "Format Data"
    i = i + 1
    aryReleaseNotes(i) = "Version: " & ReleaseVersion                       ' "Version 2006.01"
    i = i + 1
    aryReleaseNotes(i) = "Released: " & ReleaseDate(CurrentOutput)          ' "Released: November 2, 2006"
    i = i + 1
    aryReleaseNotes(i) = "1. Overhaul before going to Version 2007.01."
    i = i + 1
    aryReleaseNotes(i) = " "
    
    ' Step 4 - Add the release notes from the resource dat file.
    '-------------------------------------------------------------------------------------------------------------------------
    ' None.

    ' Step 5 - Send the Release Notes to the outside world.
    '-------------------------------------------------------------------------------------------------------------------------
    ReleaseNotes = aryReleaseNotes
    
End Property ' 804. Public Property Get ReleaseNotes()

'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' I. Procedures - Methods
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

' 101. GreaterThanZero
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub GreaterThanZero(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByVal xDecimal)

    ' For inputs that are greater than zero, format the number into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: nnnnn.nnnnn
    ' After:  nnnnn.nn, "nn,nnn.nn"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.
    
    ' Format only if valid number.
    If (IsValid) Then
        ' Valid number.
        
        ' Round the number.
        InputNum = Round(InputNum, xDecimal)
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = FormatNumber(InputNum, xDecimal) & "&nbsp;"
        Else
            OutputNum = FormatNumber(InputNum, xDecimal) & " "
        End If
        
    Else
        ' Invalid number.
            
        ' Set to zero.
        ' Previous set OutputNum to "" string.
        InputNum = 0
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = "0.00" & "&nbsp;"
        Else
            OutputNum = "0.00 "
        End If
        
    End If

End Sub ' 101. Public Sub GreaterThanZero(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
' 102. GreaterThanZeroWithNegative
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub GreaterThanZeroWithNegative(ByVal IsValid, ByRef InputNum, _
           ByRef OutputNum, ByRef OutputNum_Negative, ByVal xDecimal)

    ' This procedure is similar to GreaterThanZero but it adds a negative formating to show a negative amount.
    ' For inputs that are greater than zero, format into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: nnnnn.nnnnn, ""
    ' After:  nnnnn.nn, "nn,nnn.nn" and "(nn,nnn.nn)"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.

    ' Format only if number is valid.
    If (IsValid) Then
        ' Valid number.
        
        ' Round the number.
        InputNum = Round(InputNum, xDecimal)
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = FormatNumber(InputNum, xDecimal) & "&nbsp;"
        Else
            OutputNum = FormatNumber(InputNum, xDecimal) & " "
        End If
        
        ' Bracket version.
        OutputNum_Negative = InputNum * -1
        
        If (InputNum = 0) Then
            ' Need to add trailing space because there are no brackets.
            ' Add trailing space.
            If (bolIsUsedOnWeb) Then
                OutputNum_Negative = "0.00" & "&nbsp;"
            Else
                OutputNum_Negative = "0.00 "
            End If
        Else
            ' Format with brackets.
            OutputNum_Negative = FormatNumber(OutputNum_Negative, xDecimal, vbTrue, vbTrue)
        End If
        
    Else
        ' Invalid number.
        
        ' Set to zero.
        ' Previous set OutputNum & OutputNum_Negative to "" string.
        InputNum = 0
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = "0.00" & "&nbsp;"
            OutputNum_Negative = "0.00" & "&nbsp;"
        Else
            OutputNum = "0.00 "
            OutputNum_Negative = "0.00 "
        End If
        
    End If

End Sub ' 102. Public Sub GreaterThanZeroWithNegative(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByRef OutputNum_Negative, ByVal xDecimal)
' 103. GreaterThanZeroBlankOk
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub GreaterThanZeroBlankOk(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)

    ' For inputs that are zero or greater, format the number into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: nnnnn.nnnnn
    ' After:  nnnnn.nn, "nn,nnn.nn"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.
    
    ' Round valid number.
    If (IsNumeric(InputNum)) Then
        InputNum = Round(InputNum, xDecimal)
    Else
        InputNum = 0
    End If
    
    ' Format and add trailing space.
    If (bolIsUsedOnWeb) Then
        OutputNum = FormatNumber(InputNum, xDecimal) & "&nbsp;"
    Else
        OutputNum = FormatNumber(InputNum, xDecimal) & " "
    End If

End Sub ' 103. Public Sub GreaterThanZeroBlankOk(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
' 104. GreaterThanZeroBlankOkWithNegative
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub GreaterThanZeroBlankOkWithNegative(ByRef InputNum, _
           ByRef OutputNum_Negative, ByVal xDecimal)

    ' This procedure is similar to 03 but it adds a negative formating to show a negative amount.
    ' For inputs that are zero or greater, format into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: nnnnn.nnnnn, ""
    ' After:  nnnnn.nn, "nn,nnn.nn" and "(nn,nnn.nn)"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.

    ' Round valid number.
    If (IsNumeric(InputNum)) Then
        InputNum = Round(InputNum, xDecimal)
        OutputNum_Negative = InputNum * -1
    Else
        InputNum = 0
        OutputNum_Negative = 0
    End If
    
    ' Format and add trailing space.
    If (OutputNum_Negative < 0) Then
        OutputNum_Negative = FormatNumber(OutputNum_Negative, xDecimal, vbTrue, vbTrue)
    Else
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum_Negative = FormatNumber(OutputNum_Negative, xDecimal, vbTrue, vbTrue) & "&nbsp;"
        Else
            OutputNum_Negative = FormatNumber(OutputNum_Negative, xDecimal, vbTrue, vbTrue) & " "
        End If
    End If

End Sub ' 104. Public Sub GreaterThanZeroBlankOkWithNegative(ByRef InputNum, ByRef OutputNum_Negative, ByVal xDecimal)
' 105. PosOrNegOrBlankOk
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub PosOrNegOrBlankOk(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByVal xDecimal)

' Uses 'ByRef' to change the values of the variables in the call to this procedure.


    ' Round valid number.
    If (IsValid) Then
        ' Valid number.
        InputNum = Round(InputNum, xDecimal)
    Else
        If (IsNumeric(InputNum)) Then
            InputNum = Round(InputNum, xDecimal)
        Else
            InputNum = 0
        End If
    End If

    ' Format and add trailing space.
    If (InputNum < 0) Then
        OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue)
    Else
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue) & "&nbsp;"
        Else
            OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue) & " "
        End If
    End If

End Sub ' 105. Public Sub PosOrNegOrBlankOk(ByVal IsValid, ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
' 106. RealNumber
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub RealNumber(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)

    ' For inputs that are a real number, format into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: nnnnn.nnnnn, ""
    ' After:  nnnnn.nn, "nn,nnn.nn" OR "(nn,nnn.nn)"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.

    ' Round valid number.
    If (IsNumeric(InputNum)) Then
        InputNum = Round(InputNum, xDecimal)
    Else
        InputNum = 0
    End If
    
    ' Format and add trailing spaces.
    If (InputNum < 0) Then
        OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue)
    Else
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue) & "&nbsp;"
        Else
            OutputNum = FormatNumber(InputNum, xDecimal, vbTrue, vbTrue) & " "
        End If
    End If

End Sub ' 106. Public Sub RealNumber(ByRef InputNum, ByRef OutputNum, ByVal xDecimal)
' 107. TaxRate
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub TaxRate(ByVal InputNum, ByRef OutputNum, ByVal xDecimal)

    ' For inputs that are tax rates, format into a string with x decimal points.
    ' Example with xDecimal = 2
    ' Before: 0.nnnnn, ""
    ' After:  0.nn, "nn.n%"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.

    ' Format.
    If (IsNumeric(InputNum)) Then
        ' No brackets for negative number.
        OutputNum = FormatPercent(InputNum, xDecimal, vbTrue, vbFalse)
        
    Else
        ' Set to zero for invalid numbers.
        Select Case xDecimal
            Case 0
                OutputNum = "0%"
            Case 2
                OutputNum = "0.00%"
            Case 4
                OutputNum = "0.0000%"
            Case Else
                OutputNum = "0.00%"
        End Select
    End If
    
End Sub ' 107. Public Sub TaxRate(ByVal InputNum, ByRef OutputNum, ByVal xDecimal)
' 108. DayOfWeekLongDate
'-----------------------------------------------------------------------------------------------------------------------------
Public Function DayOfWeekLongDate(ByVal InputDate)

    ' Take a date value and create a string version of the date.
    ' Example: 30/05/2004 becomes "Sunday, May 30, 2004"
    ' Uses 'ByRef' to change the values of the variables in the call to this procedure.

    ' Local variable.
    Dim aryDayNames

    ' Define the days of the week.  Used in creating string date formated.
    aryDayNames = Array("", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

    ' Only format if valid date.
    If (IsDate(InputDate)) Then
        ' Valid date.
        
        ' Create formated date.
        DayOfWeekLongDate = aryDayNames(DatePart("w", InputDate)) & ", "
        DayOfWeekLongDate = DayOfWeekLongDate & MonthName(Month(InputDate), False) & " "
        DayOfWeekLongDate = DayOfWeekLongDate & Day(InputDate) & ", "
        DayOfWeekLongDate = DayOfWeekLongDate & Year(InputDate)
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            DayOfWeekLongDate = DayOfWeekLongDate & "&nbsp;"
        Else
            DayOfWeekLongDate = DayOfWeekLongDate & " "
        End If
    Else
        ' Invalid date.
        DayOfWeekLongDate = ""
    End If

End Function ' 108. Public Function DayOfWeekLongDate(ByVal InputDate)
' 109. ShortDate
'-----------------------------------------------------------------------------------------------------------------------------
Public Function ShortDate(ByVal InputDate)

    ' Take a date value and add leading zeros if needed.
    ' Example: 2004/5/5 becomes "2004/05/05"

    ' Format if valid date.
    If (IsDate(InputDate)) Then
    
        ' Year.
        ShortDate = Year(InputDate) & "/"
        
        ' Month.
        If (Month(InputDate) < 10) Then
            ShortDate = ShortDate & "0" & Month(InputDate) & "/"
        Else
            ShortDate = ShortDate & Month(InputDate) & "/"
        End If
        
        ' Day.
        If (Day(InputDate) < 10) Then
            ShortDate = ShortDate & "0" & Day(InputDate)
        Else
            ShortDate = ShortDate & Day(InputDate)
        End If
        
        ' Add trailing space.
        If (bolIsUsedOnWeb) Then
            ShortDate = ShortDate & "&nbsp;"
        Else
            ShortDate = ShortDate & " "
        End If

    End If

End Function ' Public Function ShortDate(ByVal InputDate)
' 110. LongDate
'-----------------------------------------------------------------------------------------------------------------------------
Public Function LongDate(ByVal InputDate)

    ' Take a date and create Month DD, YYY format.
    ' Example: 2003/10/11 becomes October 11, 2003

    ' Only format if valid date.
    If (IsDate(InputDate)) Then
    
        ' Format date.
        LongDate = MonthName(Month(InputDate), False) & " " & Day(InputDate) & ", " & Year(InputDate)

    End If

End Function ' 110. Public Function LongDate(ByVal InputDate)
' 111. AbbreviatedLongDate
'-----------------------------------------------------------------------------------------------------------------------------
Public Function AbbreviatedLongDate(ByVal InputDate)

    ' Take a date and create Month DD, YYY format.
    ' Example: 2003/10/11 becomes Oct. 11, 2003

    ' Only format if valid date.
    If (IsDate(InputDate)) Then
    
        ' Format date.
        If (Month(InputDate) = 5) Then
            ' May - don't add period.
            AbbreviatedLongDate = "May " & Day(InputDate) & ", " & Year(InputDate)
        Else
            ' Other months - add period.
            AbbreviatedLongDate = Left(MonthName(Month(InputDate), False), 3) & ". " & Day(InputDate) & ", " & Year(InputDate)
        End If
        
    End If

End Function ' 111. Public Function AbbreviatedLongDate(ByVal InputDate)
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' J. Class_Terminate()
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()


End Sub
'*****************************************************************************************************************************
'
'End Class ' clsFormatData            '(For use in VBScript - ASP - Web page.)
'
'*****************************************************************************************************************************
' Add in ASP -  %>

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidateDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' clsValidateDate for Fast by TaxCalcs
'
' Name: clsValidateDate
' Type: Class Module
'
' Version 2008.01
'
' Date Started:     2004.08.13 (Fast Version)   (Originally 2001.12.29)
' Date Completed:   2004.08.13                  (Originally 2002.02.09)
' Last Updated:     2009.03.14
'
' Author:   James Piper, BBA, CA, CMA
'
' (c) 2004 to 2009 by AarhusPiper, Inc.  *** All rights reserved.  Unauthorized use prohibited.  ***
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' A. Summary
'
' B. Purpose
' C. Dependencies
' D. Last Update & Explanation of the Changes
' E. Constants
' F. Variables
' G. Class_Initialize()
' H. Properties
' I. Procedures - Methods
' J. Class_Terminate()
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' B. Purpose.
'
'   This object is used with three text boxes to input a valid date:
'
'       1. txtDate = an 11 textbox for the date
'       2. lblErrMsg = msg of "Invalid Date"
'
'   Creating and destroying object as follows:
'
'       1. Dim dtDate1 As clsValidateDate (Declarations)
'       2. Set dtDate1 = New clsValidateDate (Form Load)
'       3. Set dtDate1 = Nothing (Form Unload)
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' C. Dependencies.
'
'   This object requires these events on the form:
'
'       1. Private Sub txtDate1_GotFocus()
'       2. Private Sub txtDate1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'       3. Private Sub txtDate1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'       4. Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
'       5. Private Sub txtDate1_KeyPress(KeyAscii As Integer)
'       6. Private Sub txtDate1_KeyUp(KeyCode As Integer, Shift As Integer)
'       7. Private Sub txtDate1_Validate()
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
' D. Last Update & Explanation of the Changes
'
'   Version 0.00 - 2004.08.13
'   1. This class object was originally created in 2002 and it worked fine.
'   2. With the restart of TaxCalcs as a stand alone program, the code is being tweaked.
'   3. Comments are being edited and added to.
'   4. Originally thought I would lose the use of the 'CommentMsg' but decided to keep it - lblDateComment on the form.
'   5. Adding comment structure as layed out in clsTestClass - including this.
'
'   Version 0.01 - [modMain Version 1.45] - 2005.05.31
'   1. Cleaning up the comments.
'   2. Cleaning up the code and further restrictions based on earliest and lastest dates allowed.
'
'   Version 0.02 - [modMain Version 1.46] - 2005.07.18
'   1. Removing code in CurrentDateKeyDown to deal with F9 key -- not needed.
'
'   Version 0.03 - [modMain Version 1.47] - 2005.08.09
'   1. Changes in Function IsDateValid to clarify the error message.
'
'   Version 0.04 - [modMain Version 2006.03] - 2006.11.10
'   1. Adding coding to allow for Feb. 29 in leap years.
'
'   Version 2008.01 - [modMain Version 2008.01] - 2009.03.14
'   1. Removing use of Errmsg box.
'
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
'
' H. Properties.
' 101. Public Property Let DefaultValue(ByVal vData As Variant)
' 101. Public Property Get DefaultValue() As Variant
' 102. Private Property Let DateOnFocus(ByVal vData As String)
' 102. Private Property Get DateOnFocus() As String
' 103. Private Property Let LastDate(ByVal vData As Date)
' 103. Private Property Get LastDate() As Date
' 104. Private Property Let EarliestDate(ByVal vData As Date)
' 104. Private Property Get EarliestDate() As Date
' 105. Private Property Let CurrentDate(ByVal vData As String)
' 106. Private Property Get CurrentDate() As String
' 107. Public Property Let CursorPostion(ByVal vData As Integer)
' 107. Public Property Get CursorPostion() As Integer
' 108. Private Property Let ErrorMessage(ByVal vData As String)
' 108. Private Property Get ErrorMessage() As String
'
' I. Methods.
' 101. Public Sub GotFocus(ByVal FocusDate As String, ByVal EDate As Date, ByVal LDate As Date)
' 102. Public Sub CurrentDateMouseDown(ByRef Button As Integer, ByRef Name As Object)
' 103. Public Sub CurrentDateMouseUp(ByRef Button As Integer, ByRef Name As Object)
' 104. Public Sub CurrentDateKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef Name As Object)
' 105. Public Sub CurrentDateKeyUp(ByRef KeyCode As Integer, ByRef Name As Object)
' 106. Public Function CurrentDateKeyPress(ByRef KeyAscii, ByRef TextBox As Object) As String
' 107. Public Function MonthDayMatch(ByRef vDate As String) As Boolean
' 108. Public Function IsDateValid(ByVal vDate As String) As Boolean
' 109. Public Function DateNotValid(ByRef DateTextBox As Object) As Boolean
'
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

' Local variable(s) to hold property value(s)
'-----------------------------------------------------------------------------------------------------------------------------

' Default value for the date (e.g., 'yyyy/mm/dd', or '2002/mm/dd').
Private strDefaultValue As String

' Original date value.
' If <esc> is pressed, the text box will be restored to this value.
Private strDateOnFocus As String

' Current value in the text box.
Private strCurrentDateValue As String

' Earliest valid date.
Private dtEarliestDate As Date

' Last valid date.
Private dtLastDate As Date

' yyYy of the EarliestDate.
' Used in limiting input.
Private strEarliestY3 As String

' For yyyY.
' Determine allowable values for Y.
Private strAllowableY4Pre2000 As String
Private strAllowableY4Post1999 As String

' Cursor position value when typing in data.
Private intCursorPostion As Integer

' comment error message
Private strErrorMessage As Variant
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
'
' G. Initialize the class object.
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    ' Set default value.
    strDefaultValue = "yyyy/mm/dd"

End Sub ' Private Sub Class_Initialize()

'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
'
' PROPERTIES
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

' 101. Public Property Let DefaultValue(ByVal vData As Variant)
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Let DefaultValue(ByVal vData As Variant)
    
    ' Set the default value.
    ' Value for the text box when the form is loaded.
    ' Press 'delete key' and this value is placed in the text box.
    ' The DefaultValue is set in Form_Load event.
    strDefaultValue = vData
    
End Property ' 101. Public Property Let DefaultValue(ByVal vData As Variant)
' 101. Public Property Get DefaultValue() As Variant
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get DefaultValue() As Variant

    ' Let the default value.
    DefaultValue = strDefaultValue
    
End Property ' 101. Public Property Get DefaultValue() As Variant
' 102. Private Property Let DateOnFocus(ByVal vData As String)
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Let DateOnFocus(ByVal vData As String)
    
    ' Set the original value of the date.
    ' This property holds the value in the text box as the focus in put on the text box.
    ' The user then can type in data, press <esc> and the text box will return to this original value.
    strDateOnFocus = vData
    
End Property ' 102. Private Property Let DateOnFocus(ByVal vData As String)
' 102. Private Property Get DateOnFocus() As String
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Get DateOnFocus() As String
    
    ' Get the original value of the date.
    ' This property holds the value in the text box as the focus in put on the text box.
    DateOnFocus = strDateOnFocus

End Property ' 102. Private Property Get DateOnFocus() As String
' 103. Private Property Let LastDate(ByVal vData As Date)
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Let LastDate(ByVal vData As Date)

    ' Set the last valid date.
    ' Any date entered after this time will create an error message.
    dtLastDate = vData

    ' Determine the list of allowable numbers for yyyY - Y4.
    ' If years are 2000 to 2005, strAllowableY4Pre2000 will be "012345".
    strAllowableY4Pre2000 = ""
    Dim x As Integer
    For x = Year(dtEarliestDate) To Year(dtLastDate)
        If (x < 2000) Then
            strAllowableY4Pre2000 = strAllowableY4Pre2000 & Right(CInt(x), 1)
        Else
            strAllowableY4Post1999 = strAllowableY4Post1999 & Right(CInt(x), 1)
        End If
    Next
    
End Property ' 103. Private Property Let LastDate(ByVal vData As Date)
' 103. Private Property Get LastDate() As Date
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Get LastDate() As Date
    
    ' Get the last valid date.
    ' Any date entered after this time will create an error message.
    LastDate = dtLastDate

End Property ' 103. Private Property Get LastDate() As Date
' 104. Private Property Let EarliestDate(ByVal vData As Date)
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Let EarliestDate(ByVal vData As Date)
    
    ' Set the last valid date.
    ' Here you tell this module what is the earliest possible day.
    ' Any date entered before this time will create an error message.
    dtEarliestDate = vData
    
    ' Get the third character of yyyy.
    ' This is either 8, 9 or 0.
    ' This value restricts inputs in CurrentDateKeyPress method.
    strEarliestY3 = Mid(Year(dtEarliestDate), 3, 1)

End Property ' 104. Private Property Let EarliestDate(ByVal vData As Date)
' 104. Private Property Get EarliestDate() As Date
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Get EarliestDate() As Date
    
    ' Set the last valid date.
    ' Any date entered before this time will create an error message.
    EarliestDate = dtEarliestDate

End Property ' 104. Private Property Get EarliestDate() As Date
' 105. Private Property Let CurrentDate(ByVal vData As String)
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Let CurrentDate(ByVal vData As String)
    
    ' Set the value of the current date in the text box.
    strCurrentDateValue = vData
    
End Property ' 105. Private Property Let CurrentDate(ByVal vData As String)
' 106. Private Property Get CurrentDate() As String
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Get CurrentDate() As String
    
    ' Get the value of the current date in the text box.
    CurrentDate = strCurrentDateValue
    
End Property ' 106. Private Property Get CurrentDate() As String
' 107. Public Property Let CursorPostion(ByVal vData As Integer)
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Let CursorPostion(ByVal vData As Integer)
    
    ' Set the current cursor position in the text box.
    intCursorPostion = vData

End Property ' 107. Public Property Let CursorPostion(ByVal vData As Integer)
' 107. Public Property Get CursorPostion() As Integer
'-----------------------------------------------------------------------------------------------------------------------------
Public Property Get CursorPostion() As Integer
    
    ' Get the current cursor position in the text box.
    CursorPostion = intCursorPostion

End Property ' 107. Public Property Get CursorPostion() As Integer
' 108. Private Property Let ErrorMessage(ByVal vData As String)
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Let ErrorMessage(ByVal vData As String)
    
    ' Set the error message about an invalid date.
    strErrorMessage = vData
    
End Property ' 108. Private Property Let ErrorMessage(ByVal vData As String)
' 108. Private Property Get ErrorMessage() As String
'-----------------------------------------------------------------------------------------------------------------------------
Private Property Get ErrorMessage() As String
    
    ' Get the error message about an invalid date.
    ErrorMessage = strErrorMessage

End Property ' 108. Private Property Get ErrorMessage() As String

'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
'
' METHODS
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

' 101. Public Sub GotFocus(ByVal FocusDate As String, ByVal EDate As Date, ByVal LDate As Date)
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub GotFocus(ByVal FocusDate As String, ByVal EDate As Date, ByVal LDate As Date)
    
    '
    ' Called by the txtDate_GotFocus event of the form.
    '
    
    ' Set value on focus - press <esc> will return this value - like undo.
    DateOnFocus = FocusDate
    
    ' Set the earliest valid date.
    EarliestDate = EDate
    
    ' Set the lastest valid date.
    LastDate = LDate
    
End Sub ' 101. Public Sub GotFocus(ByVal FocusDate As String, ByVal EDate As Date, ByVal LDate As Date)
' 102. Public Sub CurrentDateMouseDown(ByRef Button As Integer, ByRef Name As Object)
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub CurrentDateMouseDown(ByRef Button As Integer, ByRef Name As Object)
    
    '
    ' Called by the txtDate_MouseDown event on the form.
    '
    
    ' Don't want to allow any right click events like 'paste'.
    If (Button = 2) Then
        Clipboard.SetText Name.Text
        Name.SelStart = 0
        Name.SelLength = Len(Name.Text)
    End If

End Sub ' 102. Public Sub CurrentDateMouseDown(ByRef Button As Integer, ByRef Name As Object)
' 103. Public Sub CurrentDateMouseUp(ByRef Button As Integer, ByRef Name As Object)
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub CurrentDateMouseUp(ByRef Button As Integer, ByRef Name As Object)
    
    '
    ' Called by the txtDate_MouseUp event on the form.
    '
    
    ' Undo the effects of "cut" or "delete" when right click.
    If (Button = 2) Then
        If (Name.Text = "") Then
            Name.Text = DateOnFocus
        End If
    End If
    
End Sub ' 103. Public Sub CurrentDateMouseUp(ByRef Button As Integer, ByRef Name As Object)
' 104. Public Sub CurrentDateKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef Name As Object)
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub CurrentDateKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef Name As Object)
    
    '
    ' Called by the txtDate_KeyDown event on the form.
    '
    
    '
    ' Restrict certain key codes.
    '
    
    ' Don't want people "inserting" invalid characters.
    ' KeyCode = 45 for "insert", both "insert" keys.
    ' When the key is pressed, the code here kills the text.
    If (Shift) Then
        If (KeyCode = 45) Then
            ' <shift><ins> pressed.
            ' Zero out the code.
            Shift = 0
            KeyCode = 0
        End If
    End If

End Sub ' 104. Public Sub CurrentDateKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef Name As Object)
' 105. Public Sub CurrentDateKeyUp(ByRef KeyCode As Integer, ByRef Name As Object)
'-----------------------------------------------------------------------------------------------------------------------------
Public Sub CurrentDateKeyUp(ByRef KeyCode As Integer, ByRef Name As Object)
    
    '
    ' Called by the txtDate_KeyUp event on the form.
    '
    
    ' KeyCode = 46 is the "del" key, both "del" keys.
    If (KeyCode = 46) Then
        
        ' <del> key pressed.
        
        ' Restore the textbox to the default value.
        Name.Text = strDefaultValue
        CurrentDate = strDefaultValue
        
        '' Change caption to nothing.
        'ErrorMessage = ""
        'ErrMsgBox.ToolTipText = ""
        '
        '' Hide "invalid date" box.
        'ErrMsgBox.Visible = False
        '
    End If
    
End Sub ' 105. Public Sub CurrentDateKeyUp(ByRef KeyCode As Integer, ByRef Name As Object)
' 106. Public Function CurrentDateKeyPress(ByRef KeyAscii, ByRef TextBox As Object) As String
'-----------------------------------------------------------------------------------------------------------------------------
Public Function CurrentDateKeyPress(ByRef KeyAscii, ByRef TextBox As Object) As String
    
    '
    ' Validate data entry to a date textbox.
    ' Format is yyyy/mm/dd.
    ' Analyzes each keystroke to see if it should be displayed and added to the date.Text value.
    '

    ' Declare local variables.
    Dim strTextValue As String
    Dim intCursorPosition As Integer
    
    ' Set initial values.
    strTextValue = TextBox.Text
    intCursorPosition = TextBox.SelStart
    
    ' Part 1 - handle special non-character keys.
    '-----------------------------------------------------------------------------------------------------------------------------
    Select Case KeyAscii
    Case 8
        
        ' <backspace> key pressed, don't allow.
        Select Case intCursorPosition
            Case 0
                ' At the beginning, nothing to replace (i.e., backspace doesn't apply).
                CursorPostion = intCursorPosition
            Case 1, 2, 3, 4
                ' We are backspacing on one of the "years", replace with 'y'.
                Mid(strTextValue, intCursorPosition, 1) = "y"
                CursorPostion = intCursorPosition - 1
            Case 5
                ' dealing with '/", dont' change
                CursorPostion = 4
            Case 6, 7
                ' dealing with one of "m", replace
                Mid(strTextValue, intCursorPosition, 1) = "m"
                CursorPostion = intCursorPosition - 1
            Case 8
                ' dealing with '/", dont' change
                CursorPostion = 7
            Case 9, 10
                ' dealing with one of "m", replace
                Mid(strTextValue, intCursorPosition, 1) = "d"
                CursorPostion = intCursorPosition - 1
            Case Else
                ' anything else, dont' allow
                
        End Select ' Select Case intCursorPosition
        
        ' Don't want to change the cursor position.
        intCursorPosition = CursorPostion
        
    Case 13
        ' <enter> key pressed, change it to tab key.
        ' SendKeys not working with Windows 7.
        ' SendKeys "{tab}", True
        
        ' Move to the start of the textbox.
        intCursorPosition = 0
        
    Case 27
        ' <esc> key pressed.
        ' Restore the textbox to the default value.
        strTextValue = strDateOnFocus
        CurrentDate = strDateOnFocus
        
        '' Change caption to nothing.
        'ErrorMessage = ""
        'ErrMsgBox.ToolTipText = ""
        '
        '' Hide "invalid date" box.
        'ErrMsgBox.Visible = False
        
        ' Move to the start of the textbox.
        intCursorPosition = 0
        
    Case Else
        
        ' Part 2 - Entering numbers for yyyy/mm/dd, determine appropriatness based on cursor position.
        '-----------------------------------------------------------------------------------------------------------------------------
        Select Case intCursorPosition
        
        Case 0
            ' For y--- of yyyy.
            ' If it's not 1 or 2, don't allow.
            ' Keep cursor position if invalid character entered.
            
            Select Case KeyAscii
                Case 49
                    ' Where 1 is typed.
                    ' Use the 1 if the earliest date starts with a one.
                    ' The latest date will always start with 2.
                    If (Year(dtEarliestDate) - 2000 < 0) Then
                        ' Allowing 1 as first digit.
                        ' Since 1 for y--- means the rest will be 1yyy.
                        Mid(strTextValue, 1, 1) = "1"
                        
                        ' --y- is zero, change 190* to 1yyy.
                        If (Mid(strTextValue, 3, 1) = "0") Then
                            Mid(strTextValue, 1, 4) = "19yy"
                        End If
                        
                        ' Move to the cursor position for --y-.
                        intCursorPosition = 2
                        
                    Else
                        ' Ignoring 1 as first digit.
                    End If
                    
                Case 50
                    ' Where 2 is typed - always allow.
                    ' Regardless of what's in -y--, put "20y".
                    Mid(strTextValue, 1, 3) = "201"
                    ' Move to the cursor position for ---y.
                    intCursorPosition = 3
                Case Else
                    ' Any other key will be ignored.
            End Select ' Select Case KeyAscii
            
        Case 1
            ' For -y-- of yyyy.
            ' If 1--- then only 9.
            ' If 2--- then only 0.
            ' If y--- then 0 or 9 but only 0 if earliest date 2000 or later.
            ' Keep cursor position if invalid character entered.
    
            ' Determine action based on value in first y.
            Select Case Mid(strTextValue, 1, 1)
            
                Case "y"
                    ' y--- is y, allow 0 or 9.
                    
                    ' 0 typed - always allow.
                    If (KeyAscii = 48) Then
                        ' Add 200 to create 200*.
                        Mid(strTextValue, 1, 3) = "200"
                        ' Move to the cursor position ---y.
                        intCursorPosition = 3
                    End If
                    
                    ' 9 typed - allow if earliest date is before 2000.
                    If ((KeyAscii = 57) And (Year(dtEarliestDate) - 2000 < 0)) Then
                    
                        ' If --y- is not 8 or 9, replace with 19y otherwise just 19.
                        If (Mid(strTextValue, 3, 1) = "8") Or (Mid(strTextValue, 3, 1) = "9") Then
                            ' --y- is okay value, keep it.
                            ' Add 19 to create 19**.
                            Mid(strTextValue, 1, 2) = "19"
                        Else
                            ' --y- is not good value, change to y.
                            ' Add 19yy to create 19yy.
                            Mid(strTextValue, 1, 4) = "19yy"
                        End If
                        
                        ' Move to the cursor position --y-.
                        intCursorPosition = 2
                        
                    End If
                    
                Case "1"
                    ' y--- is 1, allow only 9.
                    If (KeyAscii = 57) Then
                        ' Add 9 to create 19**.
                        Mid(strTextValue, 2, 1) = "9"
                        ' Move to the cursor position --y-.
                        intCursorPosition = 2
                    End If
                    
                Case "2"
                    ' y--- is 2, allow only 0.
                    If (KeyAscii = 48) Then
                        ' Add 00 to create 200*.
                        ' Mid(strTextValue, 2, 2) = "00"
                        ' Move to the cursor position ---y.
                        ' intCursorPosition = 3
                        Mid(strTextValue, 2, 1) = "0"
                        intCursorPosition = 2
                    End If
                    
                Case Else
                    ' Ignore any other letters.
            End Select
            
            
        Case 2
            ' For --y- of yyyy.
            ' Possible values: 1, 2, 8 or 9.
            ' If yy-- then 1, 2, 8 or 9.
            ' If 19-- then 8 or 9.
            ' If 20-- then 1 or 2.
            
            ' Keep cursor position if invalid character entered.
    
            Select Case Mid(strTextValue, 1, 1)
            
                Case "y"
                    ' y--- is y.
                    
                    ' 0 typed - always allow.
                    If (KeyAscii = 48) Then
                        ' Add 201 to create 201*.
                        Mid(strTextValue, 1, 3) = "201"
                        ' Move to the cursor position ---y.
                        intCursorPosition = 3
                    End If
                    
                    ' 9 typed - allow if earliest date is before 2000.
                    If ((KeyAscii = 57) And (Year(dtEarliestDate) - 2000 < 0)) Then
                            ' Add 199 to create 199*.
                            Mid(strTextValue, 1, 3) = "199"
                            ' Move to the cursor position ---y.
                            intCursorPosition = 3
                    End If
                    
                Case "1"
                    ' y--- is 1.
                    
                    ' Allow the earliest date character of strEarliestY3.
                    If (Chr(KeyAscii) >= strEarliestY3) And (KeyAscii >= 48) And (KeyAscii <= 57) Then
                        ' If the character is >= the earliest third y, then allow it.
                        Mid(strTextValue, 3, 1) = Chr(KeyAscii)
                        ' Move to the cursor position ---y.
                        intCursorPosition = 3
                    End If
                    
                Case "2"
                    ' If the keypress character is 1 or 2 allow for 201x, 202x.
                    If (KeyAscii = 49) Or (KeyAscii = 50) Then
                        Mid(strTextValue, 3, 1) = Chr(KeyAscii)
                        ' Move to the cursor position ---y.
                        intCursorPosition = 3
                    End If
                    
                Case Else
                    ' Ignore any other letters.
            End Select
            
        Case 3
            ' For ---y of yyyy.
            ' Any 0 to 9 is possible.
            ' strAllowableY4Pre2000 is a string of allowable characters where Y--- is 2.
            ' strAllowableY4Post1999 is a string of allowable characters where Y--- is 1.
            ' strAllowableY4Pre2000 & strAllowableY4Post1999  are created in "Property Let LastDate."
            
            Select Case Mid(strTextValue, 1, 1)
            
                Case "y"
                    ' y--- is y.
                    If (InStr(strAllowableY4Pre2000, Chr(KeyAscii)) > 0) Then
                        ' Character is in the list of valid characters, accept.
                        Mid(strTextValue, 4, 1) = Chr(KeyAscii)
                        
                        ' Move to cursor position to m- of mm.
                        intCursorPosition = 5
                    End If
                    
                    If (InStr(strAllowableY4Post1999, Chr(KeyAscii)) > 0) Then
                        ' Character is in the list of valid characters, accept.
                        Mid(strTextValue, 4, 1) = Chr(KeyAscii)
                        
                        ' Move to cursor position to m- of mm.
                        intCursorPosition = 5
                    End If
                
                Case "1"
                    ' y--- is 1.
                    If (InStr(strAllowableY4Pre2000, Chr(KeyAscii)) > 0) Then
                        ' Character is in the list of valid characters, accept.
                        Mid(strTextValue, 4, 1) = Chr(KeyAscii)
                        
                        ' Move to cursor position to m- of mm.
                        intCursorPosition = 5
                    End If
                    
                Case "2"
                    ' y--- is 2.
                    If (InStr(strAllowableY4Post1999, Chr(KeyAscii)) > 0) Then
                        ' Character is in the list of valid characters, accept.
                        Mid(strTextValue, 4, 1) = Chr(KeyAscii)
                        
                        ' Move to cursor position to m- of mm.
                        intCursorPosition = 5
                    End If
                    
                Case Else
                    ' No other options.
                    
            End Select
            
        Case 4
            ' This is the position for the spacer.
            ' Allow only "/", then move to next space.
            
            If ((KeyAscii = 47)) Then
                Mid(strTextValue, 5, 1) = "/"
            End If
            
            ' Move to the cursor position /m-.
            intCursorPosition = 5
            
        Case 5
            ' For m- of mm.
            ' Allow 0 or 1.
            
            Select Case KeyAscii
                Case 48
                    ' 0 was pressed for m-.
                    If (Mid(strTextValue, 7, 1) = "0") Then
                        ' -m is 0, don't want 00, change to 0m.
                        Mid(strTextValue, 6, 2) = "0m"
                    Else
                        ' -m is m or a number, allow 0m or 01. . .
                        Mid(strTextValue, 6, 1) = Chr(KeyAscii)
                    End If
                    
                    ' Move to the next cursor position.
                    intCursorPosition = 6
                    
                Case 49
                    ' 1 was pressed for m-.
                    ' Only want 1m, 10, 11, or 12.
                    If (Mid(strTextValue, 7, 1) = "m" Or ((Mid(strTextValue, 7, 1) = "0") Or (Mid(strTextValue, 7, 1) = "1") Or (Mid(strTextValue, 7, 1) = "2"))) Then
                        ' Allow 1 and keep existing m2.
                        Mid(strTextValue, 6, 1) = "1"
                    Else
                        ' Allow 1, and change to 1m.
                        Mid(strTextValue, 6, 2) = "1m"
                    End If
                    
                    ' Move to the next cursor position.
                    intCursorPosition = 6
                    
                Case Else
                    ' Don't allow any other character.
                    
            End Select
            
        Case 6
            ' For -m of mm.
            ' If 1m, allow 0, 1 or 2.
            ' If 0m, allow 1 to 9.
            
            ' Based on the value of m-, allow certain chars. for -m
            Select Case Mid(strTextValue, 6, 1)
                Case "m"
                    ' m- of mm is m, allow any of 0 to 9
                    If ((KeyAscii >= 48) And (KeyAscii <= 57)) Then
                        Mid(strTextValue, 7, 1) = Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                
                Case "0"
                    ' m- of mm is 0, allow only 1 to 9
                    If ((KeyAscii >= 49) And (KeyAscii <= 57)) Then
                        Mid(strTextValue, 7, 1) = Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                    
                    If ((KeyAscii = 48)) Then
                        ' 0 pressed, change 00 to 10
                        Mid(strTextValue, 6, 2) = "1" & Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                    
                Case "1"
                    ' m- of mm is 1, allow only 0, 1, or 2
                    If ((KeyAscii >= 48) And (KeyAscii <= 50)) Then
                        Mid(strTextValue, 7, 1) = Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                    
                    If ((KeyAscii >= 51) And (KeyAscii <= 57)) Then
                        ' m- of mm is 1, allow 3 to 9 and change 1x to 0x
                        Mid(strTextValue, 6, 2) = "0" & Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                
                Case Else
                    ' nothing for anything else
            End Select
            
        Case 7
            ' This is the position for the spacer.
            ' Allow only "/", then move to next space.
            
            If ((KeyAscii = 47)) Then
                Mid(strTextValue, 8, 1) = "/"
            End If
            
            ' Move to the next cursor position.
            intCursorPosition = 8
        
        Case 8
            ' For d- of dd.
            ' Possible inputs are: 0, 1, 2 or 3.
            
            ' Based on -d.
            Select Case Mid(strTextValue, 10, 1)
            
                Case "d"
                    ' -d of dd is d, allow any of 0 to 3
                    If ((KeyAscii >= 48) And (KeyAscii <= 51)) Then
                        ' It's a 0 or 1 for all years.
                        Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                
                Case "0"
                    ' -d of dd is 0, allow only 1 to 3.
                    
                    ' Otherwise replace -d with d.
                    If ((KeyAscii >= 49) And (KeyAscii <= 51)) Then
                        Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                    
                    ' 0 pressed, change 00 to 0d
                    If ((KeyAscii = 48)) Then
                        Mid(strTextValue, 9, 2) = "0d"
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                
                Case "1"
                    ' -d of mm is 1, allow only 0, 1, 2, or 3.
                    If ((KeyAscii >= 48) And (KeyAscii <= 51)) Then
                        Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                    
                Case "2", "3", "4", "5", "6", "7", "8"
                    ' -d of mm is 2 to 8, allow only 0, 1, or 2.
                    If ((KeyAscii >= 48) And (KeyAscii <= 50)) Then
                        Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                    If ((KeyAscii = 51)) Then
                        ' 3 pressed, change 3x to 3d
                        Mid(strTextValue, 9, 2) = "3d"
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
            
                Case "9"
                    ' -d of dd is 9, allow only 0 or 1.
                    ' otherwise replace -d with d
                    ' 0 or 1 pressed
                    If ((KeyAscii = 48) Or (KeyAscii = 49)) Then
                        Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                    ' 2 pressed, change x9 to xd, for Feb.
                    If (KeyAscii = 50) Then
                        If (Mid(strTextValue, 6, 2) = "02") Then
                            ' dont want 29 for Feb.
                            Mid(strTextValue, 9, 2) = Chr(KeyAscii) & "d"
                            ' Jump to -d of dd.
                            intCursorPosition = 9
                        Else
                            ' not Feb so allow 29
                            Mid(strTextValue, 9, 1) = Chr(KeyAscii)
                            ' Jump to -d of dd.
                            intCursorPosition = 9
                        End If
                    End If
                    ' 3 pressed, change x9 to xd, in all cases.
                    If (KeyAscii = 51) Then
                        Mid(strTextValue, 9, 2) = Chr(KeyAscii) & "d"
                        ' Jump to -d of dd.
                        intCursorPosition = 9
                    End If
                
                Case Else
                    ' Nothing for anything else.
            End Select
           
        Case 9
            ' For -d of dd.
            ' Possible inputs are: 0 to 9.
            
            ' Nased on what's in yyyy-mm-Dd.
            Select Case Mid(strTextValue, 9, 1)
                Case "d"
                    ' d- of dd is d, allow any of 0 to 9.
                    If ((KeyAscii >= 48) And (KeyAscii <= 57)) Then
                        ' it's a 0 or 1 for all years
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                
                Case "0"
                    ' d- of dd is 0, allow 1 to 9.
                    ' otherwise replace d- with d0.
                    If ((KeyAscii >= 49) And (KeyAscii <= 57)) Then
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                    ' 0 pressed, change 00 to 0d.
                    If ((KeyAscii = 48)) Then
                        Mid(strTextValue, 9, 2) = "d" & Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
                
                Case "1"
                    ' d- of dd is 1, allow 0 to 9.
                    If ((KeyAscii >= 48) And (KeyAscii <= 57)) Then
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                
                Case "2"
                    ' d- of dd is 2, allow 0 to 8
                    ' otherwise replace 29 with d9 for February only
                    If ((KeyAscii >= 48) And (KeyAscii <= 56)) Then
                        ' 0 to 8 pressed.
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                    If (Mid(strTextValue, 6, 2) = "02") Then
                        ' 9 pressed, change 29 to d9 if Feb and not leap year.
                        If (((Left(strTextValue, 4)) Mod 4) = 0) Then
                            ' It's a leap year, allow 29.
                            Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                            ' Move to end.
                            intCursorPosition = 10
                        Else
                            ' Not a leap year, don't allow 29.
                            If ((KeyAscii = 57)) Then
                                Mid(strTextValue, 9 + 1, 2) = Chr(KeyAscii) ' "d" & Chr(KeyAscii)
                                ' Jump to d- of dd.
                                intCursorPosition = 8
                            End If
                        End If
                    Else
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                
                Case "3"
                    ' d- of dd is 3, allow 0 or 1
                    ' otherwise replace 3x with dx
                    If ((KeyAscii = 48) Or (KeyAscii = 49)) Then
                        Mid(strTextValue, 10, 1) = Chr(KeyAscii)
                        ' Move to end.
                        intCursorPosition = 10
                    End If
                    ' 2 pressed, change 3x to dx
                    If ((KeyAscii >= 50) And (KeyAscii <= 57)) Then
                        Mid(strTextValue, 9, 2) = "d" & Chr(KeyAscii)
                        ' Jump to d- of dd.
                        intCursorPosition = 8
                    End If
            
            End Select
            
        Case 10
            ' This is the 11th character which isn't used.
            ' Without this 11th character space, the text would "beep" when filled.
            intCursorPosition = 10
            
        Case Else
            ' There shouldn't be any other cases.
            
        End Select
    
    End Select
        
    ' In all cases, the key press is ignored.
    KeyAscii = 0
        
    ' Return the new date, if any.
    CurrentDateKeyPress = strTextValue
    
    ' Reset the current date.
    CurrentDate = strTextValue
    
    ' Reset the cursor position.
    CursorPostion = intCursorPosition

End Function ' 106. Public Function CurrentDateKeyPress(ByRef KeyAscii, ByRef TextBox As Object) As String
' 107. Public Function MonthDayMatch(ByRef vDate As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------
Public Function MonthDayMatch(ByRef vDate As String) As Boolean

    ' Use this function to test the validity of a date.
    ' The input is a textbox of 'yyyy/mm/dd'.
    ' The testing is based on an appropriate match between mm and dd.
    ' Testing also for valid numbers.
    
    ' Start with false, change to true if it's a valid day.
    MonthDayMatch = False
    
    ' Variables to store the portions of the date.
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    ' Get the yyyy, mm & dd portions of the date.
    strYear = Mid(vDate, 1, 4)
    strMonth = Mid(vDate, 6, 2)
    strDay = Mid(vDate, 9, 2)
    
    ' Test values.
    If (IsNumeric(strYear) And IsNumeric(strMonth) And IsNumeric(strDay)) Then
    
        ' The inputs are numbers, so we could have a valid date.
        ' Before doing the DateSerial function, it's important that the variables are good for a date.
        ' We need to check the 30 day versus 31 day months.
        ' In short that the max no. of days for a month is not exceeded.
        
        Select Case CInt(strMonth)
        
            Case 2
            
                ' It's February, can't be 29 or greater (except in leap years).
                ' Allow leap year.
                If ((strYear Mod 4) = 0) Then
                    ' Leap year.
                    If (CInt(strDay) > 29) Then
                        ' Invalid date.
                    Else
                        ' If we get to this point, then we have a valid date.
                        MonthDayMatch = True
                    End If
                Else
                    ' Not a leap year.
                    If (CInt(strDay) > 28) Then
                        ' Invalid date.
                    Else
                        ' If we get to this point, then we have a valid date.
                        MonthDayMatch = True
                    End If
                End If '
                
                
            Case 4, 6, 9, 11
            
                ' 30-day months, can't have 31 days.
                If (CInt(strDay) >= 31) Then
                    ' Invalid date.
                Else
                    MonthDayMatch = True
                End If
                
            Case Else
            
                If (CInt(strDay) >= 32) Then
                    ' Invalid date.
                Else
                    MonthDayMatch = True
                End If
            
        End Select
        
    Else
        ' It's not a valid date.
    End If

End Function ' 107. Public Function MonthDayMatch(ByRef vDate As String) As Boolean
' 108. Public Function IsDateValid(ByVal vDate As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------
Public Function IsDateValid(ByVal vDate As String) As Boolean

    ' Assume it is true, and go from there.
    IsDateValid = True
    
    ' Check is valid date (e.g., Feb 31 is not valid).
    If (MonthDayMatch(vDate)) Then
    
        ' Date is valid date.
        ' Check if date too early.
        If (DateDiff("d", dtEarliestDate, vDate) >= 0) Then
        
            ' If pos+, the current date is after the earlist date.
            ' The date is ok, don't show err msg.
            
            ' Check if date is too late.
            If (DateDiff("d", dtLastDate, vDate) <= 0) Then
                ' If neg-, the current date is before the last date.
                ' The date is ok, don't show err msg.
                ErrorMessage = ""
            Else
                ' The date is after the latest date.
                ' Set error msg.
                IsDateValid = False
                '
                ' Changed 2005.08.09.
                ' Clarify the error msg.
                '
                ErrorMessage = "Please enter a date on or after " & Year(dtEarliestDate) & "/" & Month(dtEarliestDate) & "/" & Day(dtEarliestDate) & _
                               "and on or before " & Year(dtLastDate) & "/" & Month(dtLastDate) & "/" & Day(dtLastDate) & ".  " & _
                               "Press 'Escape' to restore to the original value.  " & _
                               "Press 'Delete Key' to restore to the default value."
            End If
            
        Else
        
            ' The date is before earliest date.
            ' Set error msg.
            IsDateValid = False
            '
            ' Changed 2005.08.09.
            ' Clarify the error msg.
            '
            ErrorMessage = "Please enter a date on or after " & Year(dtEarliestDate) & "/" & Month(dtEarliestDate) & "/" & Day(dtEarliestDate) & _
                           "and on or before " & Year(dtLastDate) & "/" & Month(dtLastDate) & "/" & Day(dtLastDate) & ".  " & _
                           "Press 'Escape' to restore to the original value.  " & _
                           "Press 'Delete Key' to restore to the default value."
        End If
        
    Else
    
        ' Date may still be valid.
        If (vDate = "yyyy/mm/dd") Then
            ' Date still "yyyy/mm/dd", ok, no error message
            ErrorMessage = ""
        Else
            If (vDate = strDefaultValue) Then
                ' Date is default value - no err msg.
                ErrorMessage = ""
            Else
                ' Anything else is invalid date.
                IsDateValid = False
                '
                ' Changed 2005.08.09.
                ' Clarify the error msg.
                '
                ErrorMessage = "Please enter a valid date.  " & _
                               "Press 'Escape' to restore to the original value.  " & _
                               "Press 'Delete Key' to restore to the default value."
            End If
        End If
        
    End If

End Function ' 108. Public Function IsDateValid(ByVal vDate As String) As Boolean
' 109. Public Function DateNotValid(ByRef DateTextBox As Object) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------
Public Function DateNotValid(ByRef DateTextBox As Object) As Boolean

    ' Check validity of the date.
    ' If the date is 'bad', show error msg - ErrMsgBox becomes visible.
    ' Error msg shown by setting label.visible = true.
    ' If the date is still 'yyyy/mm/dd' then don't display error msg.
   
    ' The function return value is used 'Cancel' in txtDate_Validate event on the form.
    ' If Cancel is true, then it's not a valid input, it keeps focus.
    ' If Cancel is false, we're okay with the data, we can lose focus.
    If (IsDateValid(DateTextBox.Text)) Then
    
        ' The error msg maybe set to true, turn off.
        'ErrMsgBox.Visible = False
    
        ' Date not valid - return false.
        DateNotValid = False
    
    Else
        
        ' Show error message.
        'ErrMsgBox.Visible = True
        
        ' Date valid - return true.
        DateNotValid = True
        
    End If

    ' Set the ToolTipText for the ErrMsgBox to either "" or error message.
    'ErrMsgBox.ToolTipText = ErrorMessage

End Function ' 109. Public Function DateNotValid(ByRef DateTextBox As Object) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
'
' J. Terminate the class object.
'
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
    ' No code on terminating the object.
    
End Sub ' Private Sub Class_Terminate()

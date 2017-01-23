Attribute VB_Name = "modMain"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module: modMain
' Author: James Piper
' Date: March 2012
'
' Description:
' Main subroutine, the starting point.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Version 2
' 1. Use RT or BT instead of R10 or B10.
' 2. Don't have newline characters for available cards.
' 3. Fix bug when adding cards. RB3 s/b B3 not R3.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub Main()

    ' Starting point.
    '
    
    ' Load  and show the parent form.
    Dim objForm As Object
    Set objForm = New Parent
    objForm.Show
    
End Sub

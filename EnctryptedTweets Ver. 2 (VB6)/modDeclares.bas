Attribute VB_Name = "modDeclares"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module: modDeclares
' Author: James Piper
' Date: March 2012
'
' Description:
' 1. Declare global variables.
' 2. Declare API functions.
' 3. Declare constants for API functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Global Variables.
'
Public strUserMessage As String

Public strPlaintext As String
Public strPlaintextSpaced As String

Public strKey As String
Public strKeyWTrailingAs As String
Public strKeySpaced As String

Public strCiphertext As String
Public strCiphertextSpaced As String
Public strCiphertextPadded As String
Public strCiphertextPaddedSpaced As String

Public strCards As String
Public strCardsUsed As String
Public strAvailableCards As String

Public strCiphertextEntered As String
Public strCiphertextEnteredSpaced As String

Public strKeyEntered As String
Public strKeyEnteredSpaced As String

Public strPlaintextDeciphered As String
Public strPlaintextDecipheredSpaced As String

Public strFileName As String
Public strMsgBoxMessage  As String
Public intAnswer As Integer
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Global Constants.
'
Public Const constTitle = "EncryptedTweets"
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' For setting file attributes.
'
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileName As String, _
     ByVal dwFileAttributes As Long) As Long

Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileName As String) As Long

Public Const READONLY = &H1
Public Const HIDDEN = &H2
Public Const SYSTEM = &H4
Public Const ARCHIVE = &H20
Public Const NORMAL = &H80
'
' For setting file attributes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



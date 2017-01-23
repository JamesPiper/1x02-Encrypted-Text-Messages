Attribute VB_Name = "modLibrary"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module: modLibrary
' Author: James Piper
' Date: March 2012
'
' Description:
' Commmon subs and functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  1. Public Function ParseCardData(ByVal vText As String)
'  2. Public Sub GetKey(ByVal vText As String)
'  3. Public Sub ParseMessage(ByVal vText As String)
'  4. Public Sub ParseCiphertext(ByVal vText As String)
'  5. Public Sub ParseKey(ByVal vText As String)
'  6. Public Sub EncipherMessage(ByVal vText As String)
'  7. Public Function EncipherChar(ByVal vText As String, ByVal vKey As String)
'  8. Public Sub DecipherMessage(ByVal vText As String, ByVal vKey As String)
'  9. Public Function DecipherChar(ByVal vText As String, ByVal vKey As String)
' 10. Public Function AddSpacing(ByVal vText As String)
' 11. Public Function AddTrailingChars(ByVal vText As String, ByVal vChar As String)
' 12. Public Function RemovePadding(ByVal vText As String)
' 13. Public Sub LoadCardDataFromFile()
' 14. Public Sub SaveCardDataToFile(ByVal vText As String, ByVal vAdd As Boolean)
' 15. Public Sub RemoveUsedKeys()
' 16. Public Sub SaveMsgToFile(ByVal vDate As Date)
' 17. Public Sub SaveCardUsedDataToFile(ByVal vText As String)
' 18. Private Function TextLength(ByVal vTest As String) As String
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function ParseCardData(ByVal vText As String)

    ' Parse user card data.
    '
    ' Deck of card is 52 cards.
    ' There's 26 black and 26 red cards.
    '
    ' Red Ace (RA) is A.
    ' Red 2 (R2) is B.
    ' ...
    ' Red King (BK) is M.
    ' Black Ace (BA) is N.
    ' Black 2 (B2) is O.
    ' ...
    ' Black King (BK) is Z.
    '
    '
    ' What's done.
    ' 1. Card colour to uppercase (R or B).
    ' 2. Card value to uppercase (T, J, Q, K, A).
    ' 3. Ignore non-valid characters.
    ' 4. Build card data from colour and value.
    ' 5. Add space between each card.
    '
    
    Dim sCards As String
    Dim char As String
    Dim i As Integer
    Dim IsNewCard As Boolean
    Dim sCardColour As String
    Dim sCardValue  As String

    ' Initialize values.
    sCards = ""
    IsNewCard = False
    sCardColour = ""
    sCardValue = ""
    
    ' Remove spaces and non-valid characters.
    Do While (i <= Len(vText))
    
        ' Increment counter.
        i = i + 1
        
        ' Get current char to process
        ' and make uppercase.
        char = UCase(Mid(vText, i, 1))
        
        ' Get card value if any but only if we have B or R.
        If (IsNewCard) Then
            ' Face cards.
            If (char = "A") Or (char = "T") Or (char = "J") Or (char = "Q") Or (char = "K") Then
                sCardValue = char
                ' Reset.
                IsNewCard = False
            ElseIf IsNumeric(char) Then
                ' 2 to 9.
                If (char >= 2) And (char <= 9) Then
                    sCardValue = char
                    ' Reset.
                    IsNewCard = False
                End If
            End If
        End If

        ' Find card colour.
        If (char = "B") Or (char = "R") Then
            If (IsNewCard) Then
                ' Ignore previous start.
                ' Store value.
                sCardColour = char
            Else
                ' Previous char not B or R.
                ' Store value.
                sCardColour = char
                ' Start of card.
                IsNewCard = True
            End If
        End If
        
        ' Form card if we have colour and value.
        If (sCardColour <> "") And (sCardValue <> "") Then
            sCards = sCards + sCardColour + sCardValue + " "
            ' Reset for new card.
            sCardColour = ""
            sCardValue = ""
        End If
        
        ' Everything else is ignored.
        
    Loop
       
    ' Return parsed card data.
    ParseCardData = sCards

End Function
Public Sub GetKey(ByVal vText As String, ByVal vLen As Integer)

    ' Take card inputs and turn into a key.
    '
    ' Deck of card is 52 cards.
    ' There's 26 black and 26 red cards.
    '
    ' The key is uppercase alphabet (A-Z).
    ' There's 26 letters in the alphabet.
    ' This makes a even match between cards and letters.
    '
    ' Red Ace (RA) is A.
    ' Red 2 (R2) is B.
    ' ...
    ' Red King (BK) is M.
    ' Black Ace (BA) is N.
    ' Black 2 (B2) is O.
    ' ...
    ' Black King (BK) is Z.
    '
    ' Sample.
    ' vTest = B9 R3 BQ  R5 R4 B3 B3 R6 B2 RQ  R2 B7 B9 BK  B6 R10 B6 BK  R5 R6 R2 RA R2 B2 RQ  R8 R5 R10 B10 B9 B8 B8 B3 R5 BA RA B10
    '
    
    Dim char As String
    Dim i As Integer
    
    Dim iAdd As Integer
    Dim iCard As Integer
    Dim iNewCard As Boolean

    ' Reset counter.
    i = 0
    ' Reset key.
    strKey = ""
    strKeyWTrailingAs = ""
    strKeySpaced = ""
    
    ' Determine char value for each card.
    Do While (i <= Len(vText))

        ' Increment counter.
        i = i + 1
        
        ' Get current char to process.
        char = Mid(vText, i, 1)

        ' Start of new card.
        If (char = "R") Then
            iNewCard = True
            iAdd = 64
        End If
        If (char = "B") Then
            iNewCard = True
            iAdd = 77
        End If
        
        ' Get value of card.
        If (iNewCard) Then
        
            ' Increment counter.
            i = i + 1
            
            ' Get next char to process.
            char = Mid(vText, i, 1)
            
            ' Reset card value.
            iCard = 0
            
            Select Case char
            ' Ace.
            Case "A"
                iCard = 1
            ' King.
            Case "K"
                iCard = 13
            ' Queen.
            Case "Q"
                iCard = 12
            ' Jack.
            Case "J"
                iCard = 11
            Case "T"
                iCard = 10
            ' 2 to 9
            Case 2 To 9
                iCard = char
            End Select
            
            ' Calc key value.
            If (iNewCard) Then
                ' Value of card plus offset to get right ASCII value.
                iCard = iCard + iAdd
                ' Add to sting.
                strKey = strKey + Chr(iCard)
            End If
            
            ' Return to false.
            iNewCard = False
            
        End If

    Loop
    
    ' Shorten to length of message.
    strKey = Left(strKey, vLen)
    
    ' Add trailing As.
    strKeyWTrailingAs = AddTrailingChars(strKey, "A")
    
    ' Add space for every five characters.
    strKeySpaced = AddSpacing(strKeyWTrailingAs)
    
End Sub
Public Sub ParseMessage(ByVal vText As String)

    ' Take user message and do the following.
    ' 1. Ignore all non-alpha characters.
    ' 2. All in lowercase.
    '
    ' User msg in vText
    ' Return & store parsed text in strPlaintext
    '
    
    Dim i As Integer
    Dim char As Integer
    Dim strChar As String
    
    ' Reset strings.
    strPlaintext = ""
    strPlaintextSpaced = ""
    
    ' Loop through the message.
    For i = 1 To Len(vText)
    
        ' Get current char.
        char = Asc(Mid(vText, i, 1))
        
        ' Act based on type of char.
        Select Case char
        Case 97 To 122
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char)
        Case 65 To 90
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char + 32)
        Case Else
        ' Ignore everything else.
            strChar = ""
        End Select
        
        ' Build string.
        strPlaintext = strPlaintext & strChar
    Next

    ' Add space for every five characters.
    strPlaintextSpaced = AddSpacing(strPlaintext)
    
End Sub
Public Sub ParseCiphertext(ByVal vText As String)

    ' Parse ciphertext.
    ' 1. Ignore all non-alpha characters.
    ' 2. Make uppercase.
    '
    
    Dim i As Integer
    Dim char As Integer
    Dim strChar As String
    
    ' Reset strings.
    strCiphertextEntered = ""
    strCiphertextEnteredSpaced = ""
    
    ' Loop through the message.
    For i = 1 To Len(vText)
    
        ' Get current char.
        char = Asc(Mid(vText, i, 1))
        
        ' Act based on type of char.
        Select Case char
        Case 97 To 122
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char - 32)
        Case 65 To 90
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char)
        Case Else
        ' Ignore everything else.
            strChar = ""
        End Select
        
        ' Build string.
        strCiphertextEntered = strCiphertextEntered & strChar
    Next
    
End Sub
Public Sub ParseKey(ByVal vText As String)

    ' Parse ciphertext.
    ' 1. Ignore all non-alpha characters.
    ' 2. Uppercase.
    '
    
    Dim i As Integer
    Dim char As Integer
    Dim strChar As String
    
    ' Reset strings.
    strKeyEntered = ""
    strKeyEnteredSpaced = ""
    
    ' Loop through the message.
    For i = 1 To Len(vText)
    
        ' Get current char.
        char = Asc(Mid(vText, i, 1))
        
        ' Act based on type of char.
        Select Case char
        Case 97 To 122
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char - 32)
        Case 65 To 90
        ' Lowercase alpha chars.
        ' Allow.
            strChar = Chr(char)
        Case Else
        ' Ignore everything else.
            strChar = ""
        End Select
        
        ' Build string.
        strKeyEntered = strKeyEntered & strChar
    Next

    ' Add space for every five characters.
    strKeyEnteredSpaced = AddSpacing(strKeyEntered)
    
End Sub
Public Sub EncipherMessage(ByVal vText As String)

    ' Encipher user's message with key.
    '
    
    ' The Algorithm.
    ' 1. Covert plaintext to uppercase.
    ' 2. Convert to ASCII value.
    ' 2. Subtract 65.
    ' 3. Convert key to ASCII value.
    ' 4. Subtract 65.
    ' 5. Add 2 and 4.
    ' 6. Calculate modulus of result with 26.
    ' 7. Add 65.
    ' 8. Convert result to ASCII char (A to Z).
    '
    
    ' Counter.
    Dim i As Integer
    ' User message.
    Dim MsgChar As String
    ' Encryption Key.
    Dim KeyChar As String
    ' Chiper.
    Dim Cipher As String
    Dim strPad  As String
    
    ' Reset value.
    strCiphertext = ""
    strCiphertextSpaced = ""
    strCiphertextPadded = ""
    strCiphertextPaddedSpaced = ""

    For i = 1 To Len(vText)
    
        ' Get current character from message.
        MsgChar = Mid(vText, i, 1)
        
        ' Get current character from key.
        KeyChar = Mid(strKey, i, 1)
        
        ' Get cipher for one character.
        Cipher = EncipherChar(MsgChar, KeyChar)
        
        ' Build ciphertext.
        strCiphertext = strCiphertext + Cipher
        strCiphertextPadded = strCiphertextPadded & Cipher
        
        ' Add padding between each letter of the ciphertext.
        ' Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        ' Lower: A (65)
        ' Upper: Z (90)
        strPad = Int((90 - 65 + 1) * Rnd + 65)
        strCiphertextPadded = strCiphertextPadded & Chr(strPad)
    
    Next

    ' Add trailing Xs.
    strCiphertext = AddTrailingChars(strCiphertext, "X")
    strCiphertextSpaced = strCiphertext
    strCiphertextPadded = AddTrailingChars(strCiphertextPadded, "X")
    strCiphertextPaddedSpaced = strCiphertextPadded

    ' Add spacing.
    strCiphertextSpaced = AddSpacing(strCiphertext)
    strCiphertextPaddedSpaced = AddSpacing(strCiphertextPadded)
    
End Sub
Public Function EncipherChar(ByVal vText As String, ByVal vKey As String)

    ' Enciphers one char using the algorithm.
    '
    
    ' The Algorithm.
    ' 1. Covert plaintext to uppercase.
    ' 2. Convert to ASCII value.
    ' 2. Subtract 65.
    ' 3. Convert key to ASCII value.
    ' 4. Subtract 65.
    ' 5. Add 2 and 4.
    ' 6. Calculate modulus of result with 26.
    ' 7. Add 65.
    ' 8. Convert result to ASCII char (A to Z).
    '
    
    ' User message.
    Dim MsgChar As String
    Dim iMsgChar As Integer
    ' Encryption Key.
    Dim KeyChar As String
    Dim iKeyChar As Integer
    ' Chiper.
    Dim Cipher As String
    Dim iCipher As Integer

    ' To uppercase.
    MsgChar = UCase(vText)
    
    ' Convert to ASCII value.
    iMsgChar = Asc(MsgChar)
    
    ' Subtract 65.
    iMsgChar = iMsgChar - 65
    
    ' To uppercase
    KeyChar = UCase(vKey)
    
    ' Convert to ASCII value.
    iKeyChar = Asc(KeyChar)
    
    ' Subtract 65.
    iKeyChar = iKeyChar - 65
    
    ' Add message with key.
    iCipher = iMsgChar + iKeyChar
    
    ' Calc modulus with 26.
    iCipher = iCipher Mod 26
    
    ' Add 65.
    iCipher = iCipher + 65
    
    ' Convert to ASCII value.
    Cipher = Chr(iCipher)

    ' Return value.
    EncipherChar = Cipher
    
End Function
Public Sub DecipherMessage(ByVal vText As String, ByVal vKey As String)

    ' Deccipher user's message with key.
    '
    
    ' The Algorithm.
    ' 1. Covert plaintext to uppercase.
    ' 2. Convert to ASCII value.
    ' 2. Subtract 65.
    ' 3. Convert key to ASCII value.
    ' 4. Subtract 65.
    ' 5. Add 2 and 4.
    ' 6. Calculate modulus of result with 26.
    ' 7. Add 65.
    ' 8. Convert result to ASCII char (A to Z).
    '
    
    ' Counter.
    Dim i As Integer
    ' User message.
    Dim MsgChar As String
    ' Encryption Key.
    Dim KeyChar As String
    ' Chiper.
    Dim Cipher As String
    
    ' Reset value.
    strPlaintextDeciphered = ""
    strPlaintextDecipheredSpaced = ""

    For i = 1 To Len(vText)
    
        ' Get current character from message.
        MsgChar = Mid(vText, i, 1)
        
        ' Get current character from key.
        KeyChar = Mid(vKey, i, 1)
        
        ' Get cipher for one character.
        Cipher = DecipherChar(MsgChar, KeyChar)
        
        ' Build ciphertext.
        strPlaintextDeciphered = strPlaintextDeciphered + Cipher
    Next
    
    ' Add space for every five characters.
    strPlaintextDecipheredSpaced = AddSpacing(strPlaintextDeciphered)
    
    ' Not need to add trailing Xs.

End Sub
Public Function DecipherChar(ByVal vText As String, ByVal vKey As String)

    ' Deciphers one char using the algorithm.
    '
    
    ' The Algorithm.
    ' 1. Covert ciphertext to uppercase.
    ' 2. Convert to ASCII value.
    ' 2. Subtract 65.
    ' 3. Convert key to ASCII value.
    ' 4. Subtract 65.
    ' 5. Subtract 4 from 2.
    ' 6. Calculate modulus of result with 26.
    ' 7. Add 65 + 32.
    ' 8. Convert result to ASCII char (A to Z).
    '
    
    ' User message.
    Dim MsgChar As String
    Dim iMsgChar As Integer
    ' Encryption Key.
    Dim KeyChar As String
    Dim iKeyChar As Integer
    ' Chiper.
    Dim Cipher As String
    Dim iCipher As Integer

    ' To uppercase.
    Cipher = UCase(vText)
    
    ' Convert to ASCII value.
    iCipher = Asc(Cipher)
    
    ' Subtract 65.
    iCipher = iCipher - 65
    
    ' To uppercase
    KeyChar = UCase(vKey)
    
    ' Convert to ASCII value.
    iKeyChar = Asc(KeyChar)
    
    ' Subtract 65.
    iKeyChar = iKeyChar - 65
    
    ' Add message with key.
    iMsgChar = iCipher - iKeyChar
    
    ' Calc modulus with 26.
    If (iMsgChar >= 0) Then
        iMsgChar = iMsgChar Mod 26
    Else
        ' Doesn't work right with negative numbers.
        ' Work around.
        iMsgChar = iMsgChar + 26
    End If

    ' Add 65 for char in ASCII table.
    ' Add 32 to have lowercase.
    iMsgChar = iMsgChar + 65 + 32
    
    ' Convert to ASCII value.
    MsgChar = Chr(iMsgChar)

    ' Return value.
    DecipherChar = MsgChar
    
End Function
Public Function AddSpacing(ByVal vText As String)

    ' Add a space every 5 characters.
    '
    
    ' Counter.
    Dim i As Integer
    Dim MsgChar As String
    
    For i = 1 To Len(vText)
    
        ' Get current character from message.
        MsgChar = Mid(vText, i, 1)
        
        ' Build string.
        AddSpacing = AddSpacing & MsgChar
        
        ' Add space
        If (i Mod 5 = 0) Then
            AddSpacing = AddSpacing & Chr(32)
        End If

    Next

End Function
Public Function AddTrailingChars(ByVal vText As String, ByVal vChar As String)

    ' Add either X or A to the end of the text.
    ' It's common practice to pad the end of messages.
    '
    
    ' Counter.
    Dim i As Integer
    ' Number to get to five.
    Dim iLen As Integer
    iLen = 5 - Len(vText) Mod 5
    
    ' Variable numbers of characters.
    AddTrailingChars = vText
    For i = iLen To 1 Step -1
        AddTrailingChars = AddTrailingChars & vChar
    Next
    
    ' Block of five.
    AddTrailingChars = AddTrailingChars & vChar & vChar & vChar & vChar & vChar

End Function
Public Function RemovePadding(ByVal vText As String)
    
    ' Remove every other character of ciphertext.
    '

    ' Counter.
    Dim i As Integer
    
    ' Loop through the text.
    For i = 1 To Len(vText)
        ' Only want the odd number characters.
        If (i Mod 2 = 1) Then
            RemovePadding = RemovePadding & Mid(vText, i, 1)
        End If
    Next
    
End Function
Public Sub LoadCardDataFromFile()

    ' Load available card data from file.
    '
    
    ' Set filename.
    strFileName = App.Path & "\AvailableCardValues.dat"
    
    ' File system handles.
    Dim objFSO As Object
    Dim objFile As Object
    
    ' Create object as FSO.
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Open the file.
    ' 1. Filename.
    ' 2. iomode - ForReading (1) or ForAppending (8).
    ' 3. create - T/F.
    If (objFSO.FileExists(strFileName)) Then
        ' Set
        Set objFile = objFSO.OpenTextFile(strFileName, 1, False)
        ' Loop through file and read data.
        Do While Not objFile.AtEndOfStream
            strAvailableCards = strAvailableCards & objFile.ReadLine
        Loop
        ' Close the file.
        objFile.Close
    End If
    
    ' Set objects to nothing.
    Set objFile = Nothing
    Set objFSO = Nothing
    
End Sub
Public Sub SaveCardDataToFile(ByVal vText As String, ByVal vAdd As Boolean)

    ' Store card data to file.
    '
    
    ' Set filename.
    strFileName = App.Path & "\AvailableCardValues.dat"
    
    ' File system handles.
    Dim objFSO As Object
    Dim objFile As Object
    
    ' Create object as FSO.
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Open the file.
    ' 1. Filename.
    ' 2. iomode - ForReading (1), ForWriting (2) or ForAppending (8).
    ' 3. create - T/F.
    If (objFSO.FileExists(strFileName)) Then
        ' Set
        If (vAdd) Then
            Set objFile = objFSO.OpenTextFile(strFileName, 8, False)
        Else
            Set objFile = objFSO.OpenTextFile(strFileName, 2, False)
        End If
        ' Write line of data to file.
        objFile.Write (vText) '  & Chr(13) & Chr(10))
        ' Close the file.
        objFile.Close
    End If
    
    ' Set objects to nothing.
    Set objFile = Nothing
    Set objFSO = Nothing
    
End Sub
Public Sub RemoveUsedKeys()

    ' Remove used keys from the file.
    '
    
    Dim i As Integer
    Dim iStart As Integer
    
    ' Start at beginning
    iStart = 1
    For i = 1 To Len(strPlaintext)
        
        ' Find end of next card.
        iStart = InStr(iStart, strAvailableCards, " ") + 1
        
    Next
    
    ' Store cards used.
    strCardsUsed = Left(strAvailableCards, iStart - 1)
    
    ' Remove the cards used.
    strAvailableCards = Right(strAvailableCards, Len(strAvailableCards) - iStart + 1)
    
    ' Remove any newline characters.
    If (Left(strAvailableCards, 1) = Chr(13)) Then
        strAvailableCards = Right(strAvailableCards, Len(strAvailableCards) - 1)
    End If
    If (Left(strAvailableCards, 1) = Chr(10)) Then
        strAvailableCards = Right(strAvailableCards, Len(strAvailableCards) - 1)
    End If
    
End Sub
Public Sub SaveMsgToFile(ByVal vDate As Date)

    ' Save message, key and ciphertext to file.
    '
    
    ' Create filename.
    
    ' Path and year.
    strFileName = App.Path & "\" & Year(vDate) & "."
    
    ' Add month to filename.
    If (Month(vDate) < 10) Then
        strFileName = strFileName & "0" & Month(vDate) & "."
    Else
        strFileName = strFileName & Month(vDate) & "."
    End If
    
    ' Add day to filename.
    If (Day(vDate) < 10) Then
        strFileName = strFileName & "0" & Day(vDate) & "."
    Else
        strFileName = strFileName & Day(vDate) & "."
    End If
    
    ' Add hour to filename.
    If (Hour(Time) < 10) Then
        strFileName = strFileName & "0" & Hour(Time) & "."
    Else
        strFileName = strFileName & Hour(Time) & "."
    End If
    
    ' Add minute to filename.
    If (Minute(Time) < 10) Then
        strFileName = strFileName & "0" & Minute(Time) & "."
    Else
        strFileName = strFileName & Minute(Time) & "."
    End If
    
    ' Add second to filename.
    If (Minute(Time) < 10) Then
        strFileName = strFileName & "0" & Second(Time)
    Else
        strFileName = strFileName & Second(Time)
    End If
    
    ' Final filename.
    strFileName = strFileName & " EM.txt"
    
    ' Set objects for file system.
    ' File system handles.
    Dim objFSO As Object
    Dim objFile As Object
    
    ' Create object as FSO.
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Open the file.
    ' 1. Filename.
    ' 2. iomode - ForReading (1), ForWriting (2) or ForAppending (8).
    ' 3. create - T/F.
    ' Set
    Set objFile = objFSO.CreateTextFile(strFileName, 2, True)
    
    ' Write line of data to file.
    objFile.Write (strFileName & Chr(13) & Chr(10))
    objFile.Write (Chr(13) & Chr(10))
    objFile.Write ("    Message: " & TextLength(strUserMessage) & strUserMessage & Chr(13) & Chr(10))
    objFile.Write (Chr(13) & Chr(10))
    objFile.Write ("  Plaintext: " & TextLength(strPlaintext) & strPlaintextSpaced & Chr(13) & Chr(10))
    objFile.Write ("        Key: " & TextLength(strKey) & strKeySpaced & Chr(13) & Chr(10))
    objFile.Write (" Ciphertext: " & TextLength(strCiphertext) & strCiphertextSpaced & Chr(13) & Chr(10))
    objFile.Write ("  Ct Padded: " & TextLength(strCiphertextPadded) & strCiphertextPaddedSpaced & Chr(13) & Chr(10))
    objFile.Write (Chr(13) & Chr(10))
    objFile.Write (" Cards used: " & strCardsUsed & Chr(13) & Chr(10))
    
    ' Close the file.
    objFile.Close
    
    ' Set objects to nothing.
    Set objFile = Nothing
    Set objFSO = Nothing
    
    
End Sub
Public Sub SaveCardUsedDataToFile(ByVal vText As String)

    ' Save message, key and ciphertext to a log file.
    '
    
    ' Create filename.
    strFileName = App.Path & "\CardsUsed.txt"
    
    ' File system handles.
    Dim objFSO As Object
    Dim objFile As Object
    
    ' Create object as FSO.
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Open the file.
    ' 1. Filename.
    ' 2. iomode - ForReading (1), ForWriting (2) or ForAppending (8).
    ' 3. create - T/F.
    
    ' Set
    If (objFSO.FileExists(strFileName)) Then
        Set objFile = objFSO.OpenTextFile(strFileName, 8, False)
    Else
        Set objFile = objFSO.CreateTextFile(strFileName, 2, True)
    End If
    
    ' Write line of data to file.
    objFile.Write (vText & Chr(13) & Chr(10))
    ' Close the file.
    objFile.Close
    
    ' Set objects to nothing.
    Set objFile = Nothing
    Set objFSO = Nothing
    
End Sub
Private Function TextLength(ByVal vTest As String) As String

    ' Take a text and get a formatted length of (000).
    ' The total length should be 5 characters.
    '
    
    ' Get length.
    TextLength = Len(vTest)
    
    ' Adjust length with leading zeroes.
    If (TextLength < 10) Then
        TextLength = "00" & TextLength
    ElseIf (TextLength < 100) Then
        TextLength = "0" & TextLength
    End If
    
    ' Add brackets.
    TextLength = "(" & TextLength & ") "
    
End Function

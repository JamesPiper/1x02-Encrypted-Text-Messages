VERSION 5.00
Begin VB.Form Message 
   Caption         =   "Message"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   17
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1000
   End
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   16
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   15
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1000
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   13
      Text            =   "2012/03/17"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewMessage 
      Caption         =   "New Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCommitData 
      Caption         =   "Commit Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdEncipher 
      Caption         =   "Encipher"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetKey 
      Caption         =   "Get Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtKey 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   10300
   End
   Begin VB.CommandButton cmdGetPlaintext 
      Caption         =   "Get Plaintext"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtCiphertext 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5520
      Width           =   10300
   End
   Begin VB.TextBox txtPlaintext 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   140
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   10300
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      MaxLength       =   140
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   10300
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   14
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Encryption Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Cipertext"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Plaintext"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form: Message.frm
' Author: James Piper
' Date: March 2012
'
' Description:
' Main form to collect message from user and encrypt based on key from
' card values.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim objDate As clsValidateDate
Dim constFirstValidDate
Dim constLastValidDate

Const constPlaintext = 0
Const constKey = 1
Const constCiphertext = 2


Private Sub Form_Load()

    ' Store temp message value.
    txtMessage.Text = "MESSAGES TO ARRIVE EACH DAY.  CHECK BACK TOMORROW."

    ' The date.
    Set objDate = New clsValidateDate
    txtDate.Text = "2011/08/01"
    objDate.DefaultValue = txtDate.Text

    ' Set the range of valid dates.
    constFirstValidDate = DateSerial(2011, 1, 1)
    constLastValidDate = DateSerial(2022, 12, 31)
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    ' Free resources.
    Set objDate = Nothing

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Based on KeyCode take action.
    Select Case KeyCode
    
    ' If someone has pressed 'page up' (33).
    Case 33
    
    ' If someone has pressed 'page down' (34).
    Case 34
    
    ' If someone has pressed the 'end key' (35).
    Case 35
        ' <ctrl> + <end> is KeyCode=35 and Shift = 2
        
    ' If someone has pressed the 'home key' (36).
    Case 36
        
    ' If someone has pressed 'cursor down' (40).
    Case 40
        ' Regular cursor key, no shift.
        ' Move through the textboxes on the form.
        
'        If (TypeOf Screen.ActiveControl Is TextBox) Then
'            ' Move from input field to input field, down the form.
'            Select Case Screen.ActiveControl.Name
'            Case "txtClient"
'                txtDescription.SetFocus
'            Case "txtDescription"
'                txtPreparedBy.SetFocus
'            Case "txtPreparedBy"
'                txtDate.SetFocus
'            End Select  ' Select Case Screen.ActiveControl.Name
'        Else
'        End If
        
    ' If someone has pressed 'cursor up' (38).
    Case 38
        ' Regular cursor key, no shift.
        ' Move through the textboxes on the form.
'        If (TypeOf Screen.ActiveControl Is TextBox) Then
'            ' Move from input field to input field, up the form.
'            Select Case Screen.ActiveControl.Name
'            Case "txtDescription"
'                txtClient.SetFocus
'            Case "txtPreparedBy"
'                txtDescription.SetFocus
'            End Select  ' Select Case Screen.ActiveControl.Name
'        Else
'        End If
        
    Case 68
        ' 68 is ASCII code for 'D'.
        
        ' If shift = 6, then <ctrl>+<alt>D was pressed.
        ' Execute code to save the datadump to a text file then run notepad with the data in it.
        If (Shift = 6) Then
            Call subDataDumpToFile
        End If ' If (Shift = 6) Then
        
        ' Control <D> to replace existing value with default value.
        ' If shift = 2, then <ctrl>+D was pressed.
        If (Shift = 2) Then
            Select Case Screen.ActiveControl.Name
                Case "txtDate"
                    txtDate.Text = objDate.DefaultValue
            End Select
        End If
        
    End Select

End Sub
Private Sub cmdNewMessage_Click()

    ' Clear out the text boxes.
    ' Reset buttons.
    '
    
    txtMessage.Text = ""
    txtPlaintext.Text = ""
    txtKey.Text = ""
    txtCiphertext.Text = ""
    
    ' Disable buttons.
    cmdCommitData.Enabled = False
    cmdEncipher.Enabled = False
    
    ' Spaced checkbox.
    chkSpaced(constPlaintext).Enabled = False
    chkSpaced(constKey).Enabled = False
    chkSpaced(constCiphertext).Enabled = False

End Sub
Private Sub cmdGetPlaintext_Click()

    ' Take user message and do the following.
    ' 1. Ignore all non-alpha characters.
    ' 2. All in lowercase.
    
    ' Test if message isn't blank.
    If (txtMessage.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter a message to encrypt."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Store user message.
    strUserMessage = txtMessage.Text
    
    ' Parse text.
    Call ParseMessage(strUserMessage)
    
    ' Spaced or compact.
    If (chkSpaced(constPlaintext).Value) Then
        txtPlaintext.Text = strPlaintextSpaced
    Else
        txtPlaintext.Text = strPlaintext
    End If
    
    ' Spaced checkbox.
    chkSpaced(constPlaintext).Enabled = True

End Sub
Private Sub cmdGetKey_Click()

    ' Load key for the length of the plaintext.
    '
    
    ' Test if message isn't blank.
    If (txtPlaintext.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter a message to encrypt."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Check if key is long enough.
    Call GetKey(strAvailableCards, Len(strPlaintext))
    If (Len(strKey) < Len(strPlaintext)) Then
        
        ' Message to user.
        strMsgBoxMessage = "The key is shorter than the message." + Chr(13) + Chr(10) + "Please enter more card data."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Spaced or compact.
    If (chkSpaced(constKey).Value) Then
        txtKey.Text = strKeySpaced
    Else
        txtKey.Text = strKey
    End If
    
    ' Enable encipher button.
    cmdEncipher.Enabled = True
    
    ' Spaced checkbox.
    chkSpaced(constKey).Enabled = True

End Sub
Private Sub chkSpaced_Click(Index As Integer)

    ' Switch the key between spaced and unspaced.
    '
    Select Case Index
        Case constPlaintext
            If (chkSpaced(constPlaintext).Value) Then
                txtPlaintext.Text = strPlaintextSpaced
            Else
                txtPlaintext.Text = strPlaintext
            End If
        Case constKey
            If (chkSpaced(constKey).Value) Then
                txtKey.Text = strKeySpaced
            Else
                txtKey.Text = strKey
            End If
        Case constCiphertext
            If (chkSpaced(constCiphertext).Value) Then
                txtCiphertext.Text = strCiphertextSpaced
            Else
                txtCiphertext.Text = strCiphertext
            End If
    End Select

End Sub
Private Sub cmdEncipher_Click()

    ' Encipher user's message with key.
    '
    
    ' Test if message isn't blank.
    If (txtPlaintext.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter a message to encrypt."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Show ciphertext.
    EncipherMessage (strPlaintext)
    ' Spaced or compact.
    If (chkSpaced(constCiphertext).Value) Then
        txtCiphertext.Text = strCiphertextSpaced
    Else
        txtCiphertext.Text = strCiphertext
    End If
    
    
    ' Enable button.
    cmdCommitData.Enabled = True
    
    ' Spaced checkbox.
    chkSpaced(constCiphertext).Enabled = True

End Sub
Private Sub cmdCommitData_Click()

    ' Save data to file and remove keys as used.
    '
    
    ' Remove used keys.
    RemoveUsedKeys
    
    ' Update form.
    Key.txtAvailableCards.Text = strAvailableCards
    
    ' Update file.
    Call SaveCardDataToFile(strAvailableCards, False)

    ' Save message data to a message file.
    Call SaveMsgToFile(txtDate.Text)
    
    ' Disable button.
    cmdCommitData.Enabled = False
    
    ' Disable encipher button.
    cmdEncipher.Enabled = False

End Sub
' txtDate Events
'
Private Sub txtDate_GotFocus()

    ' Got focus on the text box for Prepared On date.
    '
    
    ' Set value, if <esc> is pressed, we can return to this value.
    ' Also set the range of valid dates.
    objDate.GotFocus txtDate.Text, constFirstValidDate, constLastValidDate
    
    ' Position cursor on the month.
    txtDate.SelStart = 5

End Sub
Private Sub txtDate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Mouse down, don't allow any right click menu features.
    ' That doesn't seem to work.
    objDate.CurrentDateMouseDown Button, txtDate

End Sub
Private Sub txtDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Button=2 is mouse down, don't allow any functions.
    objDate.CurrentDateMouseUp Button, txtDate
    
End Sub
Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Key down on the text box.
    objDate.CurrentDateKeyDown KeyCode, Shift, txtDate

End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)

    ' Key pressed.
    ' Set the textbox where key press is filtered by the validate date module.
    txtDate.Text = objDate.CurrentDateKeyPress(KeyAscii, txtDate)
    ' Set the cursor position.
    txtDate.SelStart = objDate.CursorPostion

End Sub
Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)

    ' Key up on the text box.
    objDate.CurrentDateKeyUp KeyCode, txtDate

End Sub
Private Sub txtDate_Validate(Cancel As Boolean)
    
    ' If cancel is set to true, you can't lose focus.
    ' An invalid date means you can't lose focus.
    ' Either hit <esc> to put in original value or type in a valid date.
    '
    
    ' Is date valid?
    Cancel = objDate.DateNotValid(txtDate)
    
End Sub

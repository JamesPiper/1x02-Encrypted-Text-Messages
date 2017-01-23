VERSION 5.00
Begin VB.Form Decipher 
   Caption         =   "Decipher"
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
   Begin VB.TextBox txtPlaintextLen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Width           =   500
   End
   Begin VB.TextBox txtKeyLen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1800
      Width           =   500
   End
   Begin VB.TextBox txtCiphertextLen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   240
      Width           =   500
   End
   Begin VB.CheckBox chkPadding 
      Caption         =   "Padding"
      Height          =   195
      Left            =   5500
      TabIndex        =   2
      Top             =   270
      Width           =   1000
   End
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   6700
      TabIndex        =   6
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   6700
      TabIndex        =   5
      Top             =   240
      Width           =   1000
   End
   Begin VB.CheckBox chkSpaced 
      Caption         =   "Spaced"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   6700
      TabIndex        =   7
      Top             =   3480
      Width           =   1000
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
      TabIndex        =   8
      Top             =   240
      Width           =   2295
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
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
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
      TabIndex        =   4
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
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   10300
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
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   10300
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Length: "
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   17
      Top             =   3540
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Length: "
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   15
      Top             =   1860
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Length: "
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   13
      Top             =   300
      Width           =   705
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
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   2655
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Ciphertext"
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
Attribute VB_Name = "Decipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form: Decipher.frm
' Author: James Piper
' Date: March 2012
'
' Description:
' Form to decipher a message.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const constCiphertext = 0
Const constKey = 1
Const constPlaintext = 2


Private Sub Form_Load()

' Store temp values.
'txtCiphertext.Text = "OPIKU BXAVE QSLRU ZWWDG ZNMNT WSVYO FSLND ULWDJ XXXXX "
'txtCiphertext.Text = "OSPNI PKHUH BUXAA TVVES QBSKL WRUUJ ZZWWW BDYGJ ZNNTM BNPTM WHSQV QYGOH FVSVL PNZDX UFLSW ZDGJN XXXXX XXXXX"
'txtKey.Text = "CLQSU VTICQ QBUJZ VSWBZ WNOLM SQLXO DISZR GUFPN AAAAA "
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    '
    
End Sub
Private Sub cmdNewMessage_Click()

    ' Reset variable values.
    ' Clear out the text boxes.
    ' Reset buttons.
    '
    
    ' Reset variables.
    strCiphertextEntered = ""
    strCiphertextEnteredSpaced = ""
    strKeyEntered = ""
    strKeyEnteredSpaced = ""
    strPlaintextDeciphered = ""
    strPlaintextDecipheredSpaced = ""
    
    ' Clear input boxes.
    txtPlaintext.Text = ""
    txtKey.Text = ""
    txtCiphertext.Text = ""
    
    ' Padding checkbox.
    chkPadding.Enabled = True
    
    ' Spaced checkbox.
    chkSpaced(constPlaintext).Enabled = False
    chkSpaced(constPlaintext).Value = False
    chkSpaced(constKey).Enabled = False
    chkSpaced(constKey).Value = False
    chkSpaced(constCiphertext).Enabled = False
    chkSpaced(constCiphertext).Value = False

    ' Enable button.
    cmdGetPlaintext.Enabled = True

    ' Reset lengths.
    txtCiphertextLen.Text = ""
    txtKeyLen = ""
    txtPlaintextLen = ""
    
End Sub
Private Sub cmdGetPlaintext_Click()

    ' Take ciphertext and key to decipher the plaintext.
    '
    
    ' Test if ciphertext isn't blank.
    If (txtCiphertext.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter a message to decrypt."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Test if key isn't blank.
    If (txtKey.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter a key."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Parse text.
    Call ParseCiphertext(txtCiphertext.Text)
    Call ParseKey(txtKey.Text)
    
    ' Remove padding, if any.
    If (chkPadding.Value) Then
        strCiphertextEntered = RemovePadding(strCiphertextEntered)
    End If
    
    ' Make sure key is as long as ciphertext.
    If (Len(strKeyEntered) < Len(strCiphertextEntered)) Then
        
        ' Message to user.
        strMsgBoxMessage = "The key is shorter than the message." + Chr(13) + Chr(10) + "Please enter a longer key."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    ' Add space for every five characters.
    strCiphertextEnteredSpaced = AddSpacing(strCiphertextEntered)
    
    ' Show parsed user data.
    txtCiphertext.Text = strCiphertextEntered
    txtKey.Text = strKeyEntered
    
    ' Decipher text.
    Call DecipherMessage(strCiphertextEntered, strKeyEntered)
    
    ' Spaced or compact.
    If (chkSpaced(constPlaintext).Value) Then
        txtPlaintext.Text = strPlaintextDecipheredSpaced
    Else
        txtPlaintext.Text = strPlaintextDeciphered
    End If
    
    ' Spaced checkbox.
    chkSpaced(constCiphertext).Enabled = True
    chkSpaced(constPlaintext).Enabled = True
    chkSpaced(constKey).Enabled = True

    ' Padding checkbox.
    chkPadding.Enabled = False
    
    ' Disable button.
    cmdGetPlaintext.Enabled = False

    ' Show lengths.
    txtCiphertextLen.Text = Len(txtCiphertext.Text)
    txtKeyLen = Len(txtKey.Text)
    txtPlaintextLen = Len(txtPlaintext.Text)
    
End Sub
Private Sub chkSpaced_Click(Index As Integer)

    ' Switch the key between spaced and unspaced text.
    '
    
    Select Case Index
        Case constCiphertext
            If (chkSpaced(constCiphertext).Value) Then
                txtCiphertext.Text = strCiphertextEnteredSpaced
            Else
                txtCiphertext.Text = strCiphertextEntered
            End If
            ' Show length.
            txtCiphertextLen.Text = Len(txtCiphertext.Text)
        Case constKey
            If (chkSpaced(constKey).Value) Then
                txtKey.Text = strKeyEnteredSpaced
            Else
                txtKey.Text = strKeyEntered
            End If
            ' Show length.
            txtKeyLen = Len(txtKey.Text)
        Case constPlaintext
            If (chkSpaced(constPlaintext).Value) Then
                txtPlaintext.Text = strPlaintextDecipheredSpaced
            Else
                txtPlaintext.Text = strPlaintextDeciphered
            End If
            ' Show length.
            txtPlaintextLen = Len(txtPlaintext.Text)
    End Select

End Sub

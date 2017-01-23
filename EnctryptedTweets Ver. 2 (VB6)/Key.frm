VERSION 5.00
Begin VB.Form Key 
   Caption         =   "Generate Key"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtParsedCards 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3000
      Width           =   10300
   End
   Begin VB.CommandButton cmdSaveCards 
      Caption         =   "Save Card Data"
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
      Top             =   4050
      Width           =   2295
   End
   Begin VB.TextBox txtAvailableCards 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   120
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4500
      Width           =   10300
   End
   Begin VB.CommandButton cmdValidateCardData 
      Caption         =   "Validate Card Data"
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
      TabIndex        =   2
      Top             =   2550
      Width           =   2295
   End
   Begin VB.TextBox txtCards 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1500
      Width           =   10300
   End
   Begin VB.Label Label1 
      Caption         =   "Available Card Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   10200
   End
   Begin VB.Label Label1 
      Caption         =   "BA = Black ace, R6, Six of hearts or diamonds."
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   900
      Width           =   10200
   End
   Begin VB.Label Label1 
      Caption         =   "Lowercase or uppercase. With or without spaces."
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   10200
   End
   Begin VB.Label Label1 
      Caption         =   "R for red, B for Black. 2 through 9, T(10), J, Q, K. A"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   10200
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Random Card Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   10200
   End
End
Attribute VB_Name = "Key"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form: Key.frm
' Author: James Piper
' Date: March 2012
'
' Description:
' User enters random card data to create encryption keys.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




' 2011.08.01
' R3 RQ B4 B6 B8 B9 B7 R9 R3 B4 B4 R2 B8 R10 BK B9 B6 B10 R2 BK B10 BA B2 RQ RK B6 B4 RQ BJ B2 R4 R9 B6 BK B5 R7 B8 R6 B3 BA
' MESSAGES TO ARRIVE EACH DAY.  CHECK BACK TOMORROW.
' MESSAGESTOARRIVEEACHDAYCHECKBACKTOMORROWXXXXX
' CLQSU VTICQ QBUJZ VSWBZ WNOLM SQLXO DISZR GUFPN AAAAA
' OPIKU BXAVE QSLRU ZWWDG ZNMNT WSVYO FSLND ULWDJ XXXXX

' 2011.08.02
' B9 R3 BQ R5 R4 B3 B3 R6 B2 RQ R2 B7 B9 BK B6 R10 B6 BK R5 R6 R2 RA R2 B2 RQ R8 R5 R10 B10 B9 B8 B8 B3 R5 BA RA B10
' STAGE ONE.  CAMEL TO MOVE.  ADD PINK TO GUTBUCKET.
' STAGEONECAMELTOMOVEADDPINKTOGUTBUCKETXXX
' VCYED PPFOL BTVZS JSZEF BABOL HEJWV UUPEN AWAAA
' NVYKH DCJQL NXGSG VGUIF EDQWY RXXCP NVJGX EPXXX
'

Private Sub Form_Load()
    
    ' Load sample card values.
    txtCards.Text = "" ' "R3 RQ B4 B6 B8 B9 B7 R9 R3 B4 B4 R2 B8 R10 BK B9 B6 B10 R2 BK B10 BA B2 RQ RK B6 B4 RQ BJ B2 R4 R9 B6 BK B5 R7 B8 R6 B3 BA"
    
    ' Load available card data.
    txtAvailableCards.Font = "Courier new"
    txtAvailableCards.FontSize = 11
    strAvailableCards = ""
    Call LoadCardDataFromFile
    txtAvailableCards.Text = strAvailableCards
    
End Sub
Private Sub cmdValidateCardData_Click()

    ' Take user card data and parse
    '
    
    ' Test if message isn't blank.
    If (txtCards.Text = "") Then
        
        ' Message to user.
        strMsgBoxMessage = "Please enter card data to create a key."
        intAnswer = MsgBox(strMsgBoxMessage, vbCritical, constTitle)
        
        ' Don't continue.
        Exit Sub
    End If
    
    txtParsedCards.Text = ParseCardData(txtCards.Text)
    
End Sub
Private Sub cmdSaveCards_Click()

    ' Add new data input from user to key card file.
    '

    ' Only add data if it's there.
    If (txtParsedCards.Text = "") Then
        Exit Sub
    End If
    
    ' Added card data to the file.
    ' True means keep existing data in the file.
    Call SaveCardDataToFile(txtParsedCards.Text, True)

    ' Update available card data.
    strAvailableCards = ""
    Call LoadCardDataFromFile
    txtAvailableCards.Text = strAvailableCards
    
    ' Remove card data from input fields.
    txtCards.Text = ""
    txtParsedCards.Text = ""
    
End Sub




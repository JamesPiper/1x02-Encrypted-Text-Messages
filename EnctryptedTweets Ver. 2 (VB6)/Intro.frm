VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Intro 
   Caption         =   "Introduction & Explanation"
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
   Begin RichTextLib.RichTextBox rtIntro 
      Height          =   4500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   7938
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Intro.frx":0000
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form: Intro.frm
' Author: James Piper
' Date: March 2012
'
' Description:
' Provide info to user on what this program does.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Form_Load()

    ' Set text.
    rtIntro.FileName = App.Path & "\Encrypted Message Intro Page.rtf"

End Sub
Private Sub Form_Resize()

    ' Fill richtextbox to parent form.
    '
    
    rtIntro.Width = Intro.Width
    rtIntro.Height = Intro.Height
    
End Sub

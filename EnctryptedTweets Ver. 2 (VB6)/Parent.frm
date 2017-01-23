VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Parent 
   BackColor       =   &H8000000C&
   Caption         =   "Encrypt Message"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10620
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   630
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8000
         _ExtentX        =   14102
         _ExtentY        =   1111
         ButtonWidth     =   1429
         ButtonHeight    =   1005
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Intro"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Encipher"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Key"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Decipher"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form: Parent.frm
' Author: James Piper
' Date: March 2012
'
' Description:
' Parent form to hold rest of the program.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub MDIForm_Load()

    ' Load parent form.
    '
    
    ' Load decipher form.
    Load Decipher
    Decipher.Show
    
    ' Load key form.
    Load Key
    Key.Show

    ' Load message form.
    Load Encipher
    Encipher.Show
    
    ' Load Intro form.
    Load Intro
    Intro.Show
    
End Sub
Private Sub MDIForm_Resize()

    ' Change width of toolbar.
    ' To equal width of parent form.
    Toolbar.Width = Me.Width
    
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    ' User selected form.
    '
    Select Case Button
    Case "Intro"
        Intro.Show
        Key.Hide
        Encipher.Hide
        Decipher.Hide
    Case "Encipher"
        Encipher.Show
        Key.Hide
        Decipher.Hide
        Intro.Hide
    Case "Key"
        Key.Show
        Encipher.Hide
        Decipher.Hide
        Intro.Hide
    Case "Decipher"
        Decipher.Show
        Key.Hide
        Encipher.Hide
        Intro.Hide
    End Select
    
End Sub

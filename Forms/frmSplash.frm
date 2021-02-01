VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   1
         Top             =   2700
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   3870
         Left            =   60
         Picture         =   "frmSplash.frx":000C
         Top             =   120
         Width           =   6990
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    MoveMe
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    'lblCopyright.Caption = App.LegalCopyright
    'lblCompany.Caption = App.CompanyName
    Timer1.Enabled = True
End Sub

Sub MoveMe() 'owner As Form)
    Dim L As Single: L = Screen.Width / 2 - Me.Width / 2
    Dim T As Single: T = Screen.Height / 2 - Me.Height / 2
    Me.Move L, T
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub

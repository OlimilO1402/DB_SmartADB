VERSION 5.00
Begin VB.Form FrmEditor 
   Caption         =   "Editor"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Übernehmen"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "FrmEditor.frx":0000
      Top             =   600
      Width           =   9975
   End
End
Attribute VB_Name = "FrmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbr As VbMsgBoxResult
Dim Content As String

Public Function ShowDialog(owner As Form, sContent As String) As VbMsgBoxResult
    mbr = 0
    Content = sContent
    Text1.Text = sContent
    MoveMe owner
    Me.Show vbModal, owner
    If mbr = vbOK Then sContent = Content 'Text1.Text 'Content
    ShowDialog = mbr
End Function

Sub MoveMe(owner As Form)
    Dim l As Single: l = owner.Left + owner.Width / 2 - Me.Width / 2
    Dim T As Single: T = owner.Top + owner.Height / 2 - Me.Height / 2
    Me.Move l, T
End Sub

Private Sub BtnCancel_Click()
    mbr = vbCancel
    Unload Me
End Sub

Private Sub BtnSave_Click()
    mbr = vbOK
    Content = Text1.Text
    Unload Me
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 0 '8 * Screen.TwipsPerPixelX
    Dim l As Single: l = Text1.Left
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - l - brdr
    Dim H As Single: H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
End Sub

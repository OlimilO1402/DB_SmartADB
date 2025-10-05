VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6285
      Begin VB.PictureBox PnlDBNav2 
         BorderStyle     =   0  'Kein
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   2
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private WithEvents mDN As DataNavigator
Attribute mDN.VB_VarHelpID = -1

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Set mDN = New_DataNavigator(Me, PnlDBNav2, "mDN", 1, 3, 10)
  
End Sub

Private Sub mDN_Click(Sender As Object, Button As EDataNavButton)
 If Sender Is mDN Then
   Select Case Button
   Case nbFirst:  Label1.Caption = "nbFirst"
   Case nbPrior:  Label1.Caption = "nbPrior"
   Case nbNext:   Label1.Caption = "nbNext"
   Case nbLast:   Label1.Caption = "nbLast"
   Case nbAddNew: Label1.Caption = "nbAddNew"
   Case nbInsNew: Label1.Caption = "nbInsNew"
   Case nbDelete: Label1.Caption = "nbDelete"
   Case nbEdit:   Label1.Caption = "nbEdit"
   Case nbCancel: Label1.Caption = "nbCancel"
   Case nbUpdate: Label1.Caption = "nbPost"
   End Select
 End If
End Sub


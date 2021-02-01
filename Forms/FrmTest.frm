VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Test"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text8 
      Height          =   345
      Left            =   1200
      TabIndex        =   15
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1200
      TabIndex        =   13
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "StrVal"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "DatVal"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DblVal"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "SngVal"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CurVal"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "LngVal"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IntVal"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BytVal"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbr As VbMsgBoxResult
Private m_obj As Test

Public Function ShowDialog(owner As Form, aObj As Test) As VbMsgBoxResult
    Set m_obj = aObj
    UpdateView m_obj
    MoveMe owner
    Me.Show vbModal, FMain
    ShowDialog = mbr
End Function

Sub MoveMe(owner As Form)
    Dim L As Single: L = owner.Left + owner.Width / 2 - Me.Width / 2
    Dim T As Single: T = owner.Top + owner.Height / 2 - Me.Height / 2
    Me.Move L, T
End Sub

Private Sub BtnCancel_Click()
    mbr = vbCancel
    Set m_obj = Nothing
    Unload Me
End Sub

Private Sub BtnOK_Click()
    mbr = vbOK
    UpdateData m_obj
    Unload Me
End Sub

Sub UpdateData(obj As Test)
    With obj
        .BytVal = ParseParam(.BytVal, Me.Text1.Text, obj, "BytVal")
        .IntVal = ParseParam(.IntVal, Me.Text2.Text, obj, "IntVal")
        .LngVal = ParseParam(.LngVal, Me.Text3.Text, obj, "LngVal")
        .CurVal = ParseParam(.CurVal, Me.Text4.Text, obj, "CurVal")
        .SngVal = ParseParam(.SngVal, Me.Text5.Text, obj, "SngVal")
        .DblVal = ParseParam(.DblVal, Me.Text6.Text, obj, "DblVal")
        .DatVal = ParseParam(.DatVal, Me.Text7.Text, obj, "DatVal")
        .StrVal = ParseParam(.StrVal, Me.Text8.Text, obj, "StrVal")
    End With
End Sub
Sub UpdateView(obj As Test)
    With obj
        Me.Text1.Text = .BytVal
        Me.Text2.Text = .IntVal
        Me.Text3.Text = .LngVal
        Me.Text4.Text = .CurVal
        Me.Text5.Text = .SngVal
        Me.Text6.Text = .DblVal
        Me.Text7.Text = .DatVal
        Me.Text8.Text = .StrVal
    End With
End Sub


VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Form2"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mDefAdd As IAdd
Public mPOfAdd As IAdd 'Parent Of Add '
Public mNamed As INamed

Private Sub Form_Load()
Dim iNd As INamed
  Text1.Text = vbNullString
  Set iNd = mDefAdd
  Call mPOfAdd.ToListBox(List1)
  SelectListBoxItemByName (iNd.Name)
End Sub

Private Sub SelectListBoxItemByName(StrName As String)
Dim i As Long, StrV As String
  For i = 0 To List1.ListCount - 1
    StrV = List1.List(i)
    If StrName = StrV Then
      List1.ListIndex = i
      Exit Sub
    End If
  Next
End Sub
Private Sub BtnOK_Click()
Dim StrName As String: StrName = Text1.Text
  If Len(StrName) = 0 Then
    Call BtnCancel_Click
    Exit Sub
  End If
  If Not mNamed Is Nothing Then mNamed.Name = StrName
  If List1.ListIndex > -1 Then
    Set mDefAdd = mPOfAdd.GetByName(List1.List(List1.ListIndex))
  End If
  If Not mDefAdd Is Nothing Then
    Set mNamed = mDefAdd.Add(mNamed, True)
    If mNamed Is Nothing Then MsgBox ("Object mit dem Namen: " & StrName & " schon in der Liste vorhanden")
  Else
    MsgBox "mDefAdd is nothing"
  End If
  'End If
  Unload Me
End Sub
Private Sub BtnCancel_Click()
  'Form1.mForm2OKFlag = vbCancel
  Unload Me
End Sub


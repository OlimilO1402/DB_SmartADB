VERSION 5.00
Begin VB.Form FrmCity 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Stadt"
   ClientHeight    =   3135
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
   ScaleHeight     =   3135
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbCountry 
      Height          =   345
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox TxtVorwahl 
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox TxtPLZ 
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox TxtName2 
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox TxtName 
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Land"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vorwahl"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PLZ"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zusatz"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   630
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "FrmCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbr   As VbMsgBoxResult
Dim m_obj   As City
Dim m_Countries As List

Public Function ShowDialog(owner As Form, aObj As City, aCountries As List) As VbMsgBoxResult
    Set m_obj = aObj
    Set m_Countries = aCountries
    m_Countries.ToListbox CmbCountry, True
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
    Unload Me
End Sub

Private Sub BtnOK_Click()
    mbr = vbOK
    UpdateData m_obj
    Unload Me
End Sub

Sub UpdateData(obj As City)
    With obj
        .Name = Me.TxtName.Text
        .Nam2 = Me.TxtName2.Text
        .PLZ = Me.TxtPLZ.Text
        .Vorwahl = Me.TxtVorwahl.Text
        If Len(CmbCountry.Text) Then
            Dim i As Long: i = CmbCountry.ListIndex
            If i >= 0 Then
                Set .Country = m_Countries.Item(i)
                If Not .Country.Cities.Contains(obj) Then
                    .Country.Cities.Add obj
                End If
            End If
        End If
    End With
End Sub

Sub UpdateView(obj As City)
    With obj
        Me.TxtName.Text = .Name
        Me.TxtName2.Text = .Nam2
        Me.TxtPLZ.Text = .PLZ
        Me.TxtVorwahl.Text = .Vorwahl
        If .Country Is Nothing Then Exit Sub
        Me.CmbCountry.Text = .Country.Key
    End With
End Sub


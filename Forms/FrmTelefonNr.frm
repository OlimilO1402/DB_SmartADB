VERSION 5.00
Begin VB.Form FrmTelefonNr 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "TelefonNr"
   ClientHeight    =   2175
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbCity 
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox CmbCountry 
      Height          =   345
      ItemData        =   "FrmTelefonNr.frx":0000
      Left            =   1200
      List            =   "FrmTelefonNr.frx":0002
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox TxtNumber 
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label LblCity 
      AutoSize        =   -1  'True
      Caption         =   "Stadt"
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Land"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Nummer"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   630
   End
End
Attribute VB_Name = "FrmTelefonNr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbr As VbMsgBoxResult
Private m_obj As TelefonNr
Private Countries As List

Public Function ShowDialog(owner As Form, aObj As TelefonNr, aCountries As List) As VbMsgBoxResult
    Set m_obj = aObj
    Set Countries = aCountries
    Countries.ToListbox CmbCountry, True
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

Private Sub CmbCountry_Click()
    Dim i As Long: i = CmbCountry.ListIndex
    Dim c As Country: Set c = Countries.Item(i)
    If c Is Nothing Then Exit Sub
    CmbCity.Text = ""
    c.Cities.ToListbox CmbCity, True
End Sub

Sub UpdateData(obj As TelefonNr)
    With obj
        .Number = Me.TxtNumber.Text
        Dim i As Long: i = CmbCountry.ListIndex
        If i >= 0 Then
            Dim c As Country: Set c = Countries.Item(i)
            i = CmbCity.ListIndex
            If i >= 0 Then
                Set .City = c.Cities.Item(i)
            End If
        End If
    End With
End Sub
Sub UpdateView(obj As TelefonNr)
    With obj
        Me.TxtNumber.Text = .Number
        Dim c As City: Set c = .City
        If c Is Nothing Then Exit Sub
        c.Country.Cities.ToListbox CmbCity ', True
        Me.CmbCity.Text = c.Key
        Me.CmbCountry.Text = c.Country.Key
    End With
End Sub


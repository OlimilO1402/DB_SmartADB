VERSION 5.00
Begin VB.Form FrmAddress 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Adresse"
   ClientHeight    =   2655
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
   ScaleHeight     =   2655
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbCity 
      Height          =   345
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox TxtInfo 
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox TxtHNr 
      Height          =   345
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox TxtStreet 
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LblCity 
      AutoSize        =   -1  'True
      Caption         =   "Stadt"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Info"
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label LblHNr 
      AutoSize        =   -1  'True
      Caption         =   "HNr"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   315
   End
   Begin VB.Label LblStreet 
      AutoSize        =   -1  'True
      Caption         =   "Straße"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "FrmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbr As VbMsgBoxResult
Private m_obj As Address
Private m_Cities As List

Public Function ShowDialog(owner As Form, aObj As Address, aCities As List) As VbMsgBoxResult
    Set m_obj = aObj
    Set m_Cities = aCities
    m_Cities.ToListbox CmbCity, True
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

Sub UpdateData(obj As Address)
    If obj Is Nothing Then Exit Sub
    With obj
        .Street = Me.TxtStreet.Text
        .HNr = Me.TxtHNr.Text
        .Info = Me.TxtInfo.Text
        If Len(CmbCity.Text) Then
            Dim i As Long: i = CmbCity.ListIndex
            If i >= 0 Then
                Set .City = m_Cities.Item(i)
                If Not .City.Addresses.Contains(obj) Then
                    .City.Addresses.Add obj
                End If
            End If
        End If
    End With
End Sub

Sub UpdateView(obj As Address)
    With obj
        Me.TxtStreet.Text = .Street
        Me.TxtHNr.Text = .HNr
        Me.TxtInfo.Text = .Info
        If .City Is Nothing Then Exit Sub
        Me.CmbCity.Text = .City.Key
    End With
End Sub


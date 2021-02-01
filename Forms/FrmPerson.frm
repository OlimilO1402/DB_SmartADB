VERSION 5.00
Begin VB.Form FrmPerson 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Person"
   ClientHeight    =   5055
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
   ScaleHeight     =   5055
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbPersonGender 
      Height          =   345
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
   End
   Begin VB.ComboBox CmbTelefonNr 
      Height          =   345
      Left            =   1200
      TabIndex        =   17
      Top             =   3960
      Width           =   3735
   End
   Begin VB.ComboBox CmbAddress 
      Height          =   345
      Left            =   1200
      TabIndex        =   15
      Top             =   3480
      Width           =   3735
   End
   Begin VB.ComboBox CmbFather 
      Height          =   345
      Left            =   1200
      TabIndex        =   13
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox TxtBirthD 
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox TxtFamName 
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
   Begin VB.ComboBox CmbMother 
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Top             =   2520
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
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Gender"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Telefon"
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Adresse"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vater"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Geburtsd."
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fam.Name"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name2"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mutter"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   2520
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
Attribute VB_Name = "FrmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbr As VbMsgBoxResult
Private m_obj As Person
Private m_Persons   As List
Private m_Addresses As List
Private m_TelNrs    As List

Public Function ShowDialog(owner As Form, aObj As Person, aPersons As List, aAddresses As List, aTelNrs As List) As VbMsgBoxResult
    Set m_obj = aObj
    Set m_Persons = aPersons
    Set m_Addresses = aAddresses
    Set m_TelNrs = aTelNrs
    
    m_Persons.ToListbox Me.CmbMother, True
    m_Persons.ToListbox Me.CmbFather, True  'father and mother cannot be the same ;)
    m_Addresses.ToListbox Me.CmbAddress
    m_TelNrs.ToListbox Me.CmbTelefonNr
    EGender_ToListBox Me.CmbPersonGender
    UpdateView m_obj
    MoveMe owner
    Me.Show vbModal, owner
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

Sub UpdateData(obj As Person)
    With obj
        .PreName1 = Me.TxtName.Text
        .PreName2 = Me.TxtName2.Text
        .FamName = Me.TxtFamName.Text
        .BirthD = ParseParam(.BirthD, Me.TxtBirthD.Text, obj, "BirthD")
        .Gender = EGender_Parse(CmbPersonGender.Text)
        
        Dim i As Long
        i = CmbMother.ListIndex
        'halt nur löschen wenn der Text leer ist
        If i >= 0 Then Set .Mother = m_Persons.Item(i) Else If CmbMother.Text = "" Then Set .Mother = Nothing
        
        i = CmbFather.ListIndex
        If i >= 0 Then Set .Father = m_Persons.Item(i) Else If CmbFather.Text = "" Then Set .Father = Nothing
        i = CmbAddress.ListIndex
        If i >= 0 Then Set .Address = m_Addresses.Item(i) Else If CmbAddress.Text = "" Then Set .Address = Nothing
        i = CmbTelefonNr.ListIndex
        If i >= 0 Then Set .TelNumber = m_TelNrs.Item(i) Else If CmbTelefonNr.Text = "" Then Set .TelNumber = Nothing
    End With
End Sub
Sub UpdateView(obj As Person)
    With obj
        Me.TxtName.Text = .PreName1
        Me.TxtName2.Text = .PreName2
        Me.TxtFamName.Text = .FamName
        Me.TxtBirthD.Text = IIf(.BirthD = 0, "", Format(.BirthD, "dd.mmm.yyyy"))
        Me.CmbPersonGender = EGender_ToStr(.Gender)
        
        Dim p As Person
        Set p = .Mother
        If Not p Is Nothing Then Me.CmbMother.Text = p.Key Else Me.CmbMother.Text = ""
        Set p = .Father
        If Not p Is Nothing Then Me.CmbFather.Text = p.Key Else Me.CmbFather.Text = ""
        
        Dim a As Address: Set a = .Address
        If Not a Is Nothing Then Me.CmbAddress.Text = a.Key Else Me.CmbAddress.Text = ""
        
        Dim T As TelefonNr: Set T = .TelNumber
        If Not T Is Nothing Then Me.CmbTelefonNr.Text = T.Key Else Me.CmbTelefonNr.Text = ""
    End With
End Sub


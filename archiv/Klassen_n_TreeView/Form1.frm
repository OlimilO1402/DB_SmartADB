VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuLand 
      Caption         =   "Land"
      Begin VB.Menu mnuLandExample 
         Caption         =   "Example"
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuAddLand 
         Caption         =   "Land"
      End
      Begin VB.Menu mnuAddCity 
         Caption         =   "City"
      End
      Begin VB.Menu mnuAddStreet 
         Caption         =   "Street"
      End
      Begin VB.Menu mnuAddHouse 
         Caption         =   "House"
      End
      Begin VB.Menu mnuAddFamily 
         Caption         =   "Family"
      End
      Begin VB.Menu mnuAddPerson 
         Caption         =   "Person"
      End
   End
   Begin VB.Menu mnuLandExplorer 
      Caption         =   "LandExplorer"
      Begin VB.Menu mnuOpenAllNodes 
         Caption         =   "Open All Nodes"
      End
      Begin VB.Menu mnuSaveOpenNodes 
         Caption         =   "Save Open Nodes"
      End
      Begin VB.Menu mnuRestoreOpenNodes 
         Caption         =   "Restore All Open Nodes"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'I am the World
'Private Lands As Collection 'holds objects of typ Land
'Private mLand As Land
Public mForm2OKFlag As Integer
Private mLandExplorer As LandExplorer

Private Sub Form_Load()
  'ClearLands
  'Set mLandExplorer = New_LandExplorer(TreeView1, StatusBar1)
End Sub
Private Sub ClearLands()
  'Set Lands = New Collection
End Sub
Private Sub Form_Resize()
Dim L As Single, T As Single, W As Single, H As Single
Dim Brdr As Single, STPX As Single, STPY As Single
  STPX = Screen.TwipsPerPixelX
  STPY = Screen.TwipsPerPixelY
  Brdr = 0 * STPX
  L = 1 * Brdr
  T = 1 * Brdr 'TreeView1.Top
  W = Me.ScaleWidth - L - 1 * Brdr
  H = Me.ScaleHeight - T - 1 * Brdr - StatusBar1.Height
  TreeView1.Move L, T, W, H
End Sub
Private Function Assert(mObj As Object, Optional mess As String) As Boolean
  If Not mObj Is Nothing Then Exit Function
  If Len(mess) > 0 Then MsgBox mess
  Assert = True
End Function



Private Sub mnuAddLand_Click()
  If Not Assert(mLandExplorer) Then
    If MsgBox("This will delete the current land, proceed anyway?", vbOKCancel) = vbCancel Then Exit Sub
  End If
Dim aLand As Land, StrName As String
  StrName = InputBox("Geben Sie den Namen des neuen Landes ein: ", "Name des neuen Landes", "Deutschland")
  If Len(StrName) = 0 Then Exit Sub
  Set aLand = New_Land(StrName)
  Set mLandExplorer = New_LandExplorer(TreeView1, StatusBar1, aLand)
  Call aLand.ToTreeView(TreeView1, aLand.Name)
End Sub
Private Sub mnuAddCity_Click()
  If Assert(mLandExplorer, "Add a land first, or select example.") Then Exit Sub
  'Form2 ist hier noch überflüssig
  'Set Form2.mAdd = aLand: Set Form2.mNamed = aCity
Dim StrName As String
  StrName = InputBox("Geben Sie den Namen der Stadt an.", "Name der neuen Stadt")
  If Len(StrName) > 0 Then
    Dim aLand As Land: Set aLand = mLandExplorer.GetCurLand
    If Not aLand Is Nothing Then
      If aLand.AddCity(New_City(StrName), True) Is Nothing Then
        MsgBox "Die Stadt: '" & StrName & "' ist schon vorhanden."
        Exit Sub
      End If
    Else
      MsgBox "ungültiges Object 'Land'"
    End If
  End If
  mLandExplorer.SaveOpenNodes
  Call aLand.ToTreeView(TreeView1, aLand.Name)
  Call mLandExplorer.RestoreOpenNodes
End Sub
Private Sub mnuAddStreet_Click()
  If Assert(mLandExplorer, "Add a land first, or select example.") Then Exit Sub
  mLandExplorer.SaveOpenNodes
  'jetzt Form2
  Dim aLand As Land: Set aLand = mLandExplorer.GetCurLand
  Dim aCity As City: Set aCity = mLandExplorer.GetCurCity
  Dim aStreet As New Street
  Dim StrName As String
  Set Form2.mPOfAdd = aLand
  Set Form2.mDefAdd = aCity
  Set Form2.mNamed = aStreet
  Form2.Caption = "New Street"
  Form2.Label1.Caption = "New Street:"
  Form2.Frame1.Caption = "Select City:"
  Form2.Show 1, Me
  Call aLand.ToTreeView(TreeView1, aLand.Name)
  Call mLandExplorer.RestoreOpenNodes
End Sub
Private Sub mnuAddHouse_Click()
  If Assert(mLandExplorer, "Add a land first, or select example.") Then Exit Sub
  mLandExplorer.SaveOpenNodes
  'jetzt Form2
  Dim aLand As Land: Set aLand = mLandExplorer.GetCurLand
  Dim aCity As City: Set aCity = mLandExplorer.GetCurCity
  Dim aStreet As Street: Set aStreet = mLandExplorer.GetCurStreet
  Dim aHouse As New House
  Dim StrName As String
  Set Form2.mPOfAdd = aCity
  Set Form2.mDefAdd = aStreet
  Set Form2.mNamed = aHouse
  Form2.Caption = "New House"
  Form2.Label1.Caption = "New House:"
  Form2.Frame1.Caption = "Select Street:"
  Form2.Show 1, Me
  Call aLand.ToTreeView(TreeView1, aLand.Name)
  Call mLandExplorer.RestoreOpenNodes
End Sub
Private Sub mnuAddFamily_Click()
  If Assert(mLandExplorer, "Add a land first, or select example.") Then Exit Sub
  mLandExplorer.SaveOpenNodes
  'jetzt Form2
  Dim aLand As Land: Set aLand = mLandExplorer.GetCurLand
  Dim aStreet As Street: Set aStreet = mLandExplorer.GetCurStreet
  Dim aHouse As House: Set aHouse = mLandExplorer.GetCurHouse
  Dim aFamily As New Family
  Dim StrName As String
  Set Form2.mPOfAdd = aStreet
  Set Form2.mDefAdd = aHouse
  Set Form2.mNamed = aFamily
  Form2.Caption = "New Family"
  Form2.Label1.Caption = "New Family:"
  Form2.Frame1.Caption = "Select House:"
  Form2.Show 1, Me
  Call aLand.ToTreeView(TreeView1, aLand.Name)
  Call mLandExplorer.RestoreOpenNodes
End Sub
Private Sub mnuAddPerson_Click()
  If Assert(mLandExplorer, "Add a land first, or select example.") Then Exit Sub
  mLandExplorer.SaveOpenNodes
  'jetzt Form2
  Dim aLand As Land: Set aLand = mLandExplorer.GetCurLand
  Dim aHouse As House: Set aHouse = mLandExplorer.GetCurHouse
  Dim aFamily As Family: Set aFamily = mLandExplorer.GetCurFamily
  Dim aPerson As New Person
  Dim StrName As String
  Set Form2.mPOfAdd = aHouse
  Set Form2.mDefAdd = aFamily
  Set Form2.mNamed = aPerson
  Form2.Caption = "New Person"
  Form2.Label1.Caption = "New Person:"
  Form2.Frame1.Caption = "Select Family:"
  Form2.Show 1, Me
  Call aLand.ToTreeView(TreeView1, aLand.Name)
  Call mLandExplorer.RestoreOpenNodes
End Sub

Private Sub mnuLandExample_Click()
  If Not Assert(mLandExplorer) Then
    If MsgBox("This will delete the current land, proceed anyway?", vbOKCancel) = vbCancel Then Exit Sub
  End If
Dim aLand As Land, aCity As City, aStreet As Street
Dim aHouse As House, aFamily As Family, aPerson As Person

  'die Reihenfolge, mit der die Elemente erstellt werden ist beliebig
  Set aLand = New_Land("Deutschland")
  'Set mLand = aLand
  Set mLandExplorer = New_LandExplorer(TreeView1, StatusBar1, aLand)

  'Lands.Add aLand
  
  Set aCity = New_City("Freising")
    Set aStreet = New_Street("Alte Poststraße")
    
      Set aHouse = New_House("93")
      Set aFamily = New_Family("Meyer")
        Call aFamily.AddPerson(New_Person("Josef"))
        Call aFamily.AddPerson(New_Person("Elisabeth"))
    Call aStreet.AddHouse(aHouse).AddFamily(aFamily)
    
      Set aHouse = New_House("36")
      Set aFamily = New_Family("Meyer Bichlmaier")
        Call aFamily.AddPerson(New_Person("Andreas"))
        Call aFamily.AddPerson(New_Person("Bernhard"))
        Call aFamily.AddPerson(New_Person("Christa"))
    Call aStreet.AddHouse(aHouse).AddFamily(aFamily)
  Call aCity.AddStreet(aStreet)
    
    Set aStreet = New_Street("Rudolf-Diesel-Straße")
    Call aStreet.AddHouse(New_House("8")).AddFamily(New_Family("Baufuchs"))
  Call aCity.AddStreet(aStreet)
    Set aStreet = New_Street("Prinz-Ludwig-Straße")
    Call aStreet.AddHouse(New_House("31"))
  Call aCity.AddStreet(aStreet)
      
Call aLand.AddCity(aCity)
  
  Set aCity = New_City("Bad Wörishofen")
    Set aStreet = New_Street("Dominikusstraße")
      Set aHouse = New_House("9")
        Set aFamily = New_Family("Geiger")
          Call aFamily.AddPerson(New_Person("Erich"))
          Call aFamily.AddPerson(New_Person("Renate"))
      Call aHouse.AddFamily(aFamily)
        Set aFamily = New_Family("Geiger Meyer")
          Call aFamily.AddPerson(New_Person("Sabine"))
          Call aFamily.AddPerson(New_Person("Oliver"))
          Call aFamily.AddPerson(New_Person("Lukas"))
      Call aHouse.AddFamily(aFamily)
        Set aFamily = New_Family("Geiger Thomas")
          Call aFamily.AddPerson(New_Person("Thomas"))
      Call aHouse.AddFamily(aFamily)
Call aLand.AddCity(aCity).AddStreet(aStreet).AddHouse(aHouse)
        
  Call aLand.ToTreeView(TreeView1, "")
End Sub


Private Sub mnuOpenAllNodes_Click()
  mLandExplorer.OpenAllNodes
End Sub

Private Sub mnuRestoreOpenNodes_Click()
  mLandExplorer.RestoreOpenNodes
End Sub

Private Sub mnuSaveOpenNodes_Click()
  mLandExplorer.SaveOpenNodes
End Sub


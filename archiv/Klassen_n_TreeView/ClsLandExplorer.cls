VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LandExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mTreeView As MSComctlLib.TreeView
Attribute mTreeView.VB_VarHelpID = -1
Private mStatusBar As StatusBar
Private mOpenNodes As Collection
Private mCurNode As MSComctlLib.Node
Private mLand As Land
Private pArr() As String

Private Sub Class_Initialize()
  ClearOpenNodes
End Sub
Friend Sub New_(Tvw As TreeView, StB As StatusBar, Optional aLand As Land)
  Set mTreeView = Tvw
  Set mStatusBar = StB
  Set mLand = aLand
End Sub

Private Function ClearOpenNodes()
  Set mOpenNodes = New Collection
End Function
'Public Property Get MyTreeView() As MSComctlLib.TreeView
'  Set MyTreeView = mTreeView
'End Property
'Public Property Set MyTreeView(mTvw As MSComctlLib.TreeView)
'  Set mTreeView = mTvw
'End Property
'Public Property Get MyStatusBar() As StatusBar
'  Set MyStatusBar = mStatusBar
'End Property
'Public Property Set MyStatusBar(mStBar As StatusBar)
'  Set mStatusBar = mStBar
'End Property

Public Function AddNode(Optional Relative, Optional Relationship, Optional Key, Optional Text As String, Optional Image As StdPicture, Optional selectedImage As StdPicture) As MSComctlLib.Node
  Set AddNode = mTreeView.Nodes.Add(Relative, Relationship, Key, Text, Image, selectedImage)
  Set CurNode = AddNode
End Function
Public Property Get CurNode() As Node
  Set CurNode = mCurNode
End Property
Public Property Set CurNode(Nd As Node)
  Set mCurNode = Nd
  If Not mCurNode Is Nothing Then
    'MsgBox mCurNode.Parent.Parent.Parent.Parent.Parent.Key
    pArr() = Split(mCurNode.FullPath, "\")
  End If
End Property
Public Sub CloseAllNodes()
Dim N As MSComctlLib.Node
  For Each N In mTreeView.Nodes
    N.Expanded = False
  Next
End Sub
Public Sub OpenAllNodes()
Dim N As MSComctlLib.Node
  SaveOpenNodes
  For Each N In mTreeView.Nodes
    N.Expanded = True
  Next
End Sub
Public Sub SaveOpenNodes()
Dim N As MSComctlLib.Node
  ClearOpenNodes
  For Each N In mTreeView.Nodes
    If N.Expanded Then
      mOpenNodes.Add N
    End If
  Next
End Sub
Public Sub RestoreOpenNodes()
Dim N As MSComctlLib.Node
  Call CloseAllNodes
  For Each N In mOpenNodes
    N.Expanded = True
  Next
End Sub

Private Sub mTreeView_DblClick()
  'If Button = vbLeftButton Then
    'Set CurNode = mTreeView.HitTest(x, y)
  'End If
End Sub

Private Sub mTreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    Set CurNode = mTreeView.HitTest(x, y)
  End If
End Sub

Private Sub mTreeView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Set mCurNode = mTreeView.HitTest(x, y)
  If Not mCurNode Is Nothing Then
    mStatusBar.SimpleText = mCurNode.FullPath
  End If
End Sub

'###################'  Get All Current Objects  '###################'
Public Function GetCurLand() As Land
  Set GetCurLand = mLand
End Function
Public Function GetCurCity() As City
Dim StrName As String, N As Long: N = 1
  'If mCurNode Is Nothing Then Exit Function
  If CountDim(pArr) > N - 1 Then StrName = pArr(N)
  If Len(StrName) > 0 Then Set GetCurCity = GetCurLand.City(StrName)
  If GetCurCity Is Nothing Then Set GetCurCity = GetCurLand.CurCity
End Function
Public Function GetCurStreet() As Street
Dim StrName As String, N As Long: N = 2
  'If mCurNode Is Nothing Then Exit Function
  If CountDim(pArr) > N - 1 Then StrName = pArr(N)
  If Len(StrName) > 0 Then Set GetCurStreet = GetCurCity.Street(StrName)
  If GetCurStreet Is Nothing Then Set GetCurStreet = GetCurCity.CurStreet
End Function
Public Function GetCurHouse() As House
Dim StrName As String, N As Long: N = 3
  'If mCurNode Is Nothing Then Exit Function
  If CountDim(pArr) > N - 1 Then StrName = pArr(N)
  If Len(StrName) > 0 Then Set GetCurHouse = GetCurStreet.House(StrName)
  If GetCurHouse Is Nothing Then Set GetCurHouse = GetCurStreet.CurHouse
End Function
Public Function GetCurFamily() As Family
Dim StrName As String, N As Long: N = 4
  'If mCurNode Is Nothing Then Exit Function
  If CountDim(pArr) > N - 1 Then StrName = pArr(N)
  If Len(StrName) > 0 Then Set GetCurFamily = GetCurHouse.Family(StrName)
  If GetCurFamily Is Nothing Then Set GetCurFamily = GetCurHouse.CurFamily
End Function
Public Function GetCurPerson() As Person
Dim StrName As String, N As Long: N = 5
  'If mCurNode Is Nothing Then Exit Function
  If CountDim(pArr) > N - 1 Then StrName = pArr(N)
  If Len(StrName) > 0 Then Set GetCurPerson = GetCurFamily.Person(StrName)
  If GetCurPerson Is Nothing Then Set GetCurPerson = GetCurFamily.CurPerson
End Function

Public Function CountDim(pArr, Optional nDim As Long = 1) As Long
TryE: On Error GoTo CatchE
  CountDim = UBound(pArr, nDim) - LBound(pArr, nDim)
  Exit Function
CatchE:
  CountDim = 0
End Function


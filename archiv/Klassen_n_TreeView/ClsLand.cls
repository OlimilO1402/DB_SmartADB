VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Land"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements SpecList
Implements INamed
Implements IAdd
'Private Base As SpecList
Private mName As String
Private mCol As Collection 'speichert Objekte vom Typ City
Private mNode As Node
'Private mKey As String
Private mCurCity As City

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
Friend Sub New_(LandName As String)
  mName = LandName
End Sub
Public Sub Clear()
  Set mCol = New Collection:  'Set Base = New SpecList
End Sub

Public Function AddCity(NewCity As City, Optional Check As Boolean = False) As City
  If Check Then
    If Not City(NewCity.Name) Is Nothing Then Exit Function
  End If
  Set AddCity = NewCity
  Set CurCity = AddCity
  mCol.Add NewCity
End Function

Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aCity As City, aKey As String
  Tvw.Nodes.Clear
  aKey = ParentKey & "\" & mName
  Set mNode = Tvw.Nodes.Add(, , aKey, mName)
  For Each aCity In mCol
    Call aCity.ToTreeView(Tvw, aKey)
  Next
End Sub
Public Sub UpdateTreeView()
  If Not mNode Is Nothing Then mNode.Text = mName
End Sub

'################'   Properties    '################'
Public Property Get Name() As String
  Name = mName
End Property
Public Property Let Name(StrVal As String)
  mName = StrVal
  UpdateTreeView
End Property

Public Property Get City(CityName As String) As City
  For Each City In mCol ' Base '
    If City.Name = CityName Then Set CurCity = City: Exit Property
  Next
End Property

Public Property Get Cities() As Collection
  Set Cities = mCol
End Property

Public Property Set CurCity(CValue As City)
  Set mCurCity = CValue
End Property
Public Property Get CurCity() As City
  Set CurCity = mCurCity
End Property


''#######################'   Implements SpecList   '#######################'
'Private Sub SpecList_Add(Obj As Object)
'
'End Sub
'
'Private Sub SpecList_Clear()
'  Call Clear
'End Sub
'
'Private Function SpecList_GetEnumerator() As stdole.IUnknown
'  Set SpecList_GetEnumerator = GetEnumerator
'End Function
'Public Function GetEnumerator() As stdole.IUnknown
'  Set GetEnumerator = Base.GetEnumerator
'End Function
'
'Private Property Get SpecList_Item(nIndex As Long) As Variant
'
'End Property
'
'Private Property Get SpecList_SpList() As Collection
'
'End Property

'################'   Implements INamed    '################'
Private Property Let INamed_Name(RHS As String)
  Name = RHS
End Property

Private Property Get INamed_Name() As String
  INamed_Name = Name
End Property


'################'   Implements IAdd    '################'
Private Function IAdd_Add(ObjAdd As INamed, Optional Check As Boolean = False) As INamed
  Set IAdd_Add = AddCity(ObjAdd, Check)
End Function

Private Sub IAdd_ToListBox(LB As ListBox)
  Call ToListBox(LB)
End Sub

Public Sub ToListBox(LB As ListBox)
Dim N As INamed 'Cty City
  For Each N In mCol 'Cty In Cities
    LB.AddItem N.Name 'Cty.Name
  Next
End Sub
Private Function IAdd_GetByName(StrName As String) As IAdd
  Set IAdd_GetByName = City(StrName)
End Function

Private Property Get IAdd_CurrentObj() As INamed
  Set IAdd_CurrentObj = CurCity
End Property


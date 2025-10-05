VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Street"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements SpecList
Implements INamed
Implements IAdd
'Private Base As SpecList
Private m_Name As String
Private m_Col As Collection 'speichert Objekte vom Typ House
Private m_Node As Node
'Private mKey As String
Private m_CurHouse As House

Private Sub Class_Initialize()
    Clear
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
Friend Sub New_(ByVal Name As String)
    m_Name = Name
End Sub
Public Sub Clear()
    Set m_Col = New Collection
End Sub

Public Function AddHouse(NewHouse As House, Optional Check As Boolean = False) As House
    If Check Then
        If Not House(NewHouse.Number) Is Nothing Then Exit Function
    End If
    Set AddHouse = NewHouse
    Set CurHouse = AddHouse
    mCol.Add NewHouse
End Function

Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aHouse As House, aKey As String
  aKey = ParentKey & "\" & mName
  Set mNode = Tvw.Nodes.Add(ParentKey, tvwChild, aKey, mName)
  For Each aHouse In mCol
    Call aHouse.ToTreeView(Tvw, aKey)
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

Public Property Get House(HouseNumber As String) As House
  For Each House In Houses
    If House.Number = HouseNumber Then Set CurHouse = House: Exit Property
  Next
End Property

Public Property Get Houses() As Collection
  Set Houses = mCol
End Property

Public Property Set CurHouse(CValue As House)
  Set mCurHouse = CValue
End Property
Public Property Get CurHouse() As House
  Set CurHouse = mCurHouse
End Property

'################'   Implements INamed    '################'
Private Property Let INamed_Name(RHS As String)
  Name = RHS
End Property
Private Property Get INamed_Name() As String
  INamed_Name = Name
End Property

'################'   Implements IAdd    '################'
Private Function IAdd_Add(ObjAdd As INamed, Optional Check As Boolean = False) As INamed
  Set IAdd_Add = AddHouse(ObjAdd, Check)
End Function

Private Sub IAdd_ToListBox(LB As ListBox)
  Call ToListBox(LB)
End Sub
Public Sub ToListBox(LB As ListBox)
Dim Hus As House
  For Each Hus In Houses
    LB.AddItem Hus.Number
  Next
End Sub
Private Function IAdd_GetByName(StrName As String) As IAdd
  Set IAdd_GetByName = House(StrName)
End Function

Private Property Get IAdd_CurrentObj() As INamed
  IAdd_CurrentObj = CurHouse
End Property


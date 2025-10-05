VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "City"
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
Private m_Col  As Collection 'speichert ClsStreet
Private m_Node As Node
'Private mKey As String
Private m_CurStreet As Street

Private Sub Class_Initialize()
    Clear
End Sub
Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub
Friend Sub New_(Name As String)
    m_Name = Name
End Sub
Public Sub Clear()
    Set m_Col = New Collection
End Sub

Public Function AddStreet(NewStreet As Street, Optional Check As Boolean = False) As Street
  If Check Then
    If Not Street(NewStreet.Name) Is Nothing Then Exit Function
  End If
  Set AddStreet = NewStreet
  Set CurStreet = AddStreet
  mCol.Add NewStreet
End Function

Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aStreet As Street, aKey As String
  aKey = ParentKey & "\" & mName
  Set mNode = Tvw.Nodes.Add(ParentKey, tvwChild, aKey, mName)
  For Each aStreet In mCol
    Call aStreet.ToTreeView(Tvw, aKey)
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

Public Property Get Street(StreetName As String) As Street
  For Each Street In Streets
    If Street.Name = StreetName Then Set CurStreet = Street: Exit Property
  Next
End Property

Public Property Get Streets() As Collection
  Set Streets = mCol
End Property

Public Property Set CurStreet(CValue As Street)
  Set mCurStreet = CValue
End Property
Public Property Get CurStreet() As Street
  Set CurStreet = mCurStreet
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
  Set IAdd_Add = AddStreet(ObjAdd, Check)
End Function

Private Sub IAdd_ToListBox(LB As ListBox)
  Call ToListBox(LB)
End Sub
Public Sub ToListBox(LB As ListBox)
Dim Str As Street
  For Each Str In Streets
    LB.AddItem Str.Name
  Next
End Sub
Private Function IAdd_GetByName(StrName As String) As IAdd
  Set IAdd_GetByName = Street(StrName)
End Function
Private Property Get IAdd_CurrentObj() As INamed
  Set IAdd_CurrentObj = CurStreet
End Property



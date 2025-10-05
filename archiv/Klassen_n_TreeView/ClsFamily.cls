VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Family"
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
Private mCol As Collection  'speichert Objekte vom Typ Person
Private mNode As Node
'Private mKey As String
Private mCurPerson As Person

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
Friend Sub New_(FamilyName As String)
  mName = FamilyName
End Sub
Public Sub Clear()
  Set mCol = New Collection
End Sub

Public Function AddPerson(NewPerson As Person, Optional Check As Boolean = False) As Person
  If Check Then
    If Not Person(NewPerson.Name) Is Nothing Then Exit Function
  End If
  Set AddPerson = NewPerson
  Set CurPerson = AddPerson
  mCol.Add NewPerson
End Function

Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aPerson As Person, aKey As String
  aKey = ParentKey & "\" & mName
  Set mNode = Tvw.Nodes.Add(ParentKey, tvwChild, aKey, mName)
  For Each aPerson In mCol
    Call aPerson.ToTreeView(Tvw, aKey)
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

Public Property Get Person(PersonName As String) As Person
  For Each Person In mCol
    If Person.Name = PersonName Then Set CurPerson = Person:  Exit Property
  Next
End Property

Public Property Get Persons() As Collection
  Set Persons = mCol
End Property

Public Property Set CurPerson(CValue As Person)
  Set mCurPerson = CValue
End Property
Public Property Get CurPerson() As Person
  Set CurPerson = mCurPerson
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
  Set IAdd_Add = AddPerson(ObjAdd, Check)
End Function

Private Sub IAdd_ToListBox(LB As ListBox)
  Call ToListBox(LB)
End Sub
Public Sub ToListBox(LB As ListBox)
Dim Prs As Person
  For Each Prs In Persons
    LB.AddItem Prs.Name
  Next
End Sub
Private Function IAdd_GetByName(StrName As String) As IAdd
  IAdd_GetByName = Person(StrName)
End Function

Private Property Get IAdd_CurrentObj() As INamed
  Set IAdd_CurrentObj = CurPerson
End Property



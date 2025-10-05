VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "House"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements SpecList
Implements INamed
Implements IAdd
'Private Base As SpecList
Private mStrNumber As String 'kann auch 122a oder 36b sein, etc.
Private mCol As Collection 'speichert Objekte vom Typ Family
Private mNode As Node
'Private mKey As String
Private mCurFamily As Family

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
Friend Sub New_(strHouseNr As String)
  mStrNumber = strHouseNr
End Sub
Public Sub Clear()
  Set mCol = New Collection
End Sub

Public Function AddFamily(NewFamily As Family, Optional Check As Boolean = False) As Family
  If Check Then
    If Not Family(NewFamily.Name) Is Nothing Then Exit Function
  End If
  Set AddFamily = NewFamily
  Set CurFamily = AddFamily
  mCol.Add NewFamily
End Function

Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aFamily As Family, aKey As String
  aKey = ParentKey & "\" & "HNr" & mStrNumber
  Set mNode = Tvw.Nodes.Add(ParentKey, tvwChild, aKey, mStrNumber)
  For Each aFamily In mCol
    Call aFamily.ToTreeView(Tvw, aKey)
  Next
End Sub
Public Sub UpdateTreeView()
  If Not mNode Is Nothing Then mNode.Text = mStrNumber
End Sub

'################'   Properties    '################'
Public Property Get Number() As String
  Number = mStrNumber
End Property
Public Property Let Number(StrVal As String)
  mStrNumber = StrVal
  UpdateTreeView
End Property

Public Property Get Family(FamilyName As String) As Family
  For Each Family In Families
    If Family.Name = FamilyName Then Set CurFamily = Family: Exit Property
  Next
End Property

Public Property Get Families() As Collection
  Set Families = mCol
End Property

Public Property Set CurFamily(CValue As Family)
  Set mCurFamily = CValue
End Property
Public Property Get CurFamily() As Family
  Set CurFamily = mCurFamily
End Property

'################'   Implements INamed    '################'
Private Property Let INamed_Name(RHS As String)
  Number = RHS
End Property
Private Property Get INamed_Name() As String
  INamed_Name = Number
End Property

'################'   Implements IAdd    '################'
Private Function IAdd_Add(ObjAdd As INamed, Optional Check As Boolean = False) As INamed
  Set IAdd_Add = AddFamily(ObjAdd, Check)
End Function

Private Sub IAdd_ToListBox(LB As ListBox)
  Call ToListBox(LB)
End Sub
Public Sub ToListBox(LB As ListBox)
Dim Fam As Family
  For Each Fam In Families
    LB.AddItem Fam.Name
  Next
End Sub
Private Function IAdd_GetByName(StrName As String) As IAdd
  Set IAdd_GetByName = Family(StrName)
End Function

Private Property Get IAdd_CurrentObj() As INamed
  Set IAdd_CurrentObj = CurFamily
End Property



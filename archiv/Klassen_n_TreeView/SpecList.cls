VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpecNameList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnumerable
Private mCol As Collection

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Clear()
    Set mCol = New Collection
End Sub

Public Sub AddNamedObj(Obj As Object, Optional Check As Boolean = False)
    mCol.Add Obj
End Sub

Public Property Get SpList() As Collection
    Set SpList = mCol
End Property

Public Property Get Item(nIndex As Long)
    Set Item = mCol(nIndex)
End Property

Private Function IEnumerable_GetEnumerator() As stdole.IUnknown
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function
Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
Attribute GetEnumerator.VB_MemberFlags = "40"
    Set GetEnumerator = mCol.[_NewEnum]
End Function


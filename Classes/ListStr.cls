VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As String
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Sub Add(Value As String)
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To 2 * m_Count)
    m_Arr(m_Count) = Value
    m_Count = m_Count + 1
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Item(ByVal Index As Long) As String
    Item = m_Arr(Index)
End Property
Public Property Let Item(ByVal Index As Long, Value As String)
    m_Arr(Index) = Value
End Property

Public Function ToStr() As String
    If m_Count = 0 Then Exit Function
    ReDim Preserve m_Arr(0 To m_Count - 1)
    ToStr = Join(m_Arr, vbCrLf)
End Function

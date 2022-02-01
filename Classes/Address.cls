VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Street    As String
Public HNr       As String 'could also be like "15a"
Public Info      As String 'additional info (z.B. "Villa Kunderbunt", "Stiege links" etc)
Private m_City   As City
'Private m_Person As Person








Public Property Get City() As City
    Set City = m_City
End Property
Public Property Set City(Value As City)
    Set m_City = Value
    If Not m_City.Addresses.ContainsObj(Me) Then
        m_City.Addresses.Add Me
    End If
End Property

Public Function ToStr() As String
    ToStr = Me.Key
End Function

Public Property Get Key() As String
    Key = Me.Street & " " & Me.HNr
    If City Is Nothing Then Exit Property
    Dim ck As String: ck = City.Key
    If Len(ck) Then Key = Key & ", " & ck
End Property

Public Sub Serial()
    Serializer.Serial_Address Me
End Sub

Public Sub Parse(Params As String)
    XParser.Parse_Address Me, Params
End Sub

Public Function Equals(other As Address) As Boolean
    Equals = MEqv.Eqv_Address(Me, other)
End Function

Public Function Compare(ByVal other As Address) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function


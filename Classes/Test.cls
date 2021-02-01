VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BytVal As Byte
Public IntVal As Integer
Public LngVal As Long
Public CurVal As Currency
Public SngVal As Single
Public DblVal As Double
Public DatVal As Date
Public StrVal As String

Public Property Get Key() As String
    Key = Me.BytVal & " " & Me.IntVal & " " & Me.LngVal & " " & Me.CurVal & " " & Me.SngVal & " " & Me.DblVal & " " & Me.DatVal & " " & Me.StrVal
End Property

Public Sub Serial()
    Serializer.Serial_Test Me
End Sub

Public Sub Parse(Params As String)
    XParser.Parse_Test Me, Params
End Sub

Public Function Equals(other As Address) As Boolean
    Equals = MEqv.Eqv_Test(Me, other)
End Function

Public Function Compare(ByVal other As Test) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function



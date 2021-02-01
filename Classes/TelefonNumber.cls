VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelefonNr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Country As Country
Public City    As City 'braucht von der Stadt nur die Vorwahl, eine Stadt ohne Country gibts nicht, deswegen mit der Stadt auch Landes-Vorwahl
Public Number  As String












Public Property Get Key() As String
    If Not City Is Nothing Then
        If Not City.Country Is Nothing Then
            Dim cov As String: cov = City.Country.Vorwahl
        End If
        Dim civ As String: civ = City.Vorwahl
        If Left(civ, 1) = "0" Then civ = Mid(civ, 2)
    End If
    Key = cov & " " & civ & " / " & Me.Number
End Property

Public Sub Serial()
    Serializer.Serial_TelefonNr Me
End Sub
    
Public Sub Parse(Params As String)
    XParser.Parse_TelefonNr Me, Params
End Sub

Public Function Equals(other As TelefonNr) As Boolean
    Equals = MEqv.Eqv_TelefonNr(Me, other)
End Function

Public Function Compare(ByVal other As TelefonNr) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function


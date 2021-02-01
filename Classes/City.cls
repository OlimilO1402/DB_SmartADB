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
Public Name       As String 'Name der Stadt
Public Nam2       As String 'Namenszusatz wie "am Gewässer sowieso" (z.B. am Lech) 'oder "Stadtteil soundso", (z.B. Bogenhausen) Äh nö nicht Stadtteil
Public PLZ        As String 'Postleitzahl
Public Vorwahl    As String 'Amtliche Vorwahl der Stadt
Private m_Country As Country
Public Addresses  As List

'OK alle Verbindungen 1:n müsen in zwei verbindungen da sein einmal bei 1 als einzelne Referenz und bei n als Liste
'd.h. wir brauchen hier auch eine Liste mit Telefonnummern.
'soll das Public oder Private sein?
'OK Public



Private Sub Class_Initialize()
    Set Addresses = MNew.List(vbObject)
End Sub

Public Property Get Country() As Country
    Set Country = m_Country
End Property
Public Property Set Country(Value As Country)
    Set m_Country = Value
    If Not m_Country.Cities.Contains(Me) Then
        m_Country.Cities.Add Me
    End If
End Property

Public Property Get Key() As String
    Dim k As String
    Dim nz As String: nz = Me.Nam2
    If Len(nz) Then k = " " & nz
    Dim p As String: p = Me.PLZ
    If Len(p) Then p = p & " "
    Key = p & Me.Name & k
End Property

Public Sub Serial()
    Serializer.Serial_City Me
End Sub

Public Sub Parse(Params As String)
    XParser.Parse_City Me, Params
End Sub

Public Function Equals(other As City) As Boolean
    Equals = MEqv.Eqv_City(Me, other)
End Function

Public Function Compare(ByVal other As City) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function


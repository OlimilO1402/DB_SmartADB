VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name    As String 'deutscher Name
Public NameInt As String 'internationaler Name
Public Vorwahl As String 'internationale Vorwahl
Public Cities  As List   'Liste der Städte in diesem Land






Private Sub Class_Initialize()
    Set Cities = MNew.List(vbObject)
End Sub

Public Property Get Key() As String
    Key = Me.Name 'there are no 2 countries with the same name, are there?
End Property

Public Sub Serial()
    Serializer.Serial_Country Me
End Sub

Public Sub Parse(Params As String)
    XParser.Parse_Country Me, Params
End Sub

Public Function Equals(other As Country) As Boolean
    Equals = MEqv.Eqv_Country(Me, other)
End Function

Public Function Compare(ByVal other As Country) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function


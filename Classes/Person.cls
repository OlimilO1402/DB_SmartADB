VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public PreName1  As String
Public PreName2  As String
Public FamName   As String
Public BirthD    As Date
Public Gender    As EGender
Private m_Mother As Person
Private m_Father As Person
'Public PersInLaw As Person 'Verheiratet mit
Public Address   As Address
Public TelNumber As TelefonNr
Public Children  As List 'Of Person
'und was is mit Freunden?
Public Relations As List 'are your children your friends too? yes, but not in general, but it could be and so it should be made possible
'Of Course Friends is in both directoins
'maybe Friends should be named Relations
'Beziehungen
'
'a Person should maybe have Gender
Public Enum EGender
    none    ' 'none'
    male    ' 'male'
    female  ' 'female'
    diverse ' 'diverse'
End Enum

Private Sub Class_Initialize()
    Set Children = MNew.List(vbObject)
    Set Relations = MNew.List(vbObject)
End Sub

Public Property Get Mother() As Person
    Set Mother = m_Mother
End Property
Public Property Set Mother(Value As Person)
    If Value Is Nothing Then
        If Not m_Mother Is Nothing Then m_Mother.Children.RemoveObj Me
    Else
        Set m_Mother = Value
        If Not m_Mother.Children.ContainsObj(Me) Then
            m_Mother.Children.Add Me
        End If
    End If
End Property

Public Property Get Father() As Person
    Set Father = m_Father
End Property
Public Property Set Father(Value As Person)
    If Value Is Nothing Then
        If Not m_Father Is Nothing Then m_Father.Children.RemoveObj Me
    Else
        Set m_Father = Value
        If Not m_Father.Children.ContainsObj(Me) Then
            m_Father.Children.Add Me
        End If
    End If
End Property

Public Function ToStr() As String
    ToStr = Me.Key
End Function

Public Property Get Key() As String
    Key = Me.FamName & " " & Me.PreName1
End Property

Public Sub Serial()
    Serializer.Serial_Person Me
End Sub

Public Sub Parse(Params As String)
    XParser.Parse_Person Me, Params
End Sub

Public Function Equals(other As Person) As Boolean
    Equals = MEqv.Eqv_Person(Me, other)
End Function

Public Function Compare(ByVal other As Person) As Long
    Compare = StrComp(Me.Key, other.Key, VbCompareMethod.vbTextCompare)
End Function



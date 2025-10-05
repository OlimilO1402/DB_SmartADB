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
'Implements SpecList
Implements INamed
'Private Base As SpecList
Private mName As String 'so hier hörts aber auf, jede Person kann aber noch in verschiedenen Häusern oder Wohnungen leben
Private mNode As Node
'Private mKey As String

Friend Sub New_(PersonName As String)
  mName = PersonName
End Sub

Public Property Get Name() As String
  Name = mName
End Property
Public Property Let Name(StrVal As String)
  mName = StrVal
  UpdateTreeView
End Property
Public Sub ToTreeView(ByRef Tvw As MSComctlLib.TreeView, ParentKey As String)
Dim aKey As String
  aKey = ParentKey & "\" & mName
  Set mNode = Tvw.Nodes.Add(ParentKey, tvwChild, aKey, mName)
End Sub
Public Sub UpdateTreeView()
  If Not mNode Is Nothing Then mNode.Text = mName
End Sub

'################'   Implements INamed    '################'
Private Property Get INamed_Name() As String
  INamed_Name = Name
End Property
Private Property Let INamed_Name(RHS As String)
  Name = RHS
End Property

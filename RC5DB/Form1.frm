VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows-Standard
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9551
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_memDB As cMemDB
Private Peter
Private Paul
Private Mary
Private Alice
Private Bob

Private Sub Form_Click()
    Set M = New_c.MemDB
    With M
        .Exec "Create Table Person(ID Integer Primary Key, Name Text Collate NoCase, Gender Integer)"
        .Exec "Create Table Family(ParentID Integer, ChildID Integer, Primary Key (ParentID, ChildID)) Without RowID"
    End With
    
     Peter = AddPerson("Peter", 0)
     Paul = AddPerson("Paul", 0)
     Mary = AddPerson("Mary", 1)
     
     Alice = AddPerson("Alice", 1)
     Bob = AddPerson("Bob", 0)
     
     Set MSHFlexGrid1.DataSource = M.GetTable("Person").DataSource: MsgBox "Person-Table"
     
     'define the Parents/Child-relation of Alice (as Daughter of Peter and Mary)
     AddChildOf Peter, Alice: AddChildOf Mary, Alice
     'define the Child/Parents-relation of Bob (as Son of Paul and Mary)
     AddParentOf Bob, Paul:   AddParentOf Bob, Mary
     
     Set MSHFlexGrid1.DataSource = M.GetTable("Family").DataSource: MsgBox "Family-Table"
     
     Set MSHFlexGrid1.DataSource = GetParentsOf(Bob).DataSource: MsgBox "Parents of Bob"
     
     Set MSHFlexGrid1.DataSource = GetChildrenOf(Peter).DataSource: MsgBox "Children of Peter"
     Set MSHFlexGrid1.DataSource = GetChildrenOf(Mary).DataSource: MsgBox "Children of Mary"
    
     Set MSHFlexGrid1.DataSource = GetSiblingsOf(Alice).DataSource: MsgBox "Siblings of Alice"
End Sub

Function AddPerson(Name As String, ByVal Gender As Long)
    M.ExecCmd "Insert Into Person(Name, Gender) Values(?,?)", Name, Gender
    AddPerson = M.Cnn.LastInsertAutoID
End Function

Sub AddChildOf(ParentID, ChildID)
    M.ExecCmd "Insert Into Family(ParentID, ChildID) Values(?,?)", ParentID, ChildID
End Sub
Sub AddParentOf(ChildID, ParentID)
    M.ExecCmd "Insert Into Family(ParentID, ChildID) Values(?,?)", ParentID, ChildID
End Sub

Function GetParentsOf(ChildID) As cRecordset
    Set GetParentsOf = M.GetTable("Person", "ID In (Select ParentID From Family Where ChildID=" & ChildID & ")")
End Function
Function GetChildrenOf(ParentID) As cRecordset
    Set GetChildrenOf = M.GetTable("Person", "ID In (Select ChildID From Family Where ParentID=" & ParentID & ")")
End Function
Function GetSiblingsOf(ChildID) As cRecordset
    Set GetSiblingsOf = M.GetTable("Person", "ID<>" & ChildID & " AND ID In (Select ChildID From Family Where ParentID In (Select ParentID From Family Where ChildID=" & ChildID & "))")
End Function

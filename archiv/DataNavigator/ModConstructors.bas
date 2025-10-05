Attribute VB_Name = "ModConstructors"
Option Explicit

Public Function New_DataNavigator(mParentFrm As Form, mPanel As PictureBox, StrName As String, Optional MinIndex As Long, Optional CurIndex As Long, Optional MaxIndex As Long) As DataNavigator
  'MsgBox "2"
  Set New_DataNavigator = New DataNavigator
  Call New_DataNavigator.NewC(mParentFrm, mPanel, StrName, MinIndex, CurIndex, MaxIndex)
End Function

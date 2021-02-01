Attribute VB_Name = "MNew"
Option Explicit

Public Function Document(Optional aPFN As PathFileName = Nothing) As Document
    Set Document = New Document
    Document.New_ aPFN
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List
    List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PathFileName(ByVal aPath As String, _
                             Optional ByVal aFileName As String, _
                             Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName
    PathFileName.New_ aPath, aFileName, aExt
End Function

Attribute VB_Name = "MStrings"
Option Explicit ' Zeilen: 44

Public Function PadLeft(StrVal As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    'ist der String l�nger als totalwidth, wird nur der String zur�ckgegeben
    'ansonsten wird der String mit der angegebenen L�nge zur�ckgegeben, der
    'String wird nach rechts ger�ckt, und links mit PadChar aufgef�llt
    'ist PadChar nicht angegeben, so wird mit RSet der String in
    'Spaces eingef�gt.
    Dim StringLength As Long: StringLength = Len(StrVal)
    If StringLength > totalWidth Then
        PadLeft = StrVal
    Else
        If Len(paddingChar) Then
            PadLeft = String$(totalWidth - StringLength, paddingChar) & StrVal
        Else
            PadLeft = Space$(totalWidth)
            RSet PadLeft = StrVal
        End If
    End If
End Function

Public Function PadRight(StrVal As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    'ist der String l�nger als totalwidth, wird nur der String zur�ckgegeben
    'ansonsten wird der String mit der angegebenen L�nge zur�ckgegeben, der
    'String wird nach links ger�ckt, und rechts mit PadChar aufgef�llt
    'ist PadChar nicht angegeben, so wird mit LSet der String in
    'Spaces eingef�gt.
    Dim StringLength As Long: StringLength = Len(StrVal)
    If StringLength > totalWidth Then
        PadRight = StrVal
    Else
        If Len(paddingChar) Then
            PadRight = StrVal & String$(totalWidth - StringLength, paddingChar)
        Else
            PadRight = Space$(totalWidth)
            LSet PadRight = StrVal
        End If
    End If
End Function

'String-Routinen
Public Function RemoveChars(ByVal this As String, CharsToRemove As String) As String
    Dim c As String
    Dim i As Long
    RemoveChars = this
    For i = 1 To Len(CharsToRemove)
        c = Mid$(CharsToRemove, i, 1)
        If InStr(1, this, c) Then
            RemoveChars = Replace(RemoveChars, c, vbNullString)
        End If
    Next
End Function


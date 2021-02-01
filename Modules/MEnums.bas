Attribute VB_Name = "Enums"
Option Explicit

Public Function EGender_ToStr(E As EGender) As String
    Dim s As String
    Select Case E
    Case EGender.none:    s = "" '"none"
    Case EGender.male:    s = "männlich" '"male"
    Case EGender.female:  s = "weiblich" '"female"
    Case EGender.diverse: s = "diverse"
    Case Else: '???
    End Select
    EGender_ToStr = s
End Function

Public Function EGender_Parse(s As String) As EGender
    Dim E As EGender
    Select Case LCase(s)
    Case "none":                E = EGender.none
    Case "male", "männlich":    E = EGender.male
    Case "female", "weiblich":  E = EGender.female
    Case "diverse":             E = EGender.diverse
    Case Else: '???
    End Select
    EGender_Parse = E
End Function

Public Sub EGender_ToListBox(aLBorCB)
    'to ListBox or ComboBox
    'Dim cmb As ComboBox: Set cmb = aLBorCB
    With aLBorCB
    'With cmb
        .Clear
        .AddItem EGender_ToStr(EGender.none)
        .AddItem EGender_ToStr(EGender.male)
        .AddItem EGender_ToStr(EGender.female)
        .AddItem EGender_ToStr(EGender.diverse)
    End With
End Sub



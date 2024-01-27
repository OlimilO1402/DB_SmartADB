Attribute VB_Name = "MEqv"
Option Explicit

Public Function Eqv_Address(Obj1 As Address, Obj2 As Address) As Boolean
    'Eqv_Address = False 'in VB default
    If StrComp(Obj1.Street, Obj2.Street, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.HNr, Obj2.HNr, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.Info, Obj2.Info, VbCompareMethod.vbTextCompare) Then Exit Function
    If Not Obj1.City.Equals(Obj2.City) Then Exit Function
    Eqv_Address = True
End Function

Public Function Eqv_City(Obj1 As City, Obj2 As City) As Boolean
    'Eqv_City = False 'in VB default
    If StrComp(Obj1.Name, Obj2.Name, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.Nam2, Obj2.Nam2, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.PLZ, Obj2.PLZ, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.Vorwahl, Obj2.Vorwahl, VbCompareMethod.vbTextCompare) Then Exit Function
    If Not Obj1.Country.Equals(Obj2.Country) Then Exit Function
    Eqv_City = True
End Function

Public Function Eqv_Country(Obj1 As Country, Obj2 As Country) As Boolean
    'Eqv_Country = False 'in VB default
    If StrComp(Obj1.Name, Obj2.Name, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.NameInt, Obj2.NameInt, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.Vorwahl, Obj2.Vorwahl, VbCompareMethod.vbTextCompare) Then Exit Function
    Eqv_Country = True
End Function

Public Function Eqv_Person(Obj1 As Person, Obj2 As Person) As Boolean
    'Eqv_Person = False 'in VB default
    If StrComp(Obj1.PreName1, Obj2.PreName1, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.PreName2, Obj2.PreName2, VbCompareMethod.vbTextCompare) Then Exit Function
    If StrComp(Obj1.FamName, Obj2.FamName, VbCompareMethod.vbTextCompare) Then Exit Function
    If Obj1.BirthD <> Obj2.BirthD Then Exit Function
    
    If Obj1.Mother Is Nothing Then
        If Not Obj2.Mother Is Nothing Then Exit Function
    Else
        If Obj2.Mother Is Nothing Then Exit Function
        If Not Obj1.Mother.Equals(Obj2.Mother) Then Exit Function
    End If
    'das gleich nochmal beim Father
    If Obj1.Father Is Nothing Then
        If Not Obj2.Father Is Nothing Then Exit Function
    Else
        If Obj2.Father Is Nothing Then Exit Function
        If Not Obj1.Father.Equals(Obj2.Father) Then Exit Function
    End If
    'und bei der Adresse und TelefonNr:
    If Obj1.Address Is Nothing Then
        If Not Obj2.Address Is Nothing Then Exit Function
    Else
        If Obj2.Address Is Nothing Then Exit Function
        If Not Obj1.Address.Equals(Obj2.Address) Then Exit Function
    End If
    
    If Obj1.TelNumber Is Nothing Then
        If Not Obj2.TelNumber Is Nothing Then Exit Function
    Else
        If Obj2.TelNumber Is Nothing Then Exit Function
        If Not Obj1.TelNumber.Equals(Obj2.TelNumber) Then Exit Function
    End If
    Eqv_Person = True
End Function

Public Function Eqv_TelefonNr(Obj1 As TelefonNr, Obj2 As TelefonNr) As Boolean
    'Eqv_TelefonNr = False 'in VB default
    If Not Obj1.City.Equals(Obj2.City) Then Exit Function
    If StrComp(Obj1.Number, Obj2.Number, VbCompareMethod.vbTextCompare) Then Exit Function
    Eqv_TelefonNr = True
End Function

Public Function Eqv_Test(Obj1 As Test, Obj2 As Test) As Boolean
    'Eqv_TelefonNr = False 'in VB default
    If Not Obj1.BytVal = Obj2.BytVal Then Exit Function
    If Not Obj1.IntVal = Obj2.IntVal Then Exit Function
    If Not Obj1.LngVal = Obj2.LngVal Then Exit Function
    If Not Obj1.CurVal = Obj2.CurVal Then Exit Function
    If Not Obj1.SngVal = Obj2.SngVal Then Exit Function
    If Not Obj1.DblVal = Obj2.DblVal Then Exit Function
    If Not Obj1.DatVal = Obj2.DatVal Then Exit Function
    If Not Obj1.StrVal = Obj2.StrVal Then Exit Function
    Eqv_Test = True
End Function

'Public Function Max(V1, V2)
'    If V1 > V2 Then Max = V1 Else Max = V2
'End Function
'Public Function Min(V1, V2)
'    If V1 < V2 Then Min = V1 Else Min = V2
'End Function
'

Attribute VB_Name = "XParser"
Option Explicit
Dim hashcol As Collection

Private Function ParseLine(sLine As String, hashtag_out As String, typnam_out As String, params_out As String) As Boolean
    If Left(sLine, 1) <> "#" Then Exit Function
    Dim pos1 As Long: pos1 = InStr(1, sLine, "=")
    If pos1 > 0 Then
        hashtag_out = Mid(sLine, 1, pos1 - 1)
        hashtag_out = Replace(hashtag_out, " ", "")
        Dim pos2 As Long: pos2 = InStr(pos1, sLine, "(")
        If pos2 > 0 Then
            typnam_out = Mid(sLine, pos1 + 1, pos2 - pos1 - 1)
            params_out = Mid(sLine, pos2 + 1, Len(sLine) - pos2 - 1)
        End If
    End If
    ParseLine = True
End Function

Public Sub ParseLines(sLines() As String, doc As Document)
    'soll man hier drinnen ein array mit den Objekten anlegen, oder soll die Listenklasse ein public EnsureCapacity haben, ja soll sie
    'but for now we do it just simple, oh and it could be stupid, because
    'the array could be too small according what line-numbers will come
    
    'nö hashtag: # also brauchen wir eine hashlist
    'beim einlesen brauchen wir eine Hashlist
    'beim rausschreiben brauchen wir eine Hashlist mit den Objptrs
    'wir müssen zuerst alle objekte anlegen
    'und dann nochmal alle durchwandern und parsen
Try: On Error GoTo Catch
    Set hashcol = New Collection
    Dim All As List: Set All = MNew.List(vbObject)
    
    hashcol.Add Nothing, "$" '"#0"
    Dim i As Long ', sLine As String
    For i = 0 To UBound(sLines)
        Dim hashtag As String, typenam As String, Params As String
        'Debug.Print sLines(i)
        If ParseLine(sLines(i), hashtag, typenam, Params) Then
            Dim obj As Object
            Select Case LCase(typenam)
            Case "address":   Set obj = New Address
            Case "city":      Set obj = New City
            Case "country":   Set obj = New Country
            Case "person":    Set obj = New Person
            Case "telefonnr": Set obj = New TelefonNr
            Case Else:
            End Select
            hashcol.Add obj, hashtag
            If Not obj Is Nothing Then
                doc.Add obj
                All.Add obj
            End If
            sLines(i) = Params
        Else
            'die zeile is iwie Schrott, dann hau wech
            sLines(i) = ""
        End If
    Next
    Dim c As Long
    For i = 0 To UBound(sLines)
        Dim sLine As String: sLine = sLines(i)
        If Len(sLine) Then
            'Debug.Print """" & sLine & """"
            All.Item(c).Parse sLine
            c = c + 1
        End If
    Next
    Exit Sub
Catch:
    'If Err Then
        ErrHandler "ParseLines", "line: " & sLine & vbCrLf & "nr: " & hashtag & "typ: " & typenam & "params: " & Params

        'MsgBox Err.Description & vbCrLf & sLine
    'End If
End Sub

'Alle primitiven Datentypen stehen sofort da, alle Objekttypen stehen nur mit einem Index da
Public Function ParseParam(Prop, ByVal Value As String, obj As Object, ParamName As String)
Try: On Error GoTo Catch
    Value = Trim$(Value)
    If Len(Value) = 0 Then Exit Function
    Dim v: v = Value
    Dim msg As String
    Dim vtDst As VbVarType: vtDst = VarType(Prop)
    Dim sTypDst As String: If vtDst = vbObject Then sTypDst = TypeName(Prop)
    If IsNumeric(v) Then
        'check min max range
        msg = "Out of range"
        Select Case vtDst
        Case vbByte:     v = CByte(v) 'If Not (0 <= v And v <= 255) Then 'do not check bounds here!
        Case vbInteger:  v = CInt(v)  'If Not (-32768 <= v And v <= 32767) Then 'do not check bounds here!
        Case vbLong:     v = CLng(v)  'If Not (-2147483647 <= v And v <= 2147483647) Then'do not check bounds here!
        Case vbCurrency: v = CCur(v)  'do not check bounds here!
        Case vbSingle:   v = CSng(Val(Replace(v, ",", ".")))  'do not check bounds here!
        Case vbDouble:   v = CDbl(Val(Replace(v, ",", ".")))  'do not check bounds here!
        Case vbDate:     v = CDate(v)
        End Select
    'ElseIf IsDate(p) Then
    '    ParseParam = CDate(p)
    Else
        Dim c1 As String: c1 = Left(Value, 1)
        Select Case c1
        Case "'"
            If Right(Value, 1) = "'" Then
                v = Mid$(Value, 2, Len(Value) - 2)
            Else
                v = Mid$(Value, 2, Len(Value) - 1)
            End If
        Case """"
            If Right(Value, 1) = """" Then
                v = Mid$(Value, 2, Len(Value) - 2)
            Else
                v = Mid$(Value, 2, Len(Value) - 1)
            End If
        Case "#"
            Set ParseParam = hashcol.Item(Value)
            Exit Function
        Case "$"
            Select Case vtDst
            Case vbString: v = vbNullString 'Null
            Case vbObject: Set ParseParam = Nothing: Exit Function
            Case Else:     v = 0
            End Select
        Case "." 'a Enum-type, enum types must have points in front and behind like so: ".female."
            If Right(Value, 1) = "." Then
                v = Mid$(Value, 2, Len(Value) - 2)
            Else
                v = Mid$(Value, 2, Len(Value) - 1)
            End If
            'If Right(v, 1) = "." Then
            'ParseParam = Value
            'we do parsing of strings for enum types not in here, but afterwards in the function Parse_<MyClass>
        Case Else
            Select Case vtDst
            Case vbDate
                v = Replace(v, "-", ".")
                ParseParam = CDate(v)
            Case vbByte, vbInteger, vbLong, vbCurrency, vbSingle, vbDouble, vbDate
                GoTo Catch
            Case vbObject
                'Set ParseParam = Nothing
                'Set ParseParam = Value
                Exit Function
            Case Else
                'OK versuchen wir mal irgendwas
                'ParseParam = Value
                'Exit Function
            End Select
        End Select
    End If
    ParseParam = v
    Exit Function
Catch:
    Dim vtSrc As VbVarType, sTypSrc As String
    ParseDetectType vtDst, sTypDst, Value, vtSrc, sTypSrc
    If (vtDst <> vtSrc) Or (sTypDst <> sTypSrc) Then
        'Type mismatch
        msg = "Attention type mismatch in: " & TypeName(obj) & "::" & ParamName & vbCrLf & _
              "expect datatype: " & VarType_ToStr(vtDst) & IIf(Len(sTypDst) And vtDst = vbObject, " " & sTypDst, "") & vbCrLf & _
              "given datatype: " & VarType_ToStr(vtSrc) & IIf(Len(sTypSrc) And vtSrc = vbObject, " " & sTypSrc, "") & ": """ & Value & """" & IIf(Len(msg), vbCrLf & msg, "")
        'MsgBox msg, vbCritical
        'ErrHandler "ParseParam", msg

    End If
    ErrHandler "ParseParam", msg
End Function
Private Sub ParseDetectType(ByVal vtDst As VbVarType, ByVal TypDst As String, ByVal Value As String, vt_out As VbVarType, typ_out As String)
'so diese Funktion erst im Fehlerfall aufrufen
    'Destin: vtDst, typDst;
    'Source: Value, vt_out, typ_out
    Value = Trim(Value)
    If IsNumeric(Value) Then
        If Is_Int(Value) Then
            If IsLng(Value) Then vt_out = vbLong:     typ_out = "Long"
            If IsInt(Value) Then vt_out = vbInteger:  typ_out = "Integer"
            If IsByt(Value) Then vt_out = vbByte:     typ_out = "Byte"
        Else
            If IsDec(Value) Then vt_out = vbDecimal:  typ_out = "Decimal"
            If IsCur(Value) Then vt_out = vbCurrency: typ_out = "Currency"
            If IsDbl(Value) Then vt_out = vbDouble:   typ_out = "Double"
            If IsSng(Value) Then vt_out = vbSingle:   typ_out = "Single"
            If IsDate(Value) And vtDst = vbDate Then vt_out = vbDate: typ_out = "Date"
        End If
    Else
        If Left(Value, 1) = "#" Then
            'If vtDst = vbObject Then
            vt_out = vbObject: typ_out = "Object"
            Dim obj As Object: Set obj = hashcol.Item(Value)
            If Not obj Is Nothing Then typ_out = typ_out & "(" & TypeName(obj) & ")"
            'jetzt muss man noch überprüfen ob der typ stimmt
        ElseIf Left(Value, 1) = "'" And Right(Value, 1) = "'" Then
            vt_out = vbString: typ_out = "String"
        Else
            Dim v: v = Value
            vt_out = VarType(v)
            'so wenn kein String, weil keine '' vorhanden dann kann es entweder doch ein STring sein, kommt also drauf an was typDst ist
            'oder wenn typDst angegeben ist dann kann es noch ein Enum-Typ sein.
        End If
    End If
End Sub
Function Is_Int(ByVal v) As Boolean
Try: On Error GoTo Catch
    If InStr(1, v, ".") > 0 Then Exit Function
    If InStr(1, v, ",") > 0 Then Exit Function
    Dim i: i = CDec(Int(v))
    Is_Int = i = CDec(v)
    Exit Function
Catch:
    'ErrHandler "Is_Int", v
End Function

Function IsByt(ByVal v) As Boolean
Try: On Error GoTo Catch
    v = CByte(v)
    IsByt = True
Catch:
    'ErrHandler "IsByt", v
End Function
Function IsInt(v) As Boolean
Try: On Error GoTo Catch
    v = CInt(v)
    IsInt = True
Catch:
End Function
Function IsLng(v) As Boolean
    IsLng = ((-2147483648# <= v) And (v <= 2147483647))
End Function
Function IsSng(v) As Boolean
Try: On Error GoTo Catch
    IsSng = VarType(CSng(v)) = vbSingle
    Exit Function
Catch:
End Function
Function IsDbl(v) As Boolean
Try: On Error GoTo Catch
    IsDbl = VarType(CDbl(v)) = vbDouble
    Exit Function
Catch:
End Function
Function IsCur(v) As Boolean
Try: On Error GoTo Catch
    IsCur = VarType(CCur(v)) = vbCurrency
    Exit Function
Catch:
End Function
Function IsDec(v) As Boolean
Try: On Error GoTo Catch
    IsDec = VarType(CDec(v)) = vbDecimal
    Exit Function
Catch:
End Function
'Function Double_TryParse(ByVal Value As String, ByRef DblVal_out As Double) As Boolean
'Try: On Error GoTo Catch
'    Value = Replace(Value, ",", ".")
'    DblVal_out = CDbl(Val(Value))
'    Double_TryParse = True
'    Exit Function
'Catch:
'End Function
'Function Single_TryParse(ByVal Value As String, ByRef SngVal_out As Single) As Boolean
'Try: On Error GoTo Catch
'    Value = Replace(Value, ",", ".")
'    SngVal_out = CSng(Val(Value))
'    Single_TryParse = True
'    Exit Function
'Catch:
'End Function

'Private Sub ParseProperty(ByRef Dst, ByVal p As String)
'    'es soll ein Fehler ausgegeben werden, wenn die Typen nicht übereinstimmen
'    'die Fehlermeldung soll folgende Infos ausgeben:
'    '* welche Zeilennummer bzw hashwert
'    '* welcher Objecttyp
'    '* welcher Parametername
'    '* welcher typ des Parameter wird erwartet
'    '* in welchem Bereich darf sich der Wert befinden
'    'errmsg = "Für ein Objekt vom Typ: " & Typename(obj) & " wird für das Property: " & PropName & " der Typ: " vartypetostr(vt_Dst) " erwartet." & vbcrlf & "Stattdessen wurde ein Wert vom Typ "
'
'
'    'error(13) = Typen unverträglich
'    '???????????????????????????????????????
'    p = Trim$(p)
'    Dim vt_Dst As VbVarType: vt_Dst = VarType(Dst)
'    Dim typDst As String: typDst = VarType_ToStr(vt_Dst)
'    If IsNumeric(p) Then
'        'if vt_Dst=vbDate
'        Dst = p
'    'ElseIf IsDate(p) Then
'    '    ParseParam = CDate(p)
'    Else
'        Dim c1 As String: c1 = Left(p, 1)
'        Select Case c1
'        Case "'"
'            If Right(p, 1) = "'" Then
'                ParseProperty = Mid$(p, 2, Len(p) - 2)
'            Else
'                ParseProperty = Mid$(p, 2, Len(p) - 1)
'            End If
'        Case """"
'            If Right(p, 1) = """" Then
'                ParseProperty = Mid$(p, 2, Len(p) - 2)
'            Else
'                ParseProperty = Mid$(p, 2, Len(p) - 1)
'            End If
'        Case "#"
'            Set ParseProperty = hashcol.Item(p)
'        End Select
'    End If
'End Sub
Public Sub Parse_Address(obj As Address, Params As String)
Try: On Error GoTo Catch
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    Dim i As Long
    With obj
        If u >= i Then .Street = ParseParam(.Street, sp(i), obj, "Street")
        i = i + 1
        If u >= i Then .HNr = ParseParam(.HNr, sp(i), obj, "HNr")
        i = i + 1
        If u >= i Then .Info = ParseParam(.Info, sp(i), obj, "Info")
        i = i + 1
        If u >= i Then
            Set .City = ParseParam(.City, sp(i), obj, "City")
            If Not .City Is Nothing Then
                .City.Addresses.Add obj
            End If
        End If
    End With
    Exit Sub
Catch:
    ErrHandler "Parse_Address", Params

    'If Err Then
    '    MsgBox Err.Description
    'End If
End Sub

Public Sub Parse_City(obj As City, Params As String)
Try: On Error GoTo Catch
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    Dim i As Long
    With obj
        If u >= 0 Then .Name = ParseParam(.Name, sp(0), obj, "Name")
        If u >= 1 Then .Nam2 = ParseParam(.Nam2, sp(1), obj, "Nam2")
        If u >= 2 Then .PLZ = ParseParam(.PLZ, sp(2), obj, "PLZ")
        If u >= 3 Then .Vorwahl = ParseParam(.Vorwahl, sp(3), obj, "Vorwahl")
        If u >= 4 Then
            Set .Country = ParseParam(.Country, sp(4), obj, "Country")
            If Not .Country Is Nothing Then
                .Country.Cities.Add obj
            End If
        End If
    End With
    Exit Sub
Catch:
    ErrHandler "Parse_City", Params
End Sub

Public Sub Parse_Country(obj As Country, Params As String)
Try: On Error GoTo Catch
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    With obj
        If u >= 0 Then .Name = ParseParam(.Name, sp(0), obj, "Name")
        If u >= 1 Then .NameInt = ParseParam(.NameInt, sp(1), obj, "NameInt")
        If u >= 2 Then .Vorwahl = ParseParam(.Vorwahl, sp(2), obj, "Vorwahl")
    End With
    Exit Sub
Catch:
    ErrHandler "Parse_Country", Params
End Sub

Public Sub Parse_Person(obj As Person, Params As String)
Try: On Error GoTo Catch
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    With obj
        If u >= 0 Then .PreName1 = ParseParam(.PreName1, sp(0), obj, "PreName1")
        If u >= 1 Then .PreName2 = ParseParam(.PreName2, sp(1), obj, "PreName2")
        If u >= 2 Then .FamName = ParseParam(.FamName, sp(2), obj, "FamName")
        If u >= 3 Then .BirthD = ParseParam(.BirthD, sp(3), obj, "BirthD")
        If u >= 4 Then .Gender = EGender_Parse(ParseParam("Gender", sp(4), obj, "Gender"))
        
        If u >= 5 Then Set .Mother = ParseParam(.Mother, sp(5), obj, "Mother")
        If u >= 6 Then Set .Father = ParseParam(.Father, sp(6), obj, "Father")
        If u >= 7 Then Set .Address = ParseParam(.Address, sp(7), obj, "Address")
        If u >= 8 Then Set .TelNumber = ParseParam(.TelNumber, sp(8), obj, "TelNumber")
        'If Not .Mother Is Nothing Then .Mother.Children.Add obj 'nein wird in set Mother gemacht
        'If Not .Father Is Nothing Then .Father.Children.Add obj
    End With
    Exit Sub
Catch:
    ErrHandler "Parse_Person", Params
End Sub

Public Sub Parse_TelefonNr(obj As TelefonNr, Params As String)
Try: On Error GoTo Catch
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    With obj
        If u >= 0 Then Set .City = ParseParam(.City, sp(0), obj, "City")
        If u >= 1 Then .Number = ParseParam(.Number, sp(1), obj, "Number")
    End With
    Exit Sub
Catch:
    ErrHandler "Parse_TelefonNr", Params
End Sub

Private Function VarType_ToStr(vt As VbVarType) As String
    Dim s As String
    Select Case vt
    Case vbBoolean:  s = "Boolean"
    Case vbByte:     s = "Byte"
    Case vbInteger:  s = "Integer"
    Case vbLong:     s = "Long"
    Case vbSingle:   s = "Single"
    Case vbDouble:   s = "Double"
    Case vbCurrency: s = "Currency"
    Case vbDecimal:  s = "Decimal"
    Case vbDate:     s = "Date"
    Case vbObject:   s = "Object"
    Case vbString:   s = "String"
    Case Else:
    End Select
    VarType_ToStr = s
End Function

Public Sub Parse_Test(obj As Test, Params As String)
    Dim sp() As String: sp = Split(Params, ",")
    Dim u As Long: u = UBound(sp)
    With obj
        If u >= 0 Then .BytVal = ParseParam(.BytVal, sp(0), obj, "BytVal")
        If u >= 1 Then .IntVal = ParseParam(.IntVal, sp(1), obj, "IntVal")
        If u >= 2 Then .LngVal = ParseParam(.IntVal, sp(2), obj, "LngVal")
        If u >= 3 Then .CurVal = ParseParam(.IntVal, sp(3), obj, "CurVal")
        If u >= 4 Then .SngVal = ParseParam(.IntVal, sp(4), obj, "SngVal")
        If u >= 5 Then .DblVal = ParseParam(.IntVal, sp(5), obj, "DblVal")
        If u >= 6 Then .DatVal = ParseParam(.IntVal, sp(6), obj, "DatVal")
        If u >= 7 Then .StrVal = ParseParam(.IntVal, sp(7), obj, "StrVal")
        'If u >= 1 Then .IntVal = ParseParam(.IntVal, sp(1))
    End With
End Sub

'# 1=Country('Deutschland', 'Germany', '+49')
'# 2=Country('Österreich' , 'Austria', '+43')
'# 3=Country('Dänemark'   , 'Denmark', '+45')
'# 4=Country('Schweiz'    , 'Swiss'  , '+41')
'# 5=City('Freising'      , ''       , '85356', '08161', #1)
'# 6=City('München'       , ''       , '80331', '089'  , #1)
'# 7=City('Bad Wörishofen', ''       , '86825', '08247', #1)
'# 8=City('Landsberg'     , 'am Lech', '86899', '08191', #1)
'# 9=City('Wien'          , ''       , '1010' , '01'   , #2)
'#10=City('Wien'          , ''       , '1020' , '01'   , #2)
'#11=City('Aarhus'        , ''       , '8000' , '8'    , #3)
'#12=TelefonNr(#5, '4920549')
'#13=TelefonNr(#7, '34732')
''          PreName1   , PreName2         , FamName, BirthD       , Mother, Father, Address, TelNr
'#14=Person('Oliver'   , ''               , 'Meyer', '14.feb.1970', #15   , #16   , #19    , #13)
'#15=Person('Elisabeth', 'Maria Franziska', 'Meyer', '02.jul.1944', #0    , #0    , #20    , #12)
'#16=Person('Josef'    , 'Andreas'        , 'Meyer', '06.apr.1942', #0    , #0    , #20    , #12)
'#17=Person('Andreas'  , 'Peter'          , 'Meyer', '22.dez.1968', #15   , #16   , #21    , #18)
'#18=TelefonNr(#5, '148653')
'#19=Address('Dominikusstraße'  , '9' , '', #7)
'#20=Address('Altenhauserstraße', '33', '', #5)
'#21=Address('Marzlinger Fußweg', '13', '', #5)

'copy this same function to every class, form or module
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry("XParser", FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError("XParser", FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function



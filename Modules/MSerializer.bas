Attribute VB_Name = "Serializer"
Option Explicit
Dim hashcol As Collection
'Public Str  As ListStr
Public LStr As ListStr

'Bsp:
'ICX
'          PreName1   , PreName2         , FamName, BirthD            , Mother, Father, Address, TelNr
'#1=Person("Oliver"   , ""               , "Meyer", "1970-02-14 00:00", #2    ,     #3,    #4,       #5)
'#2=Person("Elisabeth", "Maria Franziska", "Meyer", "1944-07-02 00:00", ''    , ''    , #4, #5)
Public Sub Init()
    Set hashcol = New Collection
    Set LStr = New ListStr
End Sub

Public Function ContainsObj(obj As Object) As Boolean
    On Error Resume Next
    Dim Key As String: Key = "#" & ObjPtr(obj)
    If IsEmpty(hashcol(Key)) Then:  'DoNothing
    ContainsObj = (Err.Number = 0)
    'Debug.Print Key & ": " & ContainsObj '& vbCrLf
    On Error GoTo 0
End Function

'Also jetzad wie serialisiert man ein Objekt mit allen abhängigen Objekten?
'd.h. zuerst das Objekt und dann alle Unterobjekte
'OK, so müssen wir starten
'bei Copy & Paste
Public Sub Serial_Doc(doc As Document) 'As String
    'Dim hk As String
    'Dim s As String
    'Set hashcol = New Collection
    'Set Str = New ListStr
    AddHashes doc.Countries
    AddHashes doc.Cities
    AddHashes doc.TelefonNrs
    AddHashes doc.Addresses
    AddHashes doc.Persons
'    Do Until c = 5
'        c = c + 1
'        Select Case c
'        Case 1: Set All = doc.Countries
'        Case 2: Set All = doc.Cities
'        Case 3: Set All = doc.TelefonNrs
'        Case 4: Set All = doc.Addresses
'        Case 5: Set All = doc.Persons
'        End Select
'        'erstmal alle Objekte in die Hashlist packen
'        Dim hObj As String, Key As String
'        For i = 0 To All.Count - 1
'            Set obj = All.Item(i)
'            hashcol.Add "#" & hashcol.Count + 1, "#" & ObjPtr(obj)
'        Next
'    Loop

    AddSerials doc.Countries
    AddSerials doc.Cities
    AddSerials doc.TelefonNrs
    AddSerials doc.Addresses
    AddSerials doc.Persons
    
'    Dim obj As Object
'    Dim All As List ': Set all = doc.AllObjects
'    Dim i As Long, c As Long
'    'c = 0
'    Do Until c = 5
'        c = c + 1
'        Select Case c
'        Case 1: Set All = doc.Countries
'        Case 2: Set All = doc.Cities
'        Case 3: Set All = doc.TelefonNrs
'        Case 4: Set All = doc.Addresses
'        Case 5: Set All = doc.Persons
'        End Select
'
'        For i = 0 To All.Count - 1
'            Set obj = All.Item(i): obj.Serial
'        Next
'        'dann in hashcol suchen ob schon drin
'        'wenn nicht dann neue nummer =hashcol.count und in hashcol speichern
'        'If Not ContainsObj(obj) Then
'        'hk = "#" & hashcol.Count + 1
'        'hashcol.Add hk, CStr(ObjPtr(obj))
'        's = s & hk & "=" & TypeName(obj) & "(" & obj.ToStr & ")" & vbCrLf
'        'Else
'        'nix machen
'        'End If
'        's = s & obj.ToStr & vbCrLf
'    Loop
    Set hashcol = New Collection
    'ToStr_Doc = s
End Sub

Private Sub AddHashes(aList As List)
    Dim i As Long, obj As Object
    For i = 0 To aList.Count - 1
        Set obj = aList.Item(i)
        hashcol.Add "#" & hashcol.Count + 1, "#" & ObjPtr(obj)
    Next
End Sub

Private Sub AddSerials(aList As List)
    Dim i As Long, obj As Object
    For i = 0 To aList.Count - 1
        Set obj = aList.Item(i)
        obj.Serial
    Next
End Sub

Private Function ToStr_Cls(obj As Object) As String
    Dim hObj As String
    Dim Key As String: Key = "#" & ObjPtr(obj)
    If ContainsObj(obj) Then
        'wenn hash schon drin dann rauslesen
        hObj = hashcol.Item(Key)
    Else
        'wenn hash nicht drin dann reinpacken
        hObj = "#" & hashcol.Count + 1
        hashcol.Add hObj, Key
    End If
    ToStr_Cls = hObj & "=" & TypeName(obj) & "("
End Function

Private Function ToStr_Param(p) As String
    Dim s As String
    If IsObject(p) Then
        Dim obj As Object: Set obj = p
        If obj Is Nothing Then
            s = "$"
        Else
            'wenn noch nicht drin dann serialisieren
            If Not ContainsObj(obj) Then
                obj.Serial
            End If
            'OK jetzt auf jedenfall drin dann den Hash zurückgeben:
            'Nein halt stopp!!!!
            Dim Key As String: Key = "#" & ObjPtr(obj)
            s = hashcol(Key)
        End If
    Else
        Dim vt As VbVarType: vt = VarType(p)
        Select Case vt
        Case vbString: s = IIf(Len(p), "'" & p & "'", "$")
        Case vbDate:   s = "'" & Format(p, "dd.mmm.yyyy") & "'"
        Case Else:     s = p
        End Select
    End If
    ToStr_Param = s
End Function

'Public Street As String
'Public HNr    As String 'could also be like "15a"
'Public Info   As String 'additional info (z.B. "Villa Kunderbunt", "Stiege links" etc)
'Public City   As City
'tja Idee:
'* den aktuellen count-Index holen
'* einen leeren string adden,
'* erst wenn alle Objekte serialisiert sind,
'  am Ende mit dem reservierten Index den String zuweisen
Public Sub Serial_Address(obj As Address)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.Street) & ", "
        s = s & ToStr_Param(.HNr) & ", "
        s = s & ToStr_Param(.Info) & ", "
        s = s & ToStr_Param(.City) & ")"
    End With
    LStr.Item(i) = s
    'ToStr_Address = s & vbCrLf
End Sub


'Public Name    As String 'Name der Stadt
'Public Nam2    As String 'Namenszusatz wie "am Gewässer sowieso" (z.B. am Lech) 'oder "Stadtteil soundso", (z.B. Bogenhausen) Äh nö nicht Stadtteil
'Public PLZ     As String 'Postleitzahl
'Public Vorwahl As String 'Amtliche Vorwahl der Stadt
'Public Country As Country
'Public Addresses As List
Public Sub Serial_City(obj As City)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.Name) & ", "
        s = s & ToStr_Param(.Nam2) & ", "
        s = s & ToStr_Param(.PLZ) & ", "
        s = s & ToStr_Param(.Vorwahl) & ", "
        s = s & ToStr_Param(.Country) & ")"
    End With
    LStr.Item(i) = s
    'ToStr_City = s & vbCrLf
End Sub


'Public Name    As String 'deutscher Name
'Public NameInt As String 'internationaler Name
'Public Vorwahl As String 'internationale Vorwahl
'Public Cities  As List   'Liste der Städte in diesem Land
Public Sub Serial_Country(obj As Country)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.Name) & ", "
        s = s & ToStr_Param(.NameInt) & ", "
        s = s & ToStr_Param(.Vorwahl) & ")"
    End With
    LStr.Item(i) = s
    'ToStr_Country = s & vbCrLf
End Sub


'Public PreName1  As String
'Public PreName2  As String
'Public FamName   As String
'Public BirthD    As Date
'Private m_Mother As Person
'Private m_Father As Person
''Public PersInLaw As Person 'Verheiratet mit
'Public Address   As Address
'Public TelNumber As TelefonNr
'Public Children  As List 'Of Person
''und was is mit Freunden?
'Public Friends   As List 'are your children your friends too? no, not in general, but it could be and so it should be made possible
Public Sub Serial_Person(obj As Person)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.PreName1) & ", "
        s = s & ToStr_Param(.PreName2) & ", "
        s = s & ToStr_Param(.FamName) & ", "
        s = s & ToStr_Param(.BirthD) & ", "
        s = s & "." & EGender_ToStr(.Gender) & "." & ", "
        s = s & ToStr_Param(.Mother) & ", "
        s = s & ToStr_Param(.Father) & ", "
        s = s & ToStr_Param(.Address) & ", "
        s = s & ToStr_Param(.TelNumber) & ")"
    End With
    LStr.Item(i) = s
    'ToStr_Person = s & vbCrLf
End Sub

'Public City    As City 'braucht nur City, weil ein City ohne Country gibts gar nicht
'Public Number  As String
Public Sub Serial_TelefonNr(obj As TelefonNr)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.City) & ", "
        s = s & ToStr_Param(.Number) & ")"
    End With
    LStr.Item(i) = s
    'ToStr_TelefonNr = s & vbCrLf
End Sub

Public Sub Serial_Test(obj As Test)
    Dim s As String: s = ToStr_Cls(obj)
    Dim i As Long: i = LStr.Count: LStr.Add vbNullString
    With obj
        s = s & ToStr_Param(.BytVal) & ", "
        s = s & ToStr_Param(.IntVal) & ", "
        s = s & ToStr_Param(.LngVal) & ", "
        s = s & ToStr_Param(.CurVal) & ", "
        s = s & ToStr_Param(.SngVal) & ", "
        s = s & ToStr_Param(.DblVal) & ", "
        s = s & ToStr_Param(.StrVal) & ")"
    End With
    LStr.Item(i) = s
End Sub

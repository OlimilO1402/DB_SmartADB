VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
' typedef  enum tagVARENUM
' {
Public Enum EDataType 'Public Enum Extended VbVarType
    'Values like vbEmpty, vbNull, vbError are missing here
    'beacuse these types can't be stored in an array anyway
    vbInteger = VbVarType.vbInteger    '   2   '   VT_I2          = 0x0002,
    vbLong = VbVarType.vbLong          '   3   '   VT_I4          = 0x0003,
    vbSingle = VbVarType.vbSingle      '   4   '   VT_R4          = 0x0004,
    vbDouble = VbVarType.vbDouble      '   5   '   VT_R8          = 0x0005,
    vbCurrency = VbVarType.vbCurrency  '   6   '   VT_CY          = 0x0006,
    vbDate = VbVarType.vbDate          '   7   '   VT_DATE        = 0x0007,
    vbString = VbVarType.vbString      '   8   '   VT_BSTR        = 0x0008,
    vbObject = VbVarType.vbObject      '   9   '   VT_DISPATCH    = 0x0009, 'IUnknown+IDispatch
    vbBoolean = VbVarType.vbBoolean    ' &HB&  '   VT_BOOL        = 0x000B,
    vbVariant = VbVarType.vbVariant    ' &HC&  '   VT_VARIANT     = 0x000C,
    'Of_Unknown, 'Of_DataObject = VbVarType.vbDataObject ' &HD& '   VT_UNKNOWN = 0x000D, 'IUnknown
    vbDecimal = VbVarType.vbDecimal    ' &HE&  '   VT_DECIMAL     = 0x000E,
                                        ' &HF&  ?
    vbByte = VbVarType.vbByte          ' &H10& '   VT_I1          = 0x0010,
    'vbUByte = &H11                     ' &H11& '   VT_UI1         = 0x0011,
    'vbUInteger = &H12                  ' &H12& '   VT_UI2         = 0x0012,
    'vbULong = &H13                     ' &H13& '   VT_UI4         = 0x0013,
    'vbLongLong = &H14                  ' &H14& '   VT_I8          = 0x0014,
    'vbULongLong = &H15                 ' &H15& '   VT_UI8         = 0x0015,
    'vbInt = &H16                       ' &H16& '   VT_INT         = 0x0016,
    'vbUInt = &H17                      ' &H17& '   VT_UINT        = 0x0017,
    'vbVoid = &H18                      ' &H18& '   VT_VOID        = 0x0018,
    'vbHResult = &H19                   ' &H19& '   VT_HRESULT     = 0x0019,
    'vbPtr = &H1A                       ' &H1A& '   VT_PTR         = 0x001A,
    'vbSafeArray = &H1B                 ' &H1B& '   VT_SAFEARRAY   = 0x001B,
    'vbCArray = &H1C                    ' &H1C& '   VT_CARRAY      = 0x001C,
    'vbUserdefined = &H1D               ' &H1D& '   VT_USERDEFINED = 0x001D,
    'vbLPStr = &H1E                     ' &H1E& '   VT_LPSTR       = 0x001E,
    'vbLPWStr = &H1F                    ' &H1F& '   VT_LPWSTR      = 0x001F,
    vbWChar '= &H27&                    ' &H27&
    vbUserDefinedType = VbVarType.vbUserDefinedType '   VT_RECORD = 0x0024,
    'As_Array = &H2000&                  '&H2000& '  VT_ARRAY       = 0x2000, 'Arr() As Integer
    vbByRef = &H4000&                  '&H4000& '  VT_BYREF       = 0x4000
    'I want compatiblity with ByRef to another List
    'this could be useful with Arrays as UdT with pointers
    'but that complicates things quite a bit
End Enum
' } VARENUM;
'Class List
'1432 Zeilen geballte Listen-Power für VB6, alles in einer Klasse.
'Mit einem Hauch von Generics und LINQ
'Funktionalität von Collection, List, Enumerable, Stack, Queue und Dictionary mehr braucht man nicht.
Dim m_Arr
Dim m_DataType As EDataType
Dim m_Capacity As Long '= 32
Dim m_GrowRate As Single 'growfactor
Dim m_GrowSize As Long   'chunksize
Dim m_bytLen   As Long   'die Länge eines Elements in Bytes
Dim m_Count    As Long
Dim m_IsPtr    As Boolean
Dim m_Index    As Long 'works as ListIndex for a selected item
Dim m_IsHashed As Boolean
Dim m_HashI    As Collection
'for lambdas: CallByName(m_LambObj, m_LambFunc)
Dim m_LambObj  As Object
Dim m_LambFunc As String
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLen As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal BytLen As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long

Private Sub Class_Initialize()
    'm_DataType = vbObject
    m_Capacity = 32
    m_GrowRate = 2
    'InitGenericArray m_DataType
    'ReDim m_Arr(0 To m_Capacity - 1)
    m_Index = -1
End Sub

Private Sub Class_Terminate()
    If (m_DataType And vbByRef) = vbByRef Then RtlZeroMemory ByVal VarPtr(m_Arr), 16
End Sub

Friend Sub New_(Of_T As EDataType, _
                Optional ArrColStrTypList, _
                Optional ByVal IsHashed As Boolean = False, _
                Optional ByVal Capacity As Long = 32, _
                Optional ByVal GrowRate As Single = 2, _
                Optional ByVal GrowSize As Long = 0) ', _
'                Optional ByVal ReadOnly As Boolean = False)
    m_DataType = Of_T
    If (m_DataType And vbUserDefinedType) = vbUserDefinedType Then
        If Not IsMissing(ArrColStrTypList) Then m_bytLen = LenB(ArrColStrTypList)
    End If
    m_Capacity = Capacity
    InitGenericArray m_DataType
    m_IsHashed = IsHashed
    If m_IsHashed Then Set m_HashI = New Collection
    
    'chunksize darf nicht negativ sein!
    m_GrowSize = IIf(GrowSize < 0, 0, GrowSize)
    ' Wenn ZuwachsGröße = 0 dann muss sichergestellt werden, dass die ZuwachsRate > 1 ist.
    m_GrowRate = IIf(m_GrowSize = 0, IIf(GrowRate <= 1, 2, GrowRate), IIf(GrowRate < 1, 1, GrowRate))
    'Reihenfolge wichtig, so lassen!
    If Not IsMissing(ArrColStrTypList) Then Me.AddRange ArrColStrTypList
End Sub

Friend Sub NewC(other As List)
    'Copy-Constructor wird von Clone aufgerufen
    With other
        Me.New_ .DataType, other, .IsHashed
    End With
End Sub

Private Sub InitGenericArray(T As EDataType)
    Dim u As Long: u = m_Capacity + 1
    Select Case T
    Case vbInteger:    ReDim m_Arr(0 To u) As Integer:   m_bytLen = 2
    Case vbLong:       ReDim m_Arr(0 To u) As Long:      m_bytLen = 4
    Case vbSingle:     ReDim m_Arr(0 To u) As Single:    m_bytLen = 4
    Case vbDouble:     ReDim m_Arr(0 To u) As Double:    m_bytLen = 8
    Case vbCurrency:   ReDim m_Arr(0 To u) As Currency:  m_bytLen = 8
    Case vbDate:       ReDim m_Arr(0 To u) As Date:      m_bytLen = 8
    Case vbString:     ReDim m_Arr(0 To u) As String:    m_bytLen = 4
    Case vbObject:     ReDim m_Arr(0 To u) As Object:    m_bytLen = 4
    Case vbBoolean:    ReDim m_Arr(0 To u) As Boolean:   m_bytLen = 2
    Case vbVariant:    ReDim m_Arr(0 To u) As Variant:   m_bytLen = 16
    Case vbDataObject: ReDim m_Arr(0 To u) As Variant:   m_bytLen = 16
    Case vbDecimal:    ReDim m_Arr(0 To u) As Variant:   m_bytLen = 16
    Case vbByte:       ReDim m_Arr(0 To u) As Byte:      m_bytLen = 1
    Case vbWChar:      ReDim m_Arr(0 To u) As Integer:   m_bytLen = 2
    Case vbUserDefinedType
                       'entweder so:
                       'ReDim m_Arr(0 To u * m_bytLen) As Byte
                       'oder so:
                       ReDim m_Arr(0 To u)
    'Case vbUserDefinedType '
    'Falls kein Type aus einer AXdll oder Typlib, dann muss mit "m_Arr() As MyType" genauer spezifiziert werden
    'wait for Add
    'Case vbArray:
    Case Else 'Fehlermeldung ausgeben
        MsgBox "Datentyp wird nicht unterstützt: " & EDataType_ToStr(T)
    End Select
End Sub

'Private Function SizeOf(t As EDataType) As Byte
'    Select Case t
'    Case vbInteger:    SizeOf = 2
'    Case vbLong:       SizeOf = 4
'    Case vbSingle:     SizeOf = 4
'    Case vbDouble:     SizeOf = 8
'    Case vbCurrency:   SizeOf = 8
'    Case vbDate:       SizeOf = 8
'    Case vbString:     SizeOf = 4
'    Case vbObject:     SizeOf = 4
'    Case vbBoolean:    SizeOf = 2
'    Case vbVariant:    SizeOf = 16
'    'Case Of_DataObject: SizeOf = 16
'    Case vbDecimal:    SizeOf = 16
'    Case vbByte:       SizeOf = 1
'    Case vbArray:      SizeOf = 16
'    Case vbWChar:      SizeOf = 2
'    Case Else: MsgBox "Datatype is not supported: " & EDataType_ToStr(t)
'    End Select
'End Function

Public Property Get DataType() As EDataType
    DataType = m_DataType
End Property

Friend Function GetType() As String
    GetType = TypeName(Me) & "(" & VarType2_ToStr(m_Arr) & ")"
End Function

'Siehe auch Extra Projekt in der diese Funktion ausgiebig getestet wird
Private Function EDataType_ToStr(T As EDataType) As String
    Dim s As String
    Select Case T
    Case vbInteger:    s = "Integer"    ' vbInteger    = 2
    Case vbLong:       s = "Long"       ' vbLong       = 3
    Case vbSingle:     s = "Single"     ' vbSingle     = 4
    Case vbDouble:     s = "Double"     ' vbDouble     = 5
    Case vbCurrency:   s = "Currency"   ' vbCurrency   = 6
    Case vbDate:       s = "Date"       ' vbDate       = 7
    Case vbString:     s = "String"     ' vbString     = 8
    Case vbObject:     s = "Object"     ' vbObject     = 9
    Case vbBoolean:    s = "Boolean"    ' vbBoolean    = 11
    Case vbVariant:    s = "Variant"    ' vbVariant    = 12
    Case vbDataObject: s = "DataObject" ' vbDataObject = 13
    Case vbDecimal:    s = "Decimal"    ' vbDecimal    = 14
    '15, 16?
    Case vbByte:       s = "Byte"       ' vbByte       = 17 (&H11)
    Case vbUserDefinedType: s = "UserDefinedType" 'vbUserDefinedType = 36 (&H24)
    Case vbWChar:      s = "WChar"
    'Case vbArray:      s = "Array"      ' vbArray      = 8192 (&H2000)
    Case Else:
        If T And vbByRef Then s = "ByRef "
        If T And vbArray Then s = "Array As "
    End Select
    EDataType_ToStr = s
End Function

Private Function VarType2_ToStr(var) As String
    On Error Resume Next
    Dim vt As EDataType: vt = VarType(var)
    Dim s As String: s = EDataType_ToStr(vt)
    If vt And vbArray Then
        vt = vt - vbArray
        s = s & EDataType_ToStr(vt)
        If vt = vbVariant Then
            s = s & "(" & VarType2_ToStr(var(LBound(var))) & ")"
        ElseIf vt = vbObject Then
            s = s & "(As " & TypeName(var(LBound(var))) & ")"
        End If
    ElseIf vt = vbObject Then
        s = s & "(As " & TypeName(var) & ")"
    End If
    VarType2_ToStr = s
    On Error GoTo 0
End Function

Public Function Add(avar)
    ' Fügt am Ende der Liste ein Objekt hinzu
    If IsMissing(avar) Then Exit Function
    If m_DataType = vbEmpty Then
        Dim dt As EDataType: dt = VarType(avar)
        If (dt = vbUserDefinedType And m_bytLen = 0) Then m_bytLen = LenB(avar)
        m_DataType = dt
        InitGenericArray m_DataType
    End If
    If m_IsHashed Then
        If Not HashAddKey(KeyFromVar(avar), m_Count) Then Exit Function
    End If
    Grow
    If m_DataType = vbObject Then
        Set m_Arr(m_Count) = avar
    Else
        If m_DataType = vbUserDefinedType Then
            'per cpymem die Daten hineinkopieren
            'dim pData as Long: pdata =
            'RtlMoveMemory ByVal DataPtrI(m_Count), VarPtr(avar), m_bytLen
            'm_Arr (m_bytLen * m_Count)
            m_Arr(m_Count) = avar
        Else
            If VarType(avar) = vbObject Then
                Set m_Arr(m_Count) = avar
            Else
                m_Arr(m_Count) = avar
            End If
        End If
    End If
    m_Count = m_Count + 1
End Function
Private Function KeyFromVar(avar) As String
    Dim k As String
    'If m_IsHashed Then
        Select Case m_DataType
        Case vbObject
            If avar Is Nothing Then Exit Function
            k = avar.Key ' <-- Class must have property Key
        Case vbUserDefinedType
            k = avar.Key ' <-- Type  must have property Key
        Case Else
            'OK hier könnte man noch überprüfen, ist
            'aVar ein Array (Tuple) aus zwei Elementen, dann Vereinbarung, key-value der erste ein String, der zweite Value
            'oder man macht eine zweite function AddKV
            k = CStr(avar)
            'hier wird der Index gespeichert und der Key
            'besser key + vbnullchar+cstr(Index)
            'm_HashI.Add Key & vbNullChar & CStr(m_Count), Key
            'über den Key bekommt man immer den Key und den Index
            'über den Index bekommt man das Element und so auch den Key
        End Select
    'Else
    '
    'End If
    KeyFromVar = k
End Function
Private Function HashAddKey(aKey As String, Index As Long) As Boolean
    If Len(aKey) Then
        If Contains(aKey) Then Exit Function
        'den Index mit dem Key integrieren?
        m_HashI.Add aKey & vbNullChar & CStr(Index), aKey
    End If
    HashAddKey = True
End Function
'Public Function AddKV(Key, Value)
'    '???
'End Function
Private Sub Grow()
    ' Vergrößert das Array falls erforderlich.
    ' Am Ende des Arrays muss immer ein leeres Element bleiben, für Insert!
    ' Das Array soll entweder über die ZuwachsRate (growfactor) oder über die ZuwachsGröße (chunksize) oder über Beides vergrößert werden können
    Dim ub As Long: ub = UBound(m_Arr)
    If ub <= m_Count Then
        If m_DataType = vbUserDefinedType Then
            ReDim Preserve m_Arr(0 To (m_GrowRate * ub + m_GrowSize) * m_bytLen)
        Else
            ReDim Preserve m_Arr(0 To m_GrowRate * ub + m_GrowSize)
        End If
        'Debug.Print "ReDim Preserve m_Arr(0 To " & m_GrowRate * ub + m_GrowSize & ")"
    End If
End Sub

Private Sub EnsureCapacity(ByVal aCount As Long)
    'stellt sicher dass die Kapazität ausreicht für die angegebene GesamtAnzahl an Elementen
    Dim ub As Long: ub = UBound(m_Arr)
    If ub <= aCount Then
        'If m_DataType = vbUserDefinedType Then
        '    ReDim Preserve m_Arr(0 To aCount * m_bytLen)
        'Else
            ReDim Preserve m_Arr(0 To aCount)
        'End If
    End If
End Sub
Public Property Get GrowRate() As Single
    GrowRate = m_GrowRate
End Property
Public Property Let GrowRate(ByVal Value As Single)
    m_GrowRate = Value
End Property

Public Property Get GrowSize() As Long
    GrowSize = m_GrowSize
End Property
Public Property Let GrowSize(ByVal Value As Long)
    m_GrowSize = Value
End Property


'Private Sub GetKeyValue(avar, ByRef nam_out As String, value)
'    If VarType(avar) And vbArray <> vbArray Then Exit Sub
'    If (UBound(avar) - LBound(avar) + 1) <> 2 Then Exit Sub
'
'End Sub
Public Sub AddRange(aArr) 'As List
    ' Fügt die Elemente der angegebenen Auflistung der List hinzu
    If IsEmpty(aArr) Then Exit Sub
    Dim i As Long, v
    If TypeOf aArr Is List Then
        For Each v In aArr.GetEnumerator
            Me.Add v
        Next
    Else
        If m_DataType = vbWChar Then
            'für String
            If VarType(aArr) = vbString Then
                Dim bl As Long: bl = LenB(aArr)
                Dim l As Long: l = bl / 2
                If m_IsHashed Then
                    For i = 1 To l
                        Me.Add AscW(Mid$(aArr, i, 1))
                    Next
                Else
                    EnsureCapacity m_Count + l
                    Dim pDst As Long: pDst = DataPtrI(m_Count)
                    RtlMoveMemory ByVal pDst, ByVal StrPtr(aArr), bl
                    m_Count = m_Count + l
                End If
            End If
        ElseIf m_DataType = vbUserDefinedType Then
            For i = 0 To UBound(aArr)
                v = aArr(i)
                Me.Add v
            Next
        Else
            'für Collection oder Array
            For Each v In aArr
                Me.Add v
            Next
        End If
    End If
End Sub

Public Property Get Capacity() As Long
    'Ruft die Gesamtzahl der Elemente ab, die die interne Datenstruktur ohne Änderung der Größe aufnehmen kann, oder legt diese Anzahl fest.
    Capacity = UBound(m_Arr)
End Property
Public Property Let Capacity(ByVal Value As Long)
    'das letzte Element wird nie benützt, für Insert
    ReDim Preserve m_Arr(0 To Value)
End Property

Public Sub Clear()
    'Entfernt alle Elemente aus der List
    m_DataType = vbEmpty
    m_Count = 0
    m_GrowRate = 0
    m_Capacity = 32
    m_bytLen = 0
    'm_Capacity = 0
    Set m_HashI = Nothing
    m_IsHashed = False
    Class_Initialize
End Sub

Public Function Clone() As List
    'erzeugt eine flache Kopie der Liste und gibt sie zurück
    Set Clone = New List: Clone.NewC Me
End Function

Public Function ContainsEqualObj(obj As Object) As Boolean
    Dim i As Long
    For i = 0 To m_Count - 1
        If m_Arr(i).Equals(obj) Then
            ContainsEqualObj = True: Exit Function
        End If
    Next
End Function
Public Function ContainsKey(Key As String) As Boolean
    'Bestimmt, ob sich ein Element in der Liste befindet.
    On Error Resume Next
    If m_IsHashed Then
        If IsEmpty(m_HashI(Key)) Then: 'DoNothing
        ContainsKey = (Err.Number = 0)
    Else
        Dim i As Long
        If m_DataType = vbObject Then
            For i = 0 To m_Count - 1
                If m_Arr(i).Key = Key Then ContainsKey = True: Exit Function
            Next
        Else
            For i = 0 To m_Count - 1
                If CStr(m_Arr(i)) = Key Then ContainsKey = True: Exit Function
            Next
        End If
    End If
    On Error GoTo 0
End Function
Public Function Contains(obj, Optional ByVal ObjUseKey As Boolean = False) As Boolean
    'Bestimmt, ob sich ein Element in der Liste befindet.
    On Error Resume Next
    Dim Key As String
    If m_IsHashed Then
        If m_DataType = vbObject Then Key = obj.Key Else Key = obj
        If IsEmpty(m_HashI(Key)) Then: 'DoNothing
        Contains = (Err.Number = 0)
    Else
        Dim i As Long
        Select Case m_DataType
        Case vbObject
            'entweder über Key oder über Gleichheit der Objekte
            If ObjUseKey Then
                For i = 0 To m_Count - 1
                    If m_Arr(i).Key = Key Then Contains = True: Exit Function
                Next
            Else
                Dim pObj As Long: pObj = ObjPtr(obj)
                For i = 0 To m_Count - 1
                    If ObjPtr(m_Arr(i)) = pObj Then Contains = True: Exit Function
                Next
            End If
        Case vbUserDefinedType
            'hier eine funktion aufrufen die zwei udts miteinander vergleicht
        Case Else
            'Key = Obj
            For i = 0 To m_Count - 1
                'If CStr(m_Arr(i)) = Key Then Contains = True: Exit Function
                If m_Arr(i) = obj Then Contains = True: Exit Function
            Next
        End Select
    End If
    On Error GoTo 0
End Function

Public Property Get Count() As Long
    'Ruft die Anzahl der Elemente ab, die in der List enthalten sind.
    Count = m_Count
End Property

Public Function GetEnumerator() 'As Variant ' 'As IUnknown
    'Prozedur-ID = -4
    'Gibt einen Enumerator zurück, der die Liste durchläuft
    Me.Trim
    GetEnumerator = m_Arr 'm_Arr() '.[_NewEnum]
End Function

Public Property Get UUBound() As Long
    'ähnlich wie Capacity
    'gibt den Upper Bound, die oberste Schranke, den größten Index des internen Arrays zurück
    UUBound = UBound(m_Arr)
End Property

Public Property Get IIsEmpty() As Boolean
    'gibt zurück ob die Liste leer ist
    IIsEmpty = m_Count = 0
End Property
Public Property Get IsHashed() As Boolean
    'gibt zurück ob in der Liste die Hashing-Funktionalität verwendet wird.
    IsHashed = m_IsHashed
End Property

Public Property Let Item(ByVal Index As Long, avar)
    If m_Count <= Index Then Exit Property
    Select Case m_DataType
    Case vbObject:  m_Arr(Index).Parse avar '<- class must have function Parse
    Case vbVariant: Select Case VarType(m_Arr(Index))
                    Case vbObject: m_Arr(Index).Parse avar
                    Case Else:     m_Arr(Index) = avar
                    End Select
    Case Else:      m_Arr(Index) = avar
    End Select
End Property
Public Property Set Item(ByVal Index As Long, avar)
    If m_Count <= Index Then Exit Property
    Set m_Arr(Index) = avar
End Property
Public Property Get Item(ByVal Index As Long)
    'Standard-Property
    'gibt ein Element am angegebenen Index zurück
    If m_Count <= Index Then Exit Property
    Select Case m_DataType
    Case vbObject:  Set Item = m_Arr(Index)
    'Case vbUserDefinedType: -> Case Else
    Case vbVariant: If IsObject(m_Arr(Index)) Then Set Item = m_Arr(Index) Else Item = m_Arr(Index)
    Case Else:      Item = m_Arr(Index)
    End Select
End Property

Public Property Get ItemToStr(ByVal Index As Long) As String
    If m_Count <= Index Then Exit Property
    Dim s As String
    Select Case m_DataType
    Case vbObject:  s = m_Arr(Index).Key
    Case vbUserDefinedType
    Case vbVariant: Select Case VarType(m_Arr(Index))
                    Case vbObject: s = m_Arr(Index).Key
                    Case Else:     s = m_Arr(Index)
                    End Select
    Case Else:      s = m_Arr(Index)
    End Select
    ItemToStr = s
End Property

Public Property Get ItemByKey(aKey As String)
    'liefert das Element mit dem angegebenen Key, nur mit IsHashed = True
    If Not m_IsHashed Then Exit Property
    If ContainsKey(aKey) Then
        Dim i As Long: i = m_HashI(aKey)
        Select Case m_DataType
        Case vbObject:  Set ItemByKey = m_Arr(i)
        Case vbVariant: If IsObject(m_Arr(i)) Then Set ItemByKey = m_Arr(i)
        Case Else:      ItemByKey = m_Arr(i)
        End Select
    End If
End Property

Public Property Get IndexOfKey(aKey As String) As Long
    'm_HashI.Add Key & vbNullChar & CStr(m_Count), Key
    
    'liefert den Index des ersten Vorkommens eines Elements mit dem angegebenen Key
    Dim i As Long
    If m_IsHashed Then
        If ContainsKey(aKey) Then
            Dim ki As String: ki = m_HashI(aKey)
            IndexOfKey = Split(ki, vbNullChar)(1)
        End If
    Else
        Select Case m_DataType
        Case vbObject:  For i = 0 To m_Count - 1
                            If aKey = m_Arr(i).Key Then IndexOfKey = i: Exit Property
                        Next
        Case vbString:  For i = 0 To m_Count - 1
                            If aKey = m_Arr(i) Then IndexOfKey = i: Exit Property
                        Next
        Case Else:      For i = 0 To m_Count - 1
                            If aKey = m_Arr(i) Then IndexOfKey = i: Exit Property
                        Next
        End Select
    End If
End Property

Public Property Get IndexOf(Value) As Long
    Dim i As Long, obj As Object
    IndexOf = -1 'first we assume its not in list
    If m_IsHashed Then
        Dim Key As String
        Select Case m_DataType
        Case vbObject: Set obj = Value: Key = obj.Key
        Case vbString: Key = Value
        Case Else:     Key = Value
        End Select
        If ContainsKey(Key) Then
            Dim ki As String: ki = m_HashI(Key)
            IndexOf = Split(ki, vbNullChar)(1)
        End If
    Else
        Select Case m_DataType
        Case vbObject:  'Set obj = Value: Dim pObj As Long: pObj = ObjPtr(obj)
                        'For i = 0 To m_Count - 1
                        '    If ObjPtr(m_Arr(i)) = pObj Then IndexOf = i: Exit Property
                        'Next
                        'Nein nicht ObjPtr sondern Equals verwenden
                        For i = 0 To m_Count - 1
                            If m_Arr(i).Equals(Value) Then IndexOf = i: Exit Property
                        Next
                        
        Case vbString:  For i = 0 To m_Count - 1
                            If m_Arr(i) = Value Then IndexOf = i: Exit Property
                        Next
        Case Else:      For i = 0 To m_Count - 1
                            If m_Arr(i) = Value Then IndexOf = i: Exit Property
                        Next
        End Select
    End If
End Property

'    If m_IsHashed Then
'        Select Case m_DataType
'        Case vbObject
'            If avar Is Nothing Then Exit Function
'            Key = avar.Key ' <-- Class must have property Key
'        Case vbUserDefinedType
'            Key = avar.Key ' <-- Type  must have property Key
'        Case Else
'            'OK hier könnte man noch überprüfen, ist
'            'aVar ein Array (Tuple) aus zwei Elementen, dann Vereinbarung, key-value der erste ein String, der zweite Value
'            'oder man macht eine zweite function AddKV
'            Key = CStr(avar)
'            'hier wird der Index gespeichert und der Key
'            'besser key + vbnullchar+cstr(Index)
'            'm_HashI.Add Key & vbNullChar & CStr(m_Count), Key
'        End Select
'        If Len(Key) Then
'            If Contains(Key) Then Exit Function
'            'den Index mit dem Key integrieren?
'            m_HashI.Add Key & vbNullChar & CStr(m_Count), Key
'        End If
'    End If

Private Property Get Key(ByVal Index As Long) As String
    Dim k As String
    If m_IsHashed Then
        Select Case m_DataType
        Case vbObject
            k = m_Arr(Index).Key ' <-- Class must have property Key
        Case vbUserDefinedType
            k = m_Arr(Index).Key ' <-- Type  must have property Key
        Case Else
            k = CStr(m_Arr(Index))
        End Select
        'm_HashI.Add Item, Key
        Key = k 'm_HashI(k)
    Else
        ' All objects must have property Key
        Select Case m_DataType
        Case vbObject: Key = m_Arr(Index).Key
        Case Else:     Key = m_Arr(Index)
        End Select
    End If
End Property

Public Function Find(LikeKey As String) As Long
    'gibt den Index zurück
    Dim i As Long, LKey As String
    For i = 0 To UBound(m_Arr)
        LKey = Key(i)
        'hier nicht Or verwenden, sonst wird tatsächlich jede Operation durchgeführt, selbst
        'dann wenn eine schon true liefert. OrElse wäre richtig aber gibts in VBC nicht.
        If LKey = LikeKey Then
            Find = i: Exit Function
        ElseIf LKey Like LikeKey Then
            Find = i: Exit Function
        ElseIf LikeKey Like LKey Then
            Find = i: Exit Function
        ElseIf InStr(1, LKey, LikeKey, vbTextCompare) Then
            Find = i: Exit Function
        ElseIf InStr(1, LikeKey, LKey, vbTextCompare) Then
            Find = i: Exit Function
        End If
    Next
End Function
'Function Equals() As Boolean
'Function Exists() As Boolean
'Function Find() 'As Variant
'Function FindAll() As List
'Function FindIndex() As Long
'Function FindLast() 'As Variant
'Function FindLastIndex() As Long

Public Function FindAll(LikeKey As String) As List
    'gibt eine Liste der gefundenen Elemente zurück
    Dim i As Long, LKey As String
    Set FindAll = New List: FindAll.New_ m_DataType, , m_IsHashed
    For i = 0 To UBound(m_Arr)
        LKey = Key(i)
        'hier nicht Or verwenden, sonst wird tatsächlich jede Operation durchgeführt, selbst
        'dann wenn eine schon true liefert. OrElse wäre richtig aber gibts in VBC nicht.
        If LKey = LikeKey Then
            FindAll.Add m_Arr(i)
        ElseIf LKey Like LikeKey Then
            FindAll.Add m_Arr(i)
        ElseIf LikeKey Like LKey Then
            FindAll.Add m_Arr(i)
        ElseIf InStr(1, LKey, LikeKey, vbTextCompare) Then
            FindAll.Add m_Arr(i)
        ElseIf InStr(1, LikeKey, LKey, vbTextCompare) Then
            FindAll.Add m_Arr(i)
        End If
    Next
End Function

Public Property Get SAPtr() As Long
    'liefert den Zeiger auf den SafeArray-Descriptor
    RtlMoveMemory SAPtr, ByVal VarPtr(m_Arr) + 8, 4
End Property

Public Property Get DataPtr() As Long
    'liefert den Zeiger auf das erste Datenelement
    RtlMoveMemory DataPtr, ByVal SAPtr + 12, 4
End Property

Private Property Get DataPtrI(ByVal Index As Long) As Long
    DataPtrI = DataPtr + Index * m_bytLen
End Property

Public Property Get ByteLength() As Long
    'returns the occupied memory in bytes
    ByteLength = m_bytLen * (UUBound + 1)
End Property

Public Property Let ListIndex(ByVal Value As Long)
    m_Index = Value
End Property
Public Property Get ListIndex() As Long
    ListIndex = m_Index
End Property
' Dim s As New Stack(Of String)
' s.Push ("Dings")   ' Fügt ein Objekt am Anfang des Stack(Of T) ein.
' s.Peek()           ' Gibt das Objekt am Anfang des Stack(Of T) zurück, ohne es zu entfernen.
' s.Pop()            ' Entfernt und gibt das Objekt am Anfang des Stack(Of T)
' 'Der Anfang von Stack ist selbstverständlich das Ende der Liste.
' List As Stack
' ein Stack ist ein LiFo = Last in First out
' wird meist verwendet um einen Tree aufzubauen
Public Sub SPush(avar)
    Me.Add avar
End Sub
Public Function SPeek()
    If m_DataType = vbObject Then
        Set SPeek = m_Arr(m_Count - 1)
    Else
        SPeek = m_Arr(m_Count - 1)
    End If
End Function
Public Function SPop()
    If m_DataType = vbObject Then
        Set SPop = m_Arr(m_Count - 1)
        Set m_Arr(m_Count - 1) = Nothing
    Else
        SPop = m_Arr(m_Count - 1)
        m_Arr(m_Count - 1) = Null
    End If
    m_Count = m_Count - 1
End Function

' Dim q As New Queue(Of String)
' q.Enqueue ("Dings") ' Fügt am Ende der Queue(Of T) ein Objekt hinzu.
' q.Peek()            ' Gibt das Objekt am Anfang der Queue zurück, ohne es zu entfernen.
' q.Dequeue()         ' Entfernt das Objekt am Anfang der Queue(Of T) und gibt es zurück.
' List As Queue
' eine Queue ist ein FiFo = First in First out
' Meist arbeiten zwei verschiedene Objekte an der Queue
' * eines fügt am Ende Werte hinzu
' * das andere verarbeitet am Anfang die Objekte, und entfent sie
Public Sub QEnqueue(avar)
    'Fügt am Ende der Queue ein Objekt hinzu
    Me.Add avar
End Sub
Public Function QPeek()
    'gibt das Objekt am Anfang der Queue zurück, ohne es zu entfernen.
    If m_DataType = vbObject Then
        Set QPeek = m_Arr(0)
    Else
        QPeek = m_Arr(0)
    End If
End Function
Public Function QDequeue()
    'entfernt das Objekt am Anfang der Queue und gibt es zurück
    If m_DataType = vbObject Then
        Set QDequeue = m_Arr(0)
        'Set m_Arr(0) = Nothing
    Else
        QDequeue = m_Arr(0)
        'm_Arr(0) = Null
    End If
    RemoveFirst
    'm_Count = m_Count - 1
End Function

'sodale jetzt noch ein Insert dann is die reine List.Generic-Klasse ohne LINQ fertig
Public Sub Insert(ByVal Index As Long, avar)
    If m_IsHashed Then
        'zuerst überprüfen ob schon ein ELement da ist
        If Contains(KeyFromVar(avar)) Then Exit Sub
    End If
    'If Index < 0 Then Exit Sub
    'If m_Count <= Index Then Exit Sub
    If Not (0 <= Index And Index <= m_Count) Then Exit Sub
    If Index = m_Count Then Me.Add avar: Exit Sub
    Grow
    Dim pSrc As Long: pSrc = DataPtrI(Index) ' VarPtr(m_Arr(Index)) 'pSrc - m_bytLen
    Dim pDst As Long: pDst = pSrc + m_bytLen 'VarPtr(m_Arr(Index + 1))
    Dim BLen As Long: BLen = m_bytLen * (m_Count - Index)
    RtlMoveMemory ByVal pDst, ByVal pSrc, BLen
    RtlZeroMemory ByVal pSrc, m_bytLen
    If m_DataType = vbObject Then
        Set m_Arr(Index) = avar
    Else
        m_Arr(Index) = avar
    End If
    m_Count = m_Count + 1
    If m_IsHashed Then
        ReInitIndices
    End If
End Sub

'dann kommt noch LINQ dran
Public Sub Remove(ByVal Index As Long)
    ' entfernt das Objekt am angegebenen Index
    If Index < 0 Then Exit Sub
    If m_Count <= Index Then Exit Sub
    If Index = m_Count - 1 Then RemoveLast: Exit Sub
    Dim pDst As Long: pDst = DataPtrI(Index) ' VarPtr(m_Arr(Index)) 'pSrc - m_bytLen
    Dim pSrc As Long: pSrc = pDst + m_bytLen 'VarPtr(m_Arr(Index + 1))
    Dim BLen As Long: BLen = m_bytLen * (m_Count - Index)
    'm_bytLen
    If m_DataType = vbObject Then
        'jetzt alle Elemente vom Ort Index+1 der Länge m_Count-Index
        Set m_Arr(Index) = Nothing
        'vielleicht auch Nullen mit RtlZeroMemory?
        'RtlZeroMemory ByVal pDst, m_bytLen
        'RtlMoveMemory ByVal pDst, ByVal pSrc, BLen
    ElseIf m_DataType = vbString Then
        m_Arr(Index) = vbNullString
        'RtlZeroMemory ByVal pDst, m_bytLen
    End If
    RtlZeroMemory ByVal pDst, m_bytLen
    RtlMoveMemory ByVal pDst, ByVal pSrc, BLen
    m_Count = m_Count - 1
    'Mist was machen bei Hashes?
    'egal, dann einfach neu machen
    If m_IsHashed Then ReInitIndices
End Sub

Public Sub RemoveObj(obj As Object)
    Dim i As Long: i = IndexOf(obj)
    'blöd, IndexOf muss natürlich -1 zurückgeben falls nicht drin
    If i < 0 Then Exit Sub
    Remove i
End Sub

Public Sub RemoveFirst()
    Remove 0
    'If m_IsHashed Then m_HashI.Remove m_Arr(0).Name
    'Set m_Arr(0) = Nothing
    'jetzt müßte man eigentlich alle Objekte ab diesem Index nach vorne kopieren
End Sub

Public Sub RemoveLast()
    If m_IsHashed Then
        Dim nam As String
        If m_DataType = vbObject Then nam = m_Arr(m_Count - 1).Name Else nam = m_Arr(m_Count - 1)
        m_HashI.Remove nam
    End If
    If m_DataType = vbObject Then
        Set m_Arr(m_Count - 1) = Nothing
    ElseIf m_DataType = vbString Then
        m_Arr(m_Count - 1) = vbNullString
    Else
        m_Arr(m_Count - 1) = 0
    End If
    m_Count = m_Count - 1
End Sub

 
'    If m_IsHashed Then
'        Select Case m_DataType
'        Case vbObject
'            If avar Is Nothing Then Exit Function
'            Key = avar.Key ' <-- Class must have property Key
'        Case vbUserDefinedType
'            Key = avar.Key ' <-- Type  must have property Key
'        Case Else
'            'OK hier könnte man noch überprüfen, ist
'            'aVar ein Array (Tuple) aus zwei Elementen, dann Vereinbarung, key-value der erste ein String, der zweite Value
'            'oder man macht eine zweite function AddKV
'            Key = CStr(avar)
'            'hier wird der Index gespeichert und der Key
'            'besser key + vbnullchar+cstr(Index)
'            'm_HashI.Add Key & vbNullChar & CStr(m_Count), Key
'        End Select
'        If Len(Key) Then
'            If Contains(Key) Then Exit Function
'            'den Index mit dem Key integrieren?
'            m_HashI.Add Key & vbNullChar & CStr(m_Count), Key
'        End If
'    End If

Sub ReInitIndices()
    Set m_HashI = New Collection
    Dim i As Long
    Dim Key As String
    Select Case m_DataType
    Case vbObject
                    For i = 0 To m_Count - 1
                        Key = m_Arr(i).Key ' <-- Class must have Ppoperty or function "Key"
                        m_HashI.Add Key & vbNullChar & CStr(i), Key
                    Next
    Case vbUserDefinedType
                    For i = 0 To m_Count - 1
                        Key = m_Arr(i).Key ' <-- Class must have Ppoperty or function "Key"
                        m_HashI.Add Key & vbNullChar & CStr(i), Key
                    Next
    Case vbVariant
                    For i = 0 To m_Count - 1
                        If IsObject(m_Arr(i)) Then
                            Key = m_Arr(i).Key ' <-- Class must have property or function "Key"
                        Else
                            Key = CStr(m_Arr(i))
                        End If
                        m_HashI.Add Key & vbNullChar & CStr(i), Key
                    Next
    Case Else:      For i = 0 To m_Count - 1
                        Key = CStr(m_Arr(i))
                        m_HashI.Add Key & vbNullChar & CStr(i), Key
                    Next
    End Select
End Sub

Public Sub Sort(Optional ComparerObj As Object, Optional lambdaFunc As String)
'Sub List(Of T).Sort()
'Sub List(Of T).Sort(comparison As Comparison(Of T)
'Sub List(Of T).Sort(comparer As IComparer(Of T))
'Sub List(Of T).Sort(index As Integer, count As Integer, comparer As IComparer(Of T))
    'Sortiert die Elemente in der gesamten List(Of T) mithilfe des Standardcomparers.
    'Sortiert die Liste mit einem jeweils nach Datentyp angepassten QuickSort
    If Not ComparerObj Is Nothing Then
        Set m_LambObj = ComparerObj
        m_LambFunc = lambdaFunc
    End If
    Select Case m_DataType
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        Call QuickSortVar(0, m_Count - 1)
    Case vbString
        Call QuickSortStr(0, m_Count - 1)
    Case vbObject
        Call QuickSortObj(0, m_Count - 1)
    End Select
    If m_IsHashed Then ReInitIndices
    Set m_LambObj = Nothing
    m_LambFunc = vbNullString
End Sub
Public Sub SortRev(Optional ComparerObj As Object, Optional lambdaFunc As String)
    If Not ComparerObj Is Nothing Then
        Set m_LambObj = ComparerObj
        m_LambFunc = lambdaFunc
    End If
    Select Case m_DataType
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        Call QuickSortVar(0, m_Count - 1)
    Case vbString
        Call QuickSortStr(0, m_Count - 1)
    Case vbObject
        Call QuickSortObj(0, m_Count - 1)
    End Select
    'If m_IsHashed Then ReInitIndices
    Reverse
    Set m_LambObj = Nothing
    m_LambFunc = vbNullString
End Sub
'Public Sub Sort(Optional ByVal i1 As Long = 0, Optional ByVal i2 As Long = m_Count - 1)
'
'    Call QuickSort(i1, i2)
'
'End Sub

' Die Rekursive datenunabhängige Methode QuickSort
Private Sub QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = divideVar(i1, i2)
        Call QuickSortVar(i1, T - 1)
        Call QuickSortVar(T + 1, i2)
    End If
End Sub
Private Function divideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareVar(p, j) < 0))
        If i < j Then Call SwapVar(i, j)
    Loop While (i < j)
    Call SwapVar(i, p)
    divideVar = i
End Function
Private Function CompareVar(i1 As Long, i2 As Long)
    CompareVar = m_Arr(i1) - m_Arr(i2)
End Function
Private Sub SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim tmp: tmp = m_Arr(i1)
    m_Arr(i1) = m_Arr(i2): m_Arr(i2) = tmp
End Sub

' Die Rekursive datenunabhängige Methode QuickSort
Private Sub QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = divideStr(i1, i2)
        Call QuickSortStr(i1, T - 1)
        Call QuickSortStr(T + 1, i2)
    End If
End Sub
Private Function divideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareStr(p, j) < 0))
        If i < j Then Call SwapStr(i, j)
    Loop While (i < j)
    Call SwapStr(i, p)
    divideStr = i
End Function
Private Function CompareStr(i1 As Long, i2 As Long)
    CompareStr = StrComp(m_Arr(i1), m_Arr(i2))
End Function
Private Sub SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim tmp: tmp = m_Arr(i1)
    m_Arr(i1) = m_Arr(i2): m_Arr(i2) = tmp
End Sub

' Die Rekursive datenunabhängige Methode QuickSort
Private Sub QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = divideObj(i1, i2)
        Call QuickSortObj(i1, T - 1)
        Call QuickSortObj(T + 1, i2)
    End If
End Sub
Private Function divideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareObj(p, j) < 0))
        If i < j Then Call SwapObj(i, j)
    Loop While (i < j)
    Call SwapObj(i, p)
    divideObj = i
End Function
Private Function CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    CompareObj = m_Arr(i1).Compare(m_Arr(i2))
End Function
Private Sub SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim tmp As Object: Set tmp = m_Arr(i1)
    Set m_Arr(i1) = m_Arr(i2): Set m_Arr(i2) = tmp
End Sub

Public Sub Swap(ByVal i1 As Long, ByVal i2 As Long)
    'vertauscht die beiden Elemente z.B. für MoveUp, MoveDown
    If m_DataType = vbObject Then
        SwapObj i1, i2
    ElseIf m_DataType = vbString Then
        SwapStr i1, i2
    Else
        SwapVar i1, i2
    End If
End Sub

Private Sub priv_Reverse()
    Dim i As Long
    If m_DataType = vbObject Then
        For i = 0 To (m_Count - 1) \ 2
            SwapObj i, m_Count - 1 - i
        Next
    ElseIf m_DataType = vbString Then
        For i = 0 To (m_Count - 1) \ 2
            SwapStr i, m_Count - 1 - i
        Next
    Else
        For i = 0 To (m_Count - 1) \ 2
            SwapVar i, m_Count - 1 - i
        Next
    End If
End Sub

Public Sub ToListbox(aLBorCB, Optional ByVal UseListIndex As Boolean = False)
    Dim i As Long
    aLBorCB.Clear
    For i = 0 To m_Count - 1
        aLBorCB.AddItem ItemToStr(i)
    Next
    'besser m_index zu 0 setzen
    If 0 < m_Count And m_Index < 0 Then m_Index = 0
    If UseListIndex And (m_Index >= 0) Then
        aLBorCB.ListIndex = m_Index
    End If
End Sub
Public Function ToArray()
    Dim Arr: Arr = m_Arr
    Trim 'ReDim Preserve Arr(0 To m_Count - 1)
    ToArray = Arr
End Function

Public Sub Trim()
    ReDim Preserve m_Arr(0 To m_Count - 1)
End Sub

' #################### '     Enumerable Erweiterungs-Funktionen     ' #################### '
' ########## '                           Filtern                              ' ########## '
'6 * Filtern:
'   Where, Distinct, Take, TakeWhile, Skip, SkipWhile
Public Function Where(lambObj As Object, predicateFunc As String) As List ', restype As VbVarType) As List
'Liefert eine Untermenge von Elementen zurück, die eine gegebene Bedingung erfüllen.
'1 v 2:
'<Erweiterung> Enumerable.Where(predicate As Func(Of T, Boolean)) As IEnumerable(Of T)
' Filtert eine Sequenz von Wertem basierend auf einem Prädikat.
' * predicate:  Eine Funktion, mit der jedes Element auf eine Bedingung überprüft wird.

'2 v 2:
'<Erweiterung> Enumerable.Where(predicate As Func(Of T, Integer, Boolean)) As IEnumerable(Of T)
' Filtert eine Sequenz von Werten basierend auf einem Prädikat. In der Logik der Prädikatfuntion wird der Index der einzelnen Elemente verwendet.
' * predicate: Eine Funktion, zum Überprüfen jedes Quellelements auf eine Bedingung. Der zweite Parameter der Funktion stellt den Index des Quellelements dar.

    Dim i As Long ', v
    Set Where = New List: Where.New_ m_DataType
    For i = 0 To m_Count - 1
        'If m_DataType = vbObject Then Set v = m_Arr(i) Else v = m_Arr(i)
        If CallByName(lambObj, predicateFunc, VbMethod, (m_Arr(i))) Then  ', i) Then
            Where.Add m_Arr(i) 'CallByName(lambobj, predicateFunc, VbMethod, v, i)
        End If
    Next
End Function

Public Function Distinct(lambObj As Object, comparerFunc As String) As List
'Liefert eine Collection ohne Duplicate zurück
'1 v 2:
'<Erweiterung> Enumerable.Distinct() As IEnumerable(Of T)
' Gibt mithilfe des Standardgleichheitsvergleichs zum Vergleichen von Werten unterschiedliche Elemente aus einer Sequenz zurück.

'2 v 2:
'<Erweiterung> Enumerable.Distinct(comparer as IEqualityComparer(Of T)) As IEnumerable(Of T)
' Gibt unterschiedliche Elemente aus einer Sequenz mithilfe eines angegebenen IEqualityComparer(Of In T) zum Vergleichen von Werten  zurück.
' comparer: Ein IEqualityComparer(Of In T) zum Vergleichen von Werten.

    Dim i As Long ', v
    Set Distinct = New List: Distinct.New_ m_DataType, , True
    For i = 0 To m_Count - 1
        'If m_DataType = vbObject Then Set v = m_Arr(i) Else v = m_Arr(i)
        If CallByName(lambObj, comparerFunc, VbMethod, (m_Arr(i))) Then
            Distinct.Add m_Arr(i)
        End If
    Next
End Function

Public Function Take(ByVal aCount As Long) As List
'<Erweiterung> Enumerable.Take(count As Integer) As IEnumerable(Of T)
' Gibt eine angegebene Anzahl von zusammenhängenden Elementen ab dem Anfang einer Sequenz zurück.
' * count: Die Anzahl der zurückgegebenen Elemente.
    Dim i As Long ', v
    Set Take = New List: Take.New_ m_DataType
    For i = 0 To aCount - 1
        Take.Add m_Arr(i)
    Next
End Function

Public Function TakeWhile(lambObj As Object, predicateFunc As String) As List
'1 v 2:
'<Erweiterung> Enumerable.TakeWhile(predicate As Func(Of T, Boolean)) As IEnumerable(Of T)
' Gibt Elemente aus einer Sequenz zurück, solange eine angegebene Bedingung true ist.
' * predicate: Eine Funktion, mit der jedes Element auf eine Bedingung überprüft wird.

'2 v 2:
'<Erweiterung> Enumerable.TakeWhile(predicate As Func(Of T, Integer, Boolean)) As IEnumerable(Of T)
' Gibt Elemente aus einer Sequenz zurück, solange eine angegebene Bedingung true ist.
' In der Logik der Prädikatfunktion wird der Index des Elements verwendet.
' * predicate: Eine Funktion zum Überprüfen jedes Quellelements auf eine Bedingung.
'              Der zweite Parameter der Funktion stellt den Index des Quellelements dar.

    Dim i As Long ', v
    Set TakeWhile = New List: TakeWhile.New_ m_DataType
    For i = 0 To m_Count - 1
        If CallByName(lambObj, predicateFunc, VbMethod, (m_Arr(i)), i) Then
            TakeWhile.Add m_Arr(i)
        End If
    Next
End Function

Public Function Skip(ByVal aCount As Long) As List
'<Erweiterung> Enumerable.Skip(count As Integer) As IEnumerable(Of T)
' Umgeht eine festgelegte Anzahl von Elementen in einer Sequenz und gibt dann die übrigen Elemente zurück.
' * count: Die Anzahl der Elemente, die übersprungen werden sollen, bevor die übrigen Elemente zurückgegeben werden.
    Dim i As Long ', v
    Set Skip = New List: Skip.New_ m_DataType
    For i = aCount - 1 To m_Count - 1
        'If CallByName(lambobj, predicateFunc, VbMethod, m_Arr(i), i) Then
        Skip.Add m_Arr(i)
        'End If
    Next
End Function

Public Function SkipWhile(lambObj As Object, predicateFunc As String) As List
'1 v 2:
'<Erweiterung> Enumerable.SkipWhile(predicate As Func(Of T, Boolean) As IEnumerable(Of T)
' Umgeht Elemente in einer Sequenz, solange eine angegebene Bedingung true ist, und gibt dann die übrigen Elemente zurück.
' * predicate: Eine Funktion mit der jedes Element auf eine Bedingung überprüft wird.

'2 v 2:
'<Erweiterung> Enumerable.SkipWhile(predicate As Func(Of T, Integer, Boolean) As IEnumerable(Of T)
' Umgeht Elemente in einer Sequenz, solange eine angegebene Bedingung true ist, und gibt dann die übrigen Elemente zurück.
'In der Logik der Prädikatfunktion wird der Index des Elements verwendet.
' * predicate: Eine Funktion zum Überprüfen jedes Quellelelemnts auf eine Bedingung.
'              Der zweite Parameter der Funktion stellt den Index des Quellelements dar.

    Dim i As Long ', v
    Set SkipWhile = New List: SkipWhile.New_ m_DataType
    'For i = aCount - 1 To m_Count - 1
    For i = 0 To m_Count - 1
        If CallByName(lambObj, predicateFunc, VbMethod, (m_Arr(i)), i) Then
            SkipWhile.Add m_Arr(i)
        End If
    Next
End Function

' ########## '                         Projizieren                            ' ########## '
'2 * Projizieren:
'   Select, SelectMany
Public Function SSelect(lambObj As Object, predicateFunc As String, Of_TResult As EDataType) As List
'1 v 2:
'<Erweiterung> Enumerable.Select(Of TResult)(selector As Func(Of T, TResult)) As IEnumerable(Of T)
' Projiziert jedes Element einer Sequenz in ein neues Format.
' * selector: Eine Transformationsfunktion die auf jedes Element angewendet werden soll.

'2 v 2:
'<Erweiterung> Enumerable.Select(Of TResult)(selector As Func(Of T, Integer, TResult)) As IEnumerable(Of T)
' Projiziert jedes Element einer Sequenz in ein neues Format durch den Index des Elements integriert wird.
' * selector: Eine Transformaierenfunktion, die auf jedes Quellelement angewendet werden soll.
'              Der zweite Parameter der Funktion stellt den Index des Quellelements dar.
    Dim i As Long ', v
    Set SSelect = New List: SSelect.New_ Of_TResult
    For i = 0 To m_Count - 1
        'If m_DataType = vbObject Then Set v = m_Arr(i) Else v = m_Arr(i)
        SSelect.Add CallByName(lambObj, predicateFunc, VbMethod, (m_Arr(i)))
    Next
End Function

Public Function SelectMany(Of_TResult As VbVarType, lambObj As Object, predicateFunc As String) As List
'1 v 4:
'<Erweiterung> Enumerable.SelectMany(Of TResult)(selector As Func(Of T, IEnumerable(Of TResult))) As IEnumerable(Of TResult)
' Projiziert jedes Element einer Sequenz in ein IEnumerable(Of Out T), und fasst die resultierenden Sequenzen in einer einzigen Sequenz zusammen.
' * selector: Eine Transformierenfunktion, die auf jedes Element angewendet werden soll.

'2 v 4:
'<Erweiterung> Enumerable.SelectMany(Of TResult)(selector As Func(Of T, Integer, IEnumerable(Of TResult))) As IEnumerable(Of TResult)
' Projiziert jedes Element einer Sequenz in ein IEnumerable(Of Out T), und fasst die resultierenden Sequenzen in einer einzigen Sequenz zusammen.
' Der Index jedes Quellelements wird im projizierten Format des jweiligen Elements verwendet.
' * selector: Eine Transformierenfunktion, die auf jedes Element angewendet werden soll.
'              Der zweite Parameter der Funktion stellt den Index des Quellelements dar.

'3 v 4:
'<Erweiterung> Enumerable.SelectMany(Of TCollection, TResult)(collectionSelector As Func(Of T, IEnumerable(Of TCollection)), resultselector As Func(Of T, TCollection, TResult)) As IEnumerable(Of TResult)
' Projiziert jedes Element einer Sequenz in ein IEnumerable(Of Out T), fasst die resultierenden Sequenzen in einer einzigen Sequenz zusammen und ruft für jedes Element in dieser Sequenz eine Ergebnisauswahlfuntkion auf.
' * collectionSelector: Eine Transformierenfunktion, die auf jedes Element der Eingabesequenz angewendet werden soll.

'4 v 4:
'<Erweiterung> Enumerable.SelectMany(Of TCollection, TResult)(collectionSelector As Func(Of T, Integer, IEnumerable(Of TCollection)), resultselector As Func(Of T, TCollection, TResult)) As IEnumerable(Of TResult)
' Projiziert jedes Element einer Sequenz in ein IEnumerable(Of Out T), fasst die resultierenden Sequenzen in einer einzigen Sequenz zusammen und ruft für jedes Element in dieser Sequenz eine Ergebnisauswahlfuntkion auf.
' Der Index jedes Quellelements wird im projizierten Zwischenformat des jweiligen Elements verwendet.
' * collectionSelector: Eine Transformierenfunktion, die auf jedes Quellelement angewendet werden soll.
'              Der zweite Parameter der Funktion stellt den Index des Quellelements dar.



End Function

' ########## '                         Verknüpfen                            ' ########## '
'2 * Verknüpfen:
'   Join, GroupJoin
Public Function Join(Of_TInner As EDataType, TKey As VbVarType, TResult As VbVarType, _
                     lambdas As Object, _
                     FncOuterKeySelector As String, _
                     FncInnerKeySelector As String, _
                     FncResultSelector As String, _
                     Optional Comparer As String) As List
'1 v 2:
' Enumerable.Join(Of TInner, TKey, TResult)(inner As IEnumerable(Of TInner), outerKeySelector As Func(Of T, TKey), innerKeySelector As Func(Of Tinner, TKey), resultSelector As Func(Of T, TInner, TResult)) As IEnumerable(Of TResult)
' Korreliert die Elemente zweier Sequenzen basierend auf übereinstimmenden Schlüsseln.
' Schlüssel werden mihilfe des Standardgleichheitsvergleichs verglichen.
' * inner:            Die Sequenz, die mit der ersten Sequenz verknüpft werden soll.
' * outerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der ersten Sequenz.
' * innerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der zweiten Sequenz.
' * resultSelector:   Eine Funktion zum Erstellen eines Ergebniselements aus zwei übereinstimmenden Elementen.
' * opt Comparer:     ein IEqualityComparer(Of T) zum Hashen und Vergleichen von Schlüsseln.

'1 v 2:
' Enumerable.Join(Of TInner, TKey, TResult)(inner As IEnumerable(Of TInner), outerKeySelector As Func(Of T, TKey), innerKeySelector As Func(Of Tinner, TKey), resultSelector As Func(Of T, TInner, TResult), comparer As IEqualityComparer(Of TKey)) As IEnumerable(Of TResult)
' Korreliert die Elemente zweier Sequenzen basierend auf übereinstimmenden Schlüsseln.
' Schlüssel werden mithilfe eines angegebenen IEQualityComparer(Of In T) verglichen.
' * inner:            Die Sequenz, die mit der ersten Sequenz verknüpft werden soll.
' * outerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der ersten Sequenz.
' * innerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der zweiten Sequenz.
' * resultSelector:   Eine Funktion zum Erstellen eines Ergebniselements aus zwei übereinstimmenden Elementen.
End Function

Public Function GroupJoin(Of_TInner As EDataType, TKey As EDataType, TResult As EDataType, _
                          lambdas As Object, _
                          FncOuterKeySelector As String, _
                          FncInnerKeySelector As String, _
                          FncResultSelector As String, _
                          Optional Comparer As String) As List
'1 v 2:
'<Erweiterung> Enumerable.GroupJoin(Of TInner, TKey, TResult)(inner As IEnumerable(Of TInner), outerKeySelector As Func(Of T, TKey), innerKeySelector As Func(Of Tinner, TKey), resultSelector As Func(Of T, IEnumerable(Of TInner), TResult)) As IEnumerable(Of TResult)
' Korreliert die Elemente zweier Sequenzen basierend auf der Gleichheit der Schlüssel und gruppiert die Ergebnisse.
' Schlüssel werden mihilfe des Standardgleichheitsvergleichs verglichen.
' * inner:            Die Sequenz, die mit der ersten Sequenz verknüpft werden soll.
' * outerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der ersten Sequenz.
' * innerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der zweiten Sequenz.
' * resultSelector:   Eine Funktion zum Erstellen eines Ergebniselements anhand eines Elements aus der ersten Sequenz und einer Auzflistung von übereinstimmenden Elementen aus der zweiten Sequenz.

'2 v 2:
'<Erweiterung> Enumerable.GroupJoin(Of TInner, TKey, TResult)(inner As IEnumerable(Of TInner), outerKeySelector As Func(Of T, TKey), innerKeySelector As Func(Of Tinner, TKey), resultSelector As Func(Of T, IEnumerable(Of TInner), TResult)) As IEnumerable(Of TResult)
' Korreliert die Elemente zweier Sequenzen basierend auf der Gleichheit der Schlüssel und gruppiert die Ergebnisse.
' Schlüssel werden mihilfe eines angegebenen IEqualityComparer(Of In T) verglichen.
' * inner:            Die Sequenz, die mit der ersten Sequenz verknüpft werden soll.
' * outerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der ersten Sequenz.
' * innerKeySelector: Eine Funktion zum Extrahieren des Joinschlüssels aus jedem Element der zweiten Sequenz.
' * resultSelector:   Eine Funktion zum Erstellen eines Ergebniselements anhand eines Elements aus der ersten Sequenz und einer Auflistung von übereinstimmenden Elementen aus der zweiten Sequenz.
' * Comparer:         Ein IEqualityComparer(Of T) zum Hashen und Vergleichen von Schlüsseln.

End Function

' ########## '                         Sortieren                            ' ########## '
'5 * Sortieren:
'   OrderBy, OrderByDescending, ThenBy, ThenByDescending, Reverse
Public Function OrderBy(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
'1 v 2:
'<Erweiterung> Enumerable.OrderBy(Of TKey)(keySelector As Func(Of T, TKey)) As IOrderedEnumerable(Of T)
' Sortiert die Elemente einer Sequenz in aufsteigender Reihenfolge nach einem Schlüssel.
' * keySelector: Eine Funktion zum Extrahieren eines Schlüssels aus einem Element.

'2 v 2:
'<Erweiterung> Enumerable.OrderBy(Of TKey)(keySelector As Func(Of T, TKey), comparer As IComparer(Of TKey)) As IOrderedEnumerable(Of T)
' Sortiert die Elemente einer Sequenz in aufsteigender Reihenfolge mithilfe eines angegebenen Vergleichs.
' * keySelector: Eine Funktion zum Extrahieren eines Schlüssels aus einem Element.
    
  
End Function

Public Function OrderByDescending(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
'1 v 2:
'<Erweiterung> Enumerable.OrderByDescending(Of TKey)(keySelector As Func(Of T, TKey)) As IOrderedEnumerable(Of T)
' Sortiert die Elemente einer Sequenz in absteigender Reihenfolge nach einem Schlüssel.
' * keySelector: Eine Funktion zum Extrahieren eines Schlüssels aus einem Element.

'2 v 2:
'<Erweiterung> Enumerable.OrderByDescending(Of TKey)(keySelector As Func(Of T, TKey), comparer As IComparer(Of TKey)) As IOrderedEnumerable(Of T)
' Sortiert die Elemente einer Sequenz in absteigender Reihenfolge mithilfe eines angegebenen Vergleichs.
' * keySelector: Eine Funktion zum Extrahieren eines Schlüssels aus einem Element.

End Function

Public Function ThenBy(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List

End Function

Public Function ThenByDescending(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List

End Function

Public Function Reverse(Optional lambObj As Object = Nothing, Optional keySelectorFunc As String = "", Optional comparerFunc As String = "") As List
    If lambObj Is Nothing And Len(keySelectorFunc) = 0 And Len(comparerFunc) = 0 Then
        priv_Reverse
    Else
        
    End If
    If m_IsHashed Then ReInitIndices
End Function

' ########## '                         Gruppieren                            ' ########## '
'1 * Gruppieren:
'   GroupBy
Public Function GroupBy(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List

End Function

' ########## '                            Set                                ' ########## '
'4 * Set
'   Concat, Union, Intersect, Except
Public Function Concat(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

Public Function Union(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

Public Function Intersect(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

Public Function Except(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

' ########## '                       Konvertieren(Imp)                          ' ########## '
'2 * Konvertieren(Import):
'   OfType, Cast
Public Function OfType(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function Cast(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

' ########## '                       Konvertieren(Exp)                          ' ########## '
'5(6) * Konvertieren(Export):
'   (ToArray), ToList, ToDictionary, ToLookup, AsEnumerable, AsQueryable
Public Function ToArrayy(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function ToList(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function ToDictionary(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function ToLookup(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function AsEnumerable(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function
Public Function AsQueryable(lambObj As Object, keySelectorFunc As String, Optional comparerFunc As String) As List
    '
End Function

' ########## '                          Element                             ' ########## '
'9 * Element:
'   First, FirstOrDefault, Last, LastOrDefault, Single, SingleOrDefault, ElementAt, ElementAtOrDefault, DefaultIfEmpty
Public Function First()
    Dim i As Long: i = 0
    Select Case m_DataType
    Case vbObject:  Set First = m_Arr(i)
    Case vbVariant: If VarType(m_Arr(i)) = vbObject Then Set First = m_Arr(i) Else First = m_Arr(i)
    Case Else:      First = m_Arr(i)
    End Select
End Function
Public Function FirstOrDefault()
    '
End Function
Public Function Last()
    Dim i As Long: i = m_Count - 1
    Select Case m_DataType
    Case vbObject:  Set Last = m_Arr(i)
    Case vbVariant: If VarType(m_Arr(i)) = vbObject Then Set Last = m_Arr(i) Else Last = m_Arr(i)
    Case Else:      Last = m_Arr(i)
    End Select
End Function
Public Function LastOrDefault()
    '
End Function
Public Function SSingle()
    '
End Function
Public Function SingleOrDefault()
    '
End Function
Public Function ElementAt(ByVal Index As Long)
    Dim i As Long: i = Index
    Select Case m_DataType
    Case vbObject:  Set ElementAt = m_Arr(i)
    Case vbVariant: If VarType(m_Arr(i)) = vbObject Then Set ElementAt = m_Arr(i) Else ElementAt = m_Arr(i)
    Case Else:      ElementAt = m_Arr(i)
    End Select
End Function
Public Function ElementAtOrDefault()
    '
End Function
Public Function DefaultIfEmpty()
    '
End Function

' ########## '                          Aggregation                             ' ########## '
'6(7) * Aggregation:
'   Aggregate, Average, (Count), LongCount, Sum, Max, Min
Public Function Aggregate()
    '
End Function
Public Function Average()
'<Erweiterung> Enumerable.Average(Of TResult)(selector As Func(Of T, TResult)) As TResult
' Berechnet den Durchschnitt einer Sequenz von TResult-Werte, die durch den Aufruf einer Transformationsfunktion für jedes Element der Eingabesequenz ermittelt werden.
' * selector: Eine Transformationsfunktion, die auf jedes Element angewendet werden soll.
    Dim i As Long, Sum: Sum = CDec(0)
    Select Case m_DataType
    Case vbObject '???
    Case vbVariant
                    If VarType(m_Arr(i)) = vbObject Then
                        '
                    Else
                        For i = 0 To m_Count - 1
                            Sum = Sum + CDec(m_Arr(i))
                        Next
                        Average = Sum / m_Count
                    End If
    Case Else:      'entweder vor dem Summieren teilen oder nach dem Summieren teilen
                    For i = 0 To m_Count - 1
                        Sum = Sum + m_Arr(i)
                    Next
                    Average = Sum / m_Count
    End Select
End Function
Public Function CCount()
    'ruft die Anzahl der Elemente ab, die in List(Of T) enthalten sind.
    CCount = m_Count
End Function
Public Function LongCount()
    '1 v 2:
    'gibt ein Long, die die Gesamtanzahl von Elementen in einer Sequenz darstellt.
    
    '2 v 2:
    'Gibt ein Long das darstell, wieviele Elemente in einer Sequenz eine Bedingung erfüllen.
    ' * predicate: Eine Funktion, mit der jedes Element auf eine Bedingung überprüft wird.
    LongCount = m_Count
End Function
Public Function Sum(Of_TResult As EDataType, lambObj As Object, selectorFunc As String)
'<Erweiterung> Enumerable.Sum(Of TResult)(selector As Func(Of T, TResult)) As TResult
' Berechnet die Summe einer Sequenz von TResult-Werte, die durch den Aufruf einer Transformationsfunktion für jedes Element der Eingabesequenz ermittelt werden.
' * selector: Eine Transformationsfunktion, die auf jedes Element angewendet werden soll.

End Function
Public Function Max()
    '
End Function
Public Function Min()
    '1 v 12:
    '<Erweiterung> Enumerable.Min(Of TResult)(selector As Func(Of T, TResult)) As TResult
    '<Erweiterung> Enumerable.Min(selector As Func(Of T, TResult)) As TResult
    
    'Gibt den Mindestwert in einer generischen Sequenz zurück.
    
    '2 v 12 . . . 12 v 12:
    'Ruft für jedes Element einer Sequenz eine Transformationsfunktion auf und gibt den kleinesten T Wert.
    ' * selector: Eine Transformationsfunktion, die auf jedes Element angewendet werden soll.
    
End Function

' ########## '                          Quantifikatoren                             ' ########## '
'3(4) * Quantifikatoren:
'   All, Any, (Contains), SequenceEqual
Public Function All()
    '
End Function
Public Function AAny()
    '
End Function
Public Function CContains()
    '
End Function
Public Function SequenceEqual()
    '
End Function

' ########## '                          Generieren                             ' ########## '
'3 * Generieren:
'   Empty, Range, Repeat
Public Function EEmpty()
    EEmpty = m_Count = 0
End Function
Public Function Range()
    '
End Function
Public Function Repeat()
    '
End Function



'Funktionen die keine Erweiterung sind:
'======================================
'01:      Sub List(Of T).Add(item As T)                                                                                      ' 01: Fügt am Ende der List ein Objekt hinzu
'02:      Sub List(Of T).AddRange(Collection As IEnumerable(Of Book))                                                        ' 02: Fügt die Elemente der angegebenen Auflistung der List hinzu
'03: Function List(Of T).AsReadOnly() As ObjectModel.ReadOnlyCollection(Of T)                                                ' 03: Gibt einen schreibgeschützten ObjectModel.ReadOnlyCollection(Of T)-Wrapper für die aktuelle Auflisting zurück.
'04: Function List(Of T).BinarySearch(index As Integer, count as Integer, item As T, comparer As IComparer(Of T)) As Integer ' 11: Durchsucht mithilfe des angegebenen Vergleichs einen Bereich von Elementen in der sortierten List(Of T) nach einem Element und gibt den nullbasierten Index des Elements zurück.
'05: Property List(Of T).Capacity                                                                                            ' 12: Ruft die Gesamtzahl der Elemente ab, die die interne Datenstruktur ohne Änderung der Größe aufnehmen kann, oder legt diese Anzahl fest.
'06:      Sub List(Of T).Clear()                                                                                             ' 14: Entfernt alle Elemente aus der List(Of T)
'07: Function List(Of T).Contains                                                                                            ' 16: Bestimmt, ob sich ein Element in List(Of T) befindet.
'08: Function List(Of T).ConvertAll(Of Double)()                                                                             ' 17: Konvertiert die Elemente in der aktuellen List(Of T) in einen anderen Typ und gibt eine Liste der konvertierten Elemente zurück
'09:      Sub List(Of T).CopyTo()                                                                                            ' 18: Kopiert die gesamte List(Of T) in ein kompatibles eindimensionales Array, wobei am Anfang des Zielarrays begonnen wird.
'10: Property List(Of T).Count                                                                                               ' 19: Ruft die Anzahl der Elemente ab, die in List(Of T) enthalten sind.
'11: Function List(Of T).Enumerator                                                                                          ' 24: Zählt die Elemente einer List(Of T) auf.
'12: Function List(Of T).Equals()                                                                                            ' 25: bestimmt ob das angegebene Objekt mit dem aktuellen Objekt identisch ist.
'13: Function List(Of T).Exists()                                                                                            ' 27: Bestimmt ob die List(Of T) Elemente enthält, die mit den vom angegebenen Prädikat definierten Bedingungen übereinstimmen.
'14: Function List(Of T).Find()                                                                                              ' 28: Sucht nach einem Element, das die durch das angegebene Prädikat definierten Bedingungen erfüllt, und gibt das erste Vorkommen im gesamten List(Of T) zurück.
'15: Function List(Of T).FindAll()                                                                                           ' 29: Ruft alle Elemente ab, die die vom angegebenen Prädikat definierten Bedigungen erfüllen.
'16: Function List(Of T).FindIndex()                                                                                         ' 30: Sucht nach einem Element, das die durch das angegebene Prädikat definierten Bedingungen erfüllt, und gibt den nullbasierten Index des ersten Vorkommens im gesamten List(Of T) zurück.
'17: Function List(Of T).FindLast()                                                                                          ' 31: Sucht nach einem Element, das die durch das angegebene Prädikat definierten Bedingungen erfüllt, und gibt den nullbasierten Index des letzten Vorkommens im gesamten List(Of T) zurück.
'18: Function List(Of T).FindLastIndex(match As Predicate(of T) As Integer                                                   ' 32: Sucht nach einem Element, das die durch das angegebene Prädikat definierten Bedingungen erfüllt, und gibt den nullbasierten Index des letzten Vorkommens im gesamten List(Of T) zurück.
'19:      Sub List(Of T).ForEach(action As Action(Of T))                                                                     ' 35: Führt die angegebene Aktion für jedes Element der List(Of T) aus.
'21: Function List(Of T).GetEnumerator() As List(Of T).Enumerator                                                            ' 36: Gibt einen Enumerator zurück, der die List(Of T) durchläuft.
'22: Function List(Of T).GetHashCode()                                                                                       ' 37: Fungiert als die Standardhashfunktion.
'23: Function List(Of T).GetRange(index As Integer, count as Integer) As List(Of T)                                          ' 38: Erstellt eine flache Kopie eines Bereichs von Elementen in der Quell-List(Of T)
'24: Function Object.GetType() As Type                                                                                       ' 39: Ruft den Type der aktuellen Instanz ab.
'25: Function List(Of T).IndexOf(item As T) As Integer                                                                       ' 42: Sucht nach dem angegebenen Objekt und gibt den nullbasierten Index des ersten Vorkommens innerhalb der gesamten List(Of T) zurück.
'26: Function List(Of T).Insert()                                                                                            ' 43:
'27: Function List(Of T).InsertRange()         ' 44:
'28: Function List(Of T).Item                  ' 46:
'29: Function List(Of T).LastIndexOf()         ' 49:
'30: Function List(Of T).Remove()              ' 59:
'31: Function List(Of T).RemoveAll()           ' 60:
'32: Function List(Of T).RemoveAt(Index As Integer) ' 61:
'33: Function List(Of T).RemoveRange()         ' 62:
'34: Function List(Of T).Reverse()             ' 63:
'35: Function List(Of T).Sort()                ' 71:
'36: Function List(Of T).ToArray()             ' 75:
'37: Function List(Of T).ToString()            ' 79:
'38: Function List(Of T).TrimExcess()          ' 80:

'Funktionen die Erweiterung sind:
'================================
'01: Function List(Of T).Aggregate             ' 03: Wendet eine Akkumulatorfunktion auf eine Sequenz (List) an.
'02: Function List(Of T).Aggregate(seed)       ' 04: Wendet eine Akkumulatorfunktion auf eine Sequenz (List) an. Der angegebene Startwert wird als erster Akkumulatorwert verwendet.
'03: Function List(Of T).All                   ' 05: Bestimmt ob alle Elemente einer Sequenz eine Bedingung erfüllen.
'04: Function List(Of T).Any                   ' 06: Bestimmt ob eine Sequenz Elemente enthält.
'05: Function List(Of T).AsEnumerable          ' 07: Gibt die Eingabe als IEnumerable(Of Out T) zurück
'06: Function List(Of T).AsParallel            ' 08: Ermöglicht die Parallelisierung einer Abfrage
'07: Function List(Of T).AsQueryable           ' 09: Konvertiert ein generisches IEnumerable(Of Out T) auf einen generischen IQueryable(Of Out T)
'08: Function List(Of T).Average(selector As Func(Of String, Integer)) As Double '(+ 9 Überladungen) ' 10: Berechnet den Durchschnitt einer Sequenz von Integer Werte, die durch den Aufruf einer Transformationsfunktion für jedes Element der Eingabesequenz ermittelt werden.
'09: Function List(Of T).Cast(Of Byte)         ' 13: Wandelt die Elemente einer IEnumerable in den angegebenen Typ.
'10: Function List(Of T).Concat                ' 15: Verkettet zwei Sequenzen.
'11: Function List(Of T).DefaultIfEmpty        ' 20: Gibt die Elemente der angegebenen Sequenz oder den Standardwert des Typparameters in einer Singeltonauflistung zurück, wenn die Sequenz leer ist.
'12: Function List(Of T).Distinct              ' 21: Gibt mithilfe des Standardgleichheitsvergleichs zum Vergleichen von Werten unterschiedliche Elemente aus einer Sequenz zurück.
'13: Function List(Of T).ElementAt             ' 22: Gibt das Element am angegebenen Index in einer Sequenz zurück.
'14: Function List(Of T).ElementAtOrDefault    ' 23: Gibt das Element an einem angegebenen Index in einer Sequenz oder einen Standardwert zurück, wenn der Index außerhalb des gültigen Bereichs liegt.
'15: Function List(Of T).Except                ' 26: Erzeugt die Differenzelemente von zwei Sequenzen mithilfe des Standardgleichheitsvergleichs zum Vergleichen von Werten.
'16: Function List(Of T).First                 ' 33: Gibt das erste Element einer Sequenz zurück.
'17: Function List(Of T).FirstOrDefault        ' 34: Gibt das erste Element einer Sequenz zurück oder einen Standardwert falls die Sequenz keine Elemente enthält
'18: Function List(Of T).GroupBy(Of Integer)   ' 40:
'19: Function List(Of T).GroupJoin(Of Integer) ' 41:
'20: Function List(Of T).Intersect             ' 45:
'21: Function List(Of T).Join(Of Integer)      ' 47:
'22: Function List(Of T).Last                  ' 48:
'23: Function List(Of T).LastOrDefault         ' 50:
'24: Function List(Of T).LongCount             ' 51:
'25: Function List(Of T).Max                   ' 52:
'26: Function List(Of T).Max(Of Integer)       ' 53:
'27: Function List(Of T).Min()                 ' 54:
'28: Function List(Of T).Min(Of Integer)       ' 55:
'29: Function List(Of T).OfType(Of Integer)    ' 56:
'30: Function List(Of T).OrderBy(Of Integer)   ' 57:
'31: Function List(Of T).OrderByDescending(Of Integer) ' 58:
'32: Function List(Of T).Select(Of String)     ' 64:  Projiziert jedes Element einer Sequenz in ein neues Format.
'33: Function List(Of T).SelectMany(Of String) ' 65:
'34: Function List(Of T).SequenceEqual         ' 66:
'35: Function List(Of T).Single()              ' 67:
'36: Function List(Of T).SingleOrDefault       ' 68:
'37: Function List(Of T).Skip                  ' 69:
'38: Function List(Of T).SkipWhile             ' 70:
'39: Function List(Of T).Sum                   ' 72:
'40: Function List(Of T).Take                  ' 73:
'41: Function List(Of T).TakeWhile             ' 74:
'42: Function List(Of T).ToDictionary(Of String) ' 76:
'43: Function List(Of T).ToList                ' 77:
'44: Function List(Of T).ToLookup(Of String)   ' 78:
'45: Function List(Of T).TrueForAll()          ' 81:
'46: Function List(Of T).Union                 ' 82:
'47: Function List(Of T).Where                 ' 83:
'48: Function List(Of T).Zip(Of String)        ' 84:

'6 * Filtern:
'   Where, Distinct, Take, TakeWhile, Skip, SkipWhile
'
'2 * Projizieren:
'   Select, SelectMany
'
'2 * Verknüpfen:
'   Join, GroupJoin
'
'5 * Sortieren:
'   OrderBy, OrderByDescending, ThenBy, ThenByDescending, Reverse
'
'1 * Gruppieren:
'   GroupBy
'
'4 * Set
'   Concat, Union, Intersect, Except
'
'2 * Konvertieren(Import):
'   OfType, Cast
'
'5(6) * Konvertieren(Export):
'   (ToArray), ToList, ToDictionary, ToLookup, AsEnumerable, AsQueryable
'
'9 * Element:
'   First, FirstOrDefault, Last, LastOrDefault, Single, SingleOrDefault, ElementAt, ElementAtOrDefault, DefaultIfEmpty
'
'6(7) * Aggregation:
'   Aggregate, Average, (Count), LongCount, Sum, Max, Min
'
'4 * Quantifikatoren:
'   All, Any, Contains, SequenceEqual
'
'3 * Generieren:
'   Empty, Range, Repeat
'
'Sum: 51



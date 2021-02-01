VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ' Zeilen: 636, last modified 17.11.2008
'Public Class PathFileName
'Dies ist keine Klasse im .NET-Framework
'Diese Klasse soll eigentlich nur alle Dateioperationen abdecken
'
'man kann damit beliebig den Pfad ändern, das Laufwerk herauslesen,
'die Extension herauslesen oder verändern, und den gesamten Pfad oder
'einen gekürzten Pfad bekommen
'Man kann hier auch die Datei Kopieren, Löschen und überprüfen ob die
'Datei oder der Pfad überhaupt existieren.
'Diese Klasse ist nicht dazu da eine Datei zu öffnen zu lesen oder zu
'schreiben, da die Art der Datei nicht bekannt ist.
Private mPFN As String
'Private mDirSepChar As Long ' As Long = 92 ' String = "\" ' = Chr(92)
'Private mExtSepChar As Long ' As Long = 46 ' String = "." ' = Chr(46)

'sind jetzt Konstanten dann darf Class_Initialize entfallen
Private Const C_DirSepChar As Long = 92 ' String = "\" ' = Chr(92)
Private Const C_ExtSepChar As Long = 46 ' String = "." ' = Chr(46)
Private Const C_InvalidChars As String = "*?""<>|" 'bei Dateiname noch "/\:"
Private Const C_3Pt As String = "..."
'Private Declare Function CopyFileA Lib "kernel32.dll" ( _
'    ByVal lpExistingFileName As String, _
'    ByVal lpNewFileName As String, _
'    ByVal bFailIfExists As Long) As Long
'Private Declare Function MoveFileA Lib "kernel32.dll" ( _
'    ByVal lpExistingFileName As String, _
'    ByVal lpNewFileName As String) As Long
'Private Declare Function DeleteFileA Lib "kernel32.dll" ( _
'    ByVal lpFileName As String) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" ( _
    ByVal lpPath As String) As Long
'Private Declare Function RemoveDirectoryA Lib "kernel32" ( _
'    ByVal lpPathName As String) As Long
Public IsValid As Boolean

'####################'      Constructors      '####################'
'Private Sub Class_Initialize()
'    mDirSepChar = 92 '"\"
'    mExtSepChar = 46  '"."
'End Sub

Friend Sub New_(ByVal aPath As String, _
                Optional ByVal aFileName As String, _
                Optional ByVal aExt As String)
    If Len(aFileName) = 0 And Len(aExt) = 0 Then
        mPFN = aPath 'RemoveQuotes(aPath)
    Else
        If Len(aExt) = 0 Then
            If Len(aPath) = 0 Then
                mPFN = RemoveQuotes(aFileName)
            Else
                mPFN = GetValidPath(aPath) & _
                       GetFileName(aFileName)
            End If
        Else
            If Len(aFileName) Then
                mPFN = GetValidPath(aPath) & _
                       GetFileName(aFileName) & _
                       GetValidExt(aExt)
            Else
                'nur aPath und ext
                'aPath beinhaltet auch den Dateinamen
                'blöd und was wenn nich?
                'dann müßte man noch überprüfen ob der Pfad so überhaupt existiert
                'Naja, da gibts halt einfach nix anderes, wenn der Pfad falsch ist,
                mPFN = aPath
                Ext = GetValidExt(aExt)
            End If
        End If
    End If
    IsValid = MakeValid(mPFN)
End Sub

Private Sub Class_Terminate()
    mPFN = vbNullString
End Sub

'####################'    Public Properties    '####################'
Public Property Get Name() As String
'Attribute Name.VB_UserMemId = 0
    'gibt den vollständigen Pfaddateinamen zurück
    'Laufwerk + Pfad + Dateiname + Extension
    If Not IsValid Then IsValid = MakeValid(mPFN)
    Name = mPFN
End Property
Public Property Let Name(ByVal Value As String)
    mPFN = Validate(Value)
End Property
Public Property Let ValidName(Value As String)
    Me.IsValid = True
    mPFN = Value
End Property
Public Property Get Drive() As String
    Drive = GetDrive(mPFN)
End Property
Public Property Let Drive(RHS As String)
    Dim dr As String: dr = GetDrive(mPFN)
    If Len(RHS) = Len(dr) Then Mid$(mPFN, 1, Len(RHS)) = RHS _
    Else mPFN = GetValidDrive(RHS) & Mid$(mPFN, Len(dr) + 1)
End Property

Public Property Get Path() As String
    'gibt den Pfad inkl, Laufwerk zurück
    Path = GetPath(mPFN)
End Property
Public Property Let Path(RHS As String)
    If LenB(RHS) = LenB(GetPath(mPFN)) Then
        Mid$(mPFN, 1, Len(RHS)) = GetValidPath(RHS)
    Else
        mPFN = GetValidPath(RHS) & GetFileName(mPFN)
    End If
End Property

Public Property Get OnlyPath() As String
    'gibt den Pfad ohne Laufwerk und ohne Dateinamen zurück
    OnlyPath = GetOnlyPath(mPFN)
End Property
Public Property Let OnlyPath(RHS As String)
    Dim dr As String: dr = GetDrive(mPFN)
    Dim op As String: op = GetOnlyPath(mPFN)
    If Len(RHS) = Len(op) Then
        Mid$(mPFN, Len(dr), Len(op)) = GetValidPath(RHS)
    Else
        mPFN = dr & GetValidPath(RHS) & GetFileName(mPFN)
    End If
End Property

Public Property Get OnlyFileName() As String
    'gibt den Dateinamen, ohne Pfad und ohne Extension zurück
    OnlyFileName = GetOnlyFileName(mPFN)
End Property
Public Property Let OnlyFileName(RHS As String)
    Dim pa As String: pa = GetPath(mPFN)
    Dim of As String: of = GetOnlyFileName(mPFN)
    If Len(RHS) = Len(of) Then
        Mid$(mPFN, Len(pa) + 1) = RHS
    Else
        mPFN = pa & RHS & GetValidExt(GetExt(mPFN))
    End If
End Property

Public Property Get FileName() As String
    FileName = MidB$(mPFN, LenB(GetPath(mPFN)) + 1)
End Property
Public Property Let FileName(RHS As String)
    If LenB(RHS) = LenB(GetFileName(mPFN)) Then
        Mid$(mPFN, 1, Len(RHS)) = RHS
    Else
        mPFN = GetPath(mPFN) & RHS
    End If
End Property

Public Function Equals(other As PathFileName) As Boolean
    Equals = Me.CompareTo(other.Name) = 0
End Function
Public Property Get Ext() As String
    Ext = GetExt(mPFN)
End Property
Public Property Let Ext(RHS As String)
    'If Len(RHS) = Len(GetExt(mPFN)) Then Mid$(mPFN, Len(mPFN) - Len(RHS) + 1 _
    ') = GetValidExt(RHS) Else mPFN = GetPath(mPFN) & GetValidExt(RHS)
    Dim Ex As String: Ex = GetExt(mPFN)
    If Len(RHS) = Len(Ex) Then
        Mid$(mPFN, Len(mPFN) - Len(Ex)) = GetValidExt(RHS)
    Else
        'mPFN = GetPath(mPFN) & GetOnlyFileName(mPFN) & GetValidExt(RHS)
        'oder so:
        If Len(Ex) > 0 Then
            mPFN = Mid$(mPFN, 1, Len(mPFN) - Len(Ex) - 1) & GetValidExt(RHS)
        Else
            mPFN = mPFN & GetValidExt(RHS)
        End If
    End If
End Property

Public Property Get CountDirDepth() As Long
    'zählt die Verzeichnistiefe des Pfades
    Dim op As String: op = GetValidPath(OnlyPath)
    Dim p As Long:     p = 1
    If Len(op) > 0 Then
        Dim c As String: c = Chr$(C_DirSepChar)
        Do While p > 0
            p = InStr(p + 1, op, c)
            If p Then CountDirDepth = CountDirDepth + 1
            If p = Len(op) Then Exit Do
        Loop
    End If
End Property
Public Property Get Directory(ByVal i As Long) As String
    'liefert den Namen eines Verzeichnisknotens i im Pfad
    Dim P1 As Long, P2 As Long
    If Not GetDirPart(i, P1, P2, Directory) Then
        MsgBox "i: " & CStr(i) & " ausserhalb des gültigen Bereichs"
    End If
End Property
Public Property Let Directory(ByVal i As Long, ByVal StrVal As String)
    'setzt den Namen des Verzeichnisknotens i im Pfad
    'dadurch kann schnell zwiscen verschiedenen Verzeichnisästen
    'gewechselt werden. z.B:
    'C:\DokUEinst\Oliver Meyer\Eigene Dateien\
    'C:\DokUEinst\Administrator\Eigene Dateien\
    Dim n As Long: n = CountDirDepth
    If 0 < i And i <= n Then
        Dim P1 As Long, P2 As Long
        Dim si As String
        Call GetDirPart(i, P1, P2, si)
        If Len(si) = Len(StrVal) Then
            Mid$(mPFN, P1, P2 - P1) = StrVal
        Else
            Dim ld As Long: ld = Len(Drive)
            Dim L As String, r As String
            L = Left$(mPFN, ld + P1 - 1)
            r = Mid$(mPFN, ld + P2)
            mPFN = L & StrVal & r
        End If
    Else
        MsgBox "i: " & CStr(i) & " ausserhalb des gültigen Bereichs"
    End If
End Property
Public Property Get Quoted() As String
    'Gibt den Pfaddateiname quoted zurück (mit Anführungszeichen).
    'pfnQ = Chr$(34) & mPFN & Chr$(34)
    'oder so is hübscher:
    Quoted = """" & mPFN & """"
End Property
'Public Property Get VersionInfo() As FileVersionInfo
'    'einfach auskommentieren wenn Version nicht gebraucht wird (in kleineren Projekten)
'    Set VersionInfo = New FileVersionInfo
'    Set VersionInfo = VersionInfo.GetVersionInfo(Me.Name)
'End Property

Public Property Get length() As Long
    length = Len(mPFN)
End Property
Public Property Let length(RHS As Long)
    Dim L As Long: L = length
    If L <> RHS Then
      If L > RHS Then
        'verkleinern
        mPFN = Left$(mPFN, RHS)
      Else 'If l > RHS Then
        'mit Nullzeichen rechts vergrößern
        mPFN = mPFN & String$(RHS - L, vbNullChar)
      End If
    End If
End Property

'####################'    Public Procedures    '####################'
Public Sub AddDirectory(ByVal aDir As String)
    aDir = GetValidPath(aDir)
    mPFN = Path & aDir & Chr$(C_DirSepChar) & FileName
End Sub
Public Function RemoveQuotes(Optional ByVal aStrPFN As String) As String
    Dim q1 As Long: q1 = IIf(Asc(Left$(aStrPFN, 1)) = 34, 2, 1)
    Dim q2 As Long: q2 = IIf(Asc(Right$(aStrPFN, 1)) = 34, Len(aStrPFN) - 1, Len(aStrPFN))
    RemoveQuotes = Mid$(aStrPFN, q1, q2 - q1 + 1)
End Function

Public Function pStr(Optional ByVal BufLen As Long = 1024) As Long
'übergibt den Zeiger auf den String falls noch kein Wert vorhanden, wird ein
'Puffer angelegt, und der Zeiger darauf übergeben.
    If Len(mPFN) = 0 Then
        mPFN = String$(BufLen, vbNullChar)
    End If
    pStr = StrPtr(mPFN)
End Function
Public Function Convert(ByVal convto As VbStrConv, Optional ByVal LocaleID As Long) As String
    If LocaleID Then
        mPFN = StrConv(mPFN, convto, LocaleID)
    Else
        mPFN = StrConv(mPFN, convto)
    End If
    Call MakeValid(mPFN)
    Convert = mPFN
End Function
Public Function Exists(Optional ByVal StrPFN As String) As Boolean
TryE: On Error GoTo CatchE
    If Len(StrPFN) = 0 Then StrPFN = mPFN
    Exists = ((GetAttr(StrPFN) And (vbDirectory Or vbVolume)) = 0)
CatchE:
    'hier keine Fehlerbehandlung sondern einfach raus
End Function
Public Function DirExists(Optional ByVal strPath As String) As Boolean
TryE: On Error GoTo CatchE
    ' Wenn ein Fehler aufgetreten ist, dann ist DirExists auf jeden Fall
    ' False, sonst ist es nur True, wenn auch das Directory-Attribut beim
    ' existierenden Objekt vorhanden ist.
    If Len(strPath) = 0 Then strPath = GetPath(mPFN)
    DirExists = ((GetAttr(strPath) And vbDirectory))
CatchE:
    'hier keine Fehlerbehandlung sondern einfach raus
End Function

Public Function Shorted(Optional ByVal MaxLen As Long) As String
    'gibt einen gekürzten String zurück
    If MaxLen = 0 Then MaxLen = 100
    'wenn der Dateiname länger als 100 ist, dann nur den Dateiname anzeigen
    Shorted = OnlyFileName
    If Len(Shorted) < MaxLen Then
        'Wenn kleiner dann noch das Laufwerk dazunehmen:
        Shorted = Drive & C_3Pt & OnlyFileName
        If Len(Shorted) < MaxLen Then
            'wenn kleiner dann noch die Dateierweiterung dazunehmen
            Shorted = Drive & C_3Pt & OnlyFileName & "." & Ext
            If Len(Shorted) < MaxLen Then
                'wenn kleiner dann den Pfad dazunehmen
                'für den verbleibenden Rest den Pfad in zwei Teile teilen
                Dim d As Long
                d = MaxLen - Len(Shorted)
                If d > Len(OnlyPath) Then
                'wenn Pfad kleiner als der Rest dann zeige ganzen Dateinamen
                    Shorted = Me.Name
                Else
                    Shorted = Drive & GetDividedPath(d) & _
                              OnlyFileName & "." & Ext
                End If
            End If
        End If
    End If
End Function
Public Function AssignCopy() As PathFileName
    'um ständig wiederkehrendes Validieren zu umgehen nimmt man diese Funktion
    'die einfach nur den Inhalt in ein neues Objekt kopiert.
    'Das Kopie-Object ist dann genauso valid wie das Ursprungsobjekt.
    Set AssignCopy = New PathFileName
    AssignCopy.ValidName = mPFN 'Me.Name
    AssignCopy.IsValid = Me.IsValid
End Function
Public Function CopyTo(ByVal DestinationPFN As String) As Boolean
TryE: On Error GoTo CatchE
    If Exists Then
        Debug.Print "copy to: " & DestinationPFN
        Call FileCopy(Me.Name, DestinationPFN)
        CopyTo = True
    End If
    Exit Function
CatchE:
    Call ErrHandler("CopyTo")
End Function

Public Function Delete(Optional StrPFN As String = vbNullString) As Boolean
'löscht nur was es auch gibt
TryE: On Error GoTo CatchE
    If StrPFN = vbNullString Then StrPFN = Me.Name
    If Exists(StrPFN) Then
        'was soll man verwenden? Kill oder DeleteFile
        Call Kill(StrPFN)
        Delete = True
    End If
    Exit Function
CatchE:
    '55: Datei bereits geöffnet
    Call ErrHandler("Delete")
End Function

Public Function CompareTo(ByVal StrB As String) As Long
'vergleicht diese Instanz mit der angegebenen Zeichenfolge
'Eine 32-Bit-Ganzzahl mit Vorzeichen, die die lexikalische
'Beziehung der beiden verglichenen Elemente angibt.
'Wert            Bedingung
'Kleiner als 0   Diese Instanz ist kleiner als strB.
'0               Diese Instanz ist gleich strB.
'Größer als 0    Diese Instanz ist größer als strB.
'                – oder –   strB ist null.
    Dim aStrA As String
    Dim bStrB As String
    aStrA = mPFN 'Me.Name
    bStrB = StrB
    Call MakeValid(aStrA)
    Call MakeValid(bStrB)
    CompareTo = StrComp(aStrA, bStrB, vbTextCompare)
End Function

'ist BolRenameMe = False, dann wird dieses Objekt nicht verändert,
'sondern nur die Datei auf der Festplatte umbenannt
'dieses Objekt existiert dann nicht mehr als Datei auf der Festplatte
'(FileExists liefert dann false)
Public Function Rename(ByVal ToPFN As String, _
                       Optional ByVal BolRenameMe As Boolean) As Boolean
TryE: On Error GoTo CatchE
    Dim RenFile As PathFileName: Set RenFile = MNew.PathFileName(ToPFN)
    Dim Cmd     As String
    If Not RenFile.Exists Then
        'Wird nicht renamed wenn es die Datei schon gibt
        Cmd = Chr(34) & Me.Name & Chr(34) & " " & Chr(34) & RenFile.Name & Chr(34)
        Call Shell("cmd /c ren " & Cmd, 0)
        Rename = True
        If BolRenameMe Then
            'soll nur die Datei auf der Festplatte umbenannt werden
            'oder auch dieser Dateiname selber verändert werden
            Me.Name = ToPFN
        End If
    End If
    Exit Function
CatchE:
    Call ErrHandler("Rename")
End Function
Public Function MakeDirectory(Optional ByVal sNewDir As String) As Boolean
TryE: On Error GoTo CatchE
    'ein neues Verzeichnis mit Unterverzeichnissen anlegen
    If Len(sNewDir) = 0 Then sNewDir = Me.Path
    Call MakeValid(sNewDir)
    Dim NewDir As PathFileName: Set NewDir = MNew.PathFileName(sNewDir)
    'zuerst in das jeweilige Laufwerk wechseln
    Dim sd As String: sd = NewDir.Drive()
    Call ChDir(sd)
    Dim i As Long
    For i = 1 To NewDir.CountDirDepth
        sd = NewDir.Directory(i)
        'das Verzeichnis anlegen
        Call MkDir(sd)
        'gleich ins verzeichnis wechseln, damit das nächste
        'Unterverzeichnis angelegt werden kann
        Call ChDir(sd)
    Next
    MakeDirectory = True
    Exit Function
CatchE:
    Call ErrHandler("MakeDirectory", "Der Pfad konnte nicht angelegt werden: " & vbCrLf & sNewDir)
End Function
Public Function DeleteDir(Optional ByVal sDelDir As String) As Boolean
TryE: On Error GoTo CatchE
    If Len(sDelDir) <> 0 Then
        sDelDir = Me.Path
        If Me.Exists() Then Exit Function
    End If
    'Call RemoveDirectoryA(deldir)
    'nur ein komplett leeres Directory
    Call RmDir(sDelDir)
    DeleteDir = True
    Exit Function
CatchE:
    Call ErrHandler("DeleteDirectory", sDelDir)
End Function
Public Function Validate(Optional sPath As String) As String
    If Len(sPath) = 0 Then
        Validate = mPFN
    Else
        Validate = sPath
    End If
    Call MakeValid(Validate)
    'IsValid = True
End Function

Public Function PFNwithChangedExt(ByVal newExt As String) As PathFileName
    Set PFNwithChangedExt = MNew.PathFileName(mPFN, , newExt)
    If Len(newExt) = 0 Then PFNwithChangedExt.Ext = vbNullString
End Function
'####################'   Private Procedures   '####################'
'Ein Pfaddateiname besteht aus:
'Laufwerk  |      Nur Pfad       |   Nur Dateiname  |  Extension
'  C:\     |  Windows\System32\  |      user32      |    .dll
'
'bzw:
'          Pfad          |  Dateiname
'  C:\Windows\System32\  |  user32.dll
'
'####################'        Parsers         '####################'
Private Function GetDrive(aPFN As String) As String
    If Len(aPFN) Then
        Dim p As Long: p = InStr(1, aPFN, Chr$(C_DirSepChar))
        If p Then GetDrive = Left$(aPFN, p)
    End If
End Function
Private Function GetPath(aPFN As String) As String
    If Len(aPFN) Then
        Dim p As Long: p = InStrRev(aPFN, Chr$(C_DirSepChar))
        If p > 0 Then GetPath = Left$(aPFN, p)
    End If
End Function
Private Function GetOnlyPath(aPFN As String) As String
    If Len(aPFN) Then
        Dim sp As String: sp = GetPath(aPFN)
        GetOnlyPath = Mid$(sp, Len(GetDrive(sp)) + 1)
    End If
End Function
Private Function GetOnlyFileName(aPFN As String) As String
    If Len(aPFN) Then
        Dim FN As String: FN = GetFileName(aPFN)
        If Len(FN) > 0 Then GetOnlyFileName = Mid$(FN, 1, Len(FN) - Len(GetExt(FN)) - 1)
    End If
End Function
Private Function GetFileName(aPFN As String) As String
    If Len(aPFN) Then
        GetFileName = Mid$(aPFN, Len(GetPath(aPFN)) + 1)
    End If
End Function
Private Function GetExt(aPFN As String) As String
    If Len(aPFN) Then
        Dim E As Long: E = InStrRev(aPFN, Chr$(C_ExtSepChar))
        If E > 0 Then GetExt = Mid$(aPFN, E + 1)
    End If
End Function
Private Function GetDirPart(ByVal n As Long, _
                            ByRef pos1 As Long, _
                            ByRef pos2 As Long, _
                            ByRef sDir As String) As Boolean
'gibt einen Teil des Directories  zurück
'OK man könnte das Ganze auch mit einem Array machen.
'Aber da man normalerweise nur selten einzelne Teile
'des Pfades ändert, lohnt sich der Verwaltungsaufwand
'nicht.
    Dim op As String: op = GetValidPath(OnlyPath)
    Dim i  As Long
    Dim P1 As Long, P2 As Long
    For i = 0 To n - 1
        P1 = P2 + 1
        P2 = InStr(P1, op, Chr$(C_DirSepChar))
    Next
    If P2 > P1 Then
        sDir = Mid$(op, P1, P2 - P1)
        pos1 = P1
        pos2 = P2
        GetDirPart = True
    End If
End Function

'####################'       Validierer       '####################'
Private Function GetValidDrive(aDrive As String) As String
'falls nicht ganz rechts vorhanden,
'stattet eine Laufwerksangabe mit einem Pfad-Trennzeichen ("\") aus.
    If LenB(aDrive) > 0 Then
        If Asc(Mid$(aDrive, Len(aDrive) - 1)) <> C_DirSepChar Then
            GetValidDrive = aDrive & Chr$(C_DirSepChar)
        Else
            GetValidDrive = aDrive
        End If
    End If
End Function
Private Function GetValidPath(aPath As String) As String
'falls nicht ganz rechts vorhanden,
'stattet einen Pfad mit einem Pfad-Trennzeichen ("\") aus.
    If LenB(aPath) > 0 Then
        If Asc(Mid$(aPath, Len(aPath))) <> C_DirSepChar Then
            GetValidPath = aPath & Chr$(C_DirSepChar)
        Else
            GetValidPath = aPath
        End If
    End If
End Function
Private Function IsValidFileName(aFileName As String) As Boolean
    Dim i As Long
    Dim X As String: X = "\/:*?""<>|"
    If Len(aFileName) Then
        For i = 1 To Len(X)
            IsValidFileName = (InStr(1, aFileName, Mid$(X, i, 1)) <> 0)
            If Not IsValidFileName Then Exit Function
        Next
    End If
End Function
Private Function GetValidFileName(aFileName As String) As String
    Dim i As Long
    Dim c As String, X As String: X = "\/:" & C_InvalidChars '*?""<>|"
    If Len(aFileName) Then
        'GetValidFileName = aFileName
        GetValidFileName = RemoveChars(aFileName, X)
        'For i = 1 To Len(x)
        '    c = Mid$(x, i, 1)
        '    GetValidFileName = Replace$(GetValidFileName, c, vbNullString)
        'Next
    End If
End Function
Private Function GetValidExt(aExt As String) As String
'falls nicht ganz links vorhanden,
'stattet eine  Erweiterung mit einem Extensions-Trennzeichen (".") aus.
    If Len(aExt) > 0 Then
        Dim c As String: c = Chr$(C_ExtSepChar)
        If InStr(1, aExt, c) = 0 Then
            GetValidExt = c & aExt
        Else
            GetValidExt = aExt
        End If
    End If
End Function
Private Function GetDividedPath(MaxLen As Long) As String
    Dim half As Long: half = MaxLen \ 2 - 2 '-2 wegen den Pünktchen
    If half > 0 Then
        Dim o As String: o = OnlyPath
        GetDividedPath = Left$(o, half) & C_3Pt & Right$(o, half)
    Else 'dann nur die Pünktchen ausgeben
        GetDividedPath = C_3Pt
    End If
End Function
Private Function MakeValid(ByRef aStrPath As String) As Boolean
    'Nullen von hinten entfernen
TryE: On Error GoTo CatchE
    Dim p As Long: p = InStr(1, aStrPath, vbNullChar)
    If p <> 0 Then mPFN = Left$(aStrPath, p - 1)
    'und alle Zeichen entfernen die nicht hinein dürfen
    Dim i As Long
    Dim c As String, X As String: X = C_InvalidChars '"*?""<>|"
    Select Case C_DirSepChar
    Case 92: X = X & "/"
    Case 47: X = X & "\"
    End Select
    aStrPath = MStrings.RemoveChars(aStrPath, X)
    
    
'    'und alle doppelten Backslashes entfernen,
'    'die nicht am Anfang sind, denn ansonsten ist es ein Serverlaufwerk
    'braucht man das wirklich? nein vielleicht gibt es da Fehler irgendwo
'    c = Chr$(C_DirSepChar)
'    X = c & c
'    Dim posDblBackSlash As Long
'    posDblBackSlash = InStr(1, aStrPath, X)
'    If posDblBackSlash > 1 Then
'        aStrPath = Replace(aStrPath, X, c)
'    End If
    'und doppelpunkte ":" entfernen ausser in der Laufwerksangabe
    Dim sd As String: sd = GetDrive(aStrPath)
    Dim ld As Long:   ld = Len(sd)
    c = Chr$(58) '":"
    If InStr(ld + 1, aStrPath, c) Then
       'grummel grummel, replace gibt nur einen Teil zurück
        aStrPath = sd & Replace$(aStrPath, c, vbNullString, ld + 1)
    End If
    MakeValid = True
    Exit Function
CatchE:
    '
End Function

'String-Routinen
'jetzt in MStrings
'Private Function RemoveChars(ByVal this As String, CharsToRemove As String) As String
'    Dim c As String
'    Dim i As Long
'    RemoveChars = this
'    For i = 1 To Len(CharsToRemove)
'        c = Mid$(CharsToRemove, i, 1)
'        If InStr(1, this, c) Then
'            RemoveChars = Replace$(RemoveChars, c, vbNullString)
'        End If
'    Next
'End Function
'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal PrcName As String, _
                            Optional AddInfo As String, _
                            Optional BolLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    'ErrHandler = GlobalErrHandler(Me, PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
End Function
'Public Function GlobalErrHandler(ByVal ClsName As String, _
'                                 ByVal PrcName As String, _
'                                 Optional AddInfo As String, _
'                                 Optional BolLoud As Boolean = True, _
'                                 Optional bErrLog As Boolean = True, _
'                                 Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
'                                 ) As VbMsgBoxResult
'    Dim m As String
'    m = "In " & ClsName & "::" & PrcName & " ist ein Fehler aufgetreten." & vbCrLf & _
'        CStr(Err.Number) & ": " & Err.Description
'    If Len(AddInfo) Then
'        m = m & vbCrLf & "Info: " & AddInfo
'    End If
'    GlobalErrHandler = MsgBox(m, vbDecor)
'End Sub





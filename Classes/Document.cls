VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Persons    As List 'Of Person
Public Addresses  As List 'Of Address
Public TelefonNrs As List 'Of TelefonNr
Public Cities     As List 'Of City
Public Countries  As List 'Of Country
Public Tests      As List 'Of Test
Public AllObjects As List

Private m_pfn     As PathFilename

Private m_DocVersion   As String 'Datei-Dokumentversion
Private m_DataVersion  As Long   'Aktuelle Datenversion (wird immer um eins hochgezählt)
Private m_DataVerSaved As Long   'Datenversion des letzten gespeicherten Zustands

Private Sub Class_Initialize()
    'OK wir brauchen
    '* Datei-Dokumentversionsnummer
    '  -> ändert sich nur wenn die Datenstruktur sich irgendwannmal ändert
    '  -> Meldung ausgeben wenn Nummer unterschiedlich
    '     was wird gemacht, wird trotzdem eingelesen, werden Daten konvertiert etc.
    '* aktuelle Datenversionsnummer
    '  -> wird bei jeder Änderung um eins hochgezählt, weiß ob dirty oder clean
    '
    Set Persons = MNew.List(vbObject)    'Of Person
    Set Addresses = MNew.List(vbObject)    'Of Address
    Set TelefonNrs = MNew.List(vbObject)    'Of TelefonNr
    Set Cities = MNew.List(vbObject)    'Of City
    Set Countries = MNew.List(vbObject)    'Of Country
    Set Tests = MNew.List(vbObject)
    Set AllObjects = New List 'Of all above
    Set m_pfn = MNew.PathFilename(Application.DefaultFileName)
End Sub

Friend Sub New_(Optional ByVal aPFN As PathFilename = Nothing)
    Set m_pfn = aPFN
    If Not m_pfn Is Nothing Then
        If m_pfn.Exists Then LoadFile
    Else
        Set m_pfn = MNew.PathFilename(Application.DefaultFileName)
    End If
    'soll bei New gleich ein load erfolgen?
    '
    'OK wir brauchen einen default-dateiname -> wird in Application gespeichert
    'Programm startet, es wird ein Doc erzeugt mit dem default Dateinamen
    'es kann sein dass bereits eine datei mit dem Namen vorhanden ist, weil früher schonmal mit dem default-dateinamen
    'abgespeichert wurde.
    'wie kommen wir jetzt zum speichern unter dialog
    'OK doc hat ein Exists property
    'Exists liefert false:
    ' * wenn Pathfilename = nothing
    ' * wenn Pathfilename nicht existiert
End Sub

Public Property Get Exists() As Boolean
    If m_pfn Is Nothing Then Exit Property
    Exists = m_pfn.Exists
End Property

Public Property Get PathFilename() As PathFilename
    Set PathFilename = m_pfn
End Property

Public Sub Add(aObj As Object)
    If Not aObj Is Nothing Then
        Select Case True
        Case TypeOf aObj Is Address:   Addresses.Add aObj
        Case TypeOf aObj Is City:      Cities.Add aObj
        Case TypeOf aObj Is Country:   Countries.Add aObj
        Case TypeOf aObj Is Person:    Persons.Add aObj
        Case TypeOf aObj Is TelefonNr: TelefonNrs.Add aObj
        Case TypeOf aObj Is Test:      Tests.Add aObj
        'case else
        End Select
    End If
    'AllObjects dient imom für Paste
    AllObjects.Add aObj
End Sub

Public Sub Remove(aObj As Object)
    If Not aObj Is Nothing Then
        Select Case True
        Case TypeOf aObj Is Country:   Countries.RemoveObj aObj
        Case TypeOf aObj Is City:      Cities.RemoveObj aObj
        Case TypeOf aObj Is Address:   Addresses.RemoveObj aObj
        Case TypeOf aObj Is TelefonNr: TelefonNrs.RemoveObj aObj
        Case TypeOf aObj Is Person:    Persons.RemoveObj aObj
        Case TypeOf aObj Is Test:      Tests.RemoveObj aObj
        'case else
        End Select
        AllObjects.RemoveObj aObj
    End If
End Sub

'Public Property Get Item(Index As Long) 'As Object
'    Set Item = AllObjects(Index)
'End Property
Public Sub Paste(aDoc As Document)
'OK, so nciht, das musskomplett überarbeitet werden
'wir brauchen eine All-List
'und es darf nur das erste Objekt betrachtet werden,
'weil nur das erste OBjekt is the objekct desired
'das beabsichtigte Objekt beim Copy-Paste Vorgang
'alle anderen Objekte sind nur der Vollständigkeit halber drin.
'eine Person die Mutter und Vater hat, weiß über die Children die Verbindung
'
'

    Dim i As Long
    Dim co As Country, ci As City, ad As Address, tn As TelefonNr, pr As Person
    With aDoc
        For i = 0 To .Countries.Count - 1
            Set co = .Countries.Item(i)
            If Not Me.Countries.ContainsEqualObj(co) Then
                Me.Countries.Add co
            End If
        Next
        For i = 0 To .Cities.Count - 1
            Set ci = .Cities.Item(i)
            If Not Me.Cities.ContainsEqualObj(ci) Then
                Me.Cities.Add ci
            End If
        Next
        For i = 0 To .Addresses.Count - 1
            Set ad = .Addresses.Item(i)
            If Not Me.Addresses.ContainsEqualObj(ad) Then
                Me.Cities.Add ad
            End If
        Next
        For i = 0 To .TelefonNrs.Count - 1
            Set tn = .TelefonNrs.Item(i)
            If Not Me.TelefonNrs.ContainsEqualObj(tn) Then
                Me.TelefonNrs.Add tn
            End If
        Next
        For i = 0 To .Persons.Count - 1
            Set pr = .Persons.Item(i)
            If Not Me.Persons.ContainsEqualObj(pr) Then
                Me.Persons.Add pr
            End If
        Next
    End With
End Sub

'These two are Public just for debugging reasons
Friend Function ReadFileContent(sFileContent_out As String) As Boolean
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open m_pfn.Name For Binary Access Read As FNr
    sFileContent_out = Space$(LOF(FNr))
    Get FNr, , sFileContent_out
Finally:
    Close FNr
    ReadFileContent = Err = 0
End Function
Friend Function SaveFileContent(sFileContent As String) As Boolean
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    If m_pfn.Exists Then m_pfn.Delete
    Open m_pfn.Name For Binary As FNr
    Put FNr, , sFileContent
Finally:
    Close FNr
    SaveFileContent = Err = 0
End Function
'OK hier wird nur die Datei eingelesen
'geparst wird dann komplett im MXParser der wird hier aufgerufen und Doc übergibt sich selber

'Public Sub Parse(aPFN As String)
'    Dim sLines() As String: ReadFileLines aPFN, sLines
'    XParser.ParseLines sLines, Me
'End Sub
'
'Oh mann eh so ist das ja mal wieder totaler Müll!
Private Sub LoadFile() '_aPFN As String)
    'pfn = aPFN
    Dim sContent As String
    ReadFileContent sContent
    Parse sContent
End Sub
Public Sub SaveFile(Optional ByVal aPFN As PathFilename = Nothing)
    'If Len(aPFN) Then pfn = aPFN 'speichern unter
    'If Len(pfn) = 0 Then Exit Sub
    If Not aPFN Is Nothing Then Set m_pfn = aPFN
    Serial
    Dim sContent As String: sContent = Serializer.LStr.ToStr
    SaveFileContent sContent
End Sub

Public Sub Parse(sContent As String)
    Dim sLines() As String: sLines = Split(sContent, vbCrLf)
    XParser.ParseLines sLines, Me
End Sub

Public Sub Serial()
    Serializer.Init
    Serializer.Serial_Doc Me
End Sub

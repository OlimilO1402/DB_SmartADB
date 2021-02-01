VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Info Small&Smart"
   ClientHeight    =   4215
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8415
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDescription 
      Height          =   2775
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   840
      Width           =   6015
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Systeminfo..."
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1359.015
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   1411.69
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Small && Smart Address Database"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   3150
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.2019.09.04"
      Height          =   225
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   2100
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Registrierungsschlüssel-Sicherheitsoptionen...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Registrierungsschlüssel-Stammtypen...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
Const REG_DWORD = 4                      ' 32-Bit-Zahl

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Info zu " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    TxtDescription.Text = _
    "Object-Oriented Relational Address Database," & vbCrLf & _
    "without ever using a data base driver." & vbCrLf & _
    "Works like every typical windows desktop-applikation," & vbCrLf & _
    "shows the following features, how to . . .: " & vbCrLf & _
    " * serialize complex class-hierarchies in " & vbCrLf & _
    "   file format IFC aka STEP aka ISO10303-21 " & vbCrLf & _
    " * implement cut&copy&paste" & vbCrLf & _
    " * implement undo&redo" & vbCrLf & _
    " * enable visual styles with manifest " & vbCrLf & _
    " * have your own file-icon and register it" '& vbCrLf
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Versuchen, den Systeminfo-Programmpfad/-namen aus der Registrierung abzurufen...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Versuchen, nur den Systeminfo-Programmpfad aus der Registrierung abzurufen...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Überprüfen, ob bekannte 32-Dateiversion vorhanden ist
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Fehler - Datei wurde nicht gefunden...
        Else
            GoTo SysInfoErr
        End If
    ' Fehler - Registrierungseintrag wurde nicht gefunden...
    Else
        GoTo SysInfoErr
    End If
    
    If Len(SysInfoPath) = 0 Then SysInfoPath = "MSInfo32.exe"
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Systeminformationen sind momentan nicht verfügbar", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Schleifenzähler
    Dim rc As Long                                          ' Rückgabe-Code
    Dim hKey As Long                                        ' Zugriffsnummer für einen offenen Registrierungsschlüssel
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschlüssels
    Dim tmpVal As String                                    ' Temporärer Speicher eines Registrierungsschlüsselwertes
    Dim KeyValSize As Long                                  ' Größe der Registrierungsschlüsselvariablen
    '------------------------------------------------------------
    ' Registrierungsschlüssel unter KeyRoot {HKEY_LOCAL_MACHINE...} öffnen
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschlüssel öffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
    
    tmpVal = String$(1024, 0)                             ' Platz für Variable reservieren
    KeyValSize = 1024                                       ' Größe der Variable markieren
    
    '------------------------------------------------------------
    ' Registrierungsschlüsselwert abrufen...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)  ' Schlüsselwert abrufen/erstellen
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 fügt null-terminierte Zeichenfolge hinzu...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
    Else                                                    ' Keine null-terminierte Zeichenfolge für WinNT...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
    End If
    '------------------------------------------------------------
    ' Schlüsselwerttyp für Konvertierung bestimmen...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Datentypen durchsuchen...
    Case REG_SZ                                             ' Zeichenfolge für Registrierungsschlüsseldatentyp
        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
    Case REG_DWORD                                          ' Registrierungsschlüsseldatentyp DWORD
        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
    End Select
    
    GetKeyValue = True                                      ' Erfolgreiche Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
    Exit Function                                           ' Beenden
    
GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
    KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
    GetKeyValue = False                                     ' Fehlgeschlagene Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
End Function

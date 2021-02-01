Attribute VB_Name = "Registry"
Option Explicit ' Zeilen: 439
Private Const REG_NONE                As Long = 0 ' No value type
Private Const REG_SZ                  As Long = 1 ' Unicode null terminated string
Private Const REG_EXPAND_SZ           As Long = 2 ' Unicode null terminated string (with environment variable references)
Private Const REG_BINARY              As Long = 3 ' Free form binary
Private Const REG_DWORD               As Long = 4 ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4 ' ja, zweimal 4 ' 32-bit Nummer wie REG_DWORD
Private Const REG_DWORD_BIG_ENDIAN    As Long = 5 ' 32-bit Nummer
Private Const REG_LINK                As Long = 6 ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ            As Long = 7 ' Multiple Unicode strings
Private Const REG_OPTION_NON_VOLATILE As Long = &H0
Private Const REG_CREATED_NEW_KEY     As Long = &H1

Private Const GW_CHILD     As Long = 5
Private Const GW_HWNDFIRST As Long = 0
Private Const GW_HWNDLAST  As Long = 1
Private Const GW_HWNDNEXT  As Long = 2
Private Const GW_HWNDPREV  As Long = 3
Private Const GW_OWNER     As Long = 4
Private Const GW_MAX       As Long = 5
Private Const MaxBuff      As Long = 255

Public Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
'  HKCR = HKEY_CLASSES_ROOT
'  HKCU = HKEY_CURRENT_USER
'  HKLM = HKEY_LOCAL_MACHINE
'  HKU = HKEY_USERS
'  HKPD = HKEY_PERFORMANCE_DATA
'  HKCC = HKEY_CURRENT_CONFIG
'  HKDD = HKEY_DYN_DATA
End Enum

Private Const ERROR_SUCCESS            As Long = 0&
Private Const SYNCHRONIZE              As Long = &H100000
Private Const READ_CONTROL             As Long = &H20000
Private Const STANDARD_RIGHTS_ALL      As Long = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE  As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ     As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_WRITE    As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE          As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS   As Long = &H8
Private Const KEY_NOTIFY               As Long = &H10
Private Const KEY_SET_VALUE            As Long = &H2
Private Const KEY_CREATE_SUB_KEY       As Long = &H4
Private Const KEY_CREATE_LINK          As Long = &H20

Private Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or _
                                         KEY_QUERY_VALUE Or _
                                         KEY_SET_VALUE Or _
                                         KEY_CREATE_SUB_KEY Or _
                                         KEY_ENUMERATE_SUB_KEYS Or _
                                         KEY_NOTIFY Or _
                                         KEY_CREATE_LINK) And (Not SYNCHRONIZE))
                                          
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or _
                                   KEY_QUERY_VALUE Or _
                                   KEY_ENUMERATE_SUB_KEYS Or _
                                   KEY_NOTIFY) And (Not SYNCHRONIZE))
                                   
Private Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or _
                                    KEY_SET_VALUE Or _
                                    KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE              As Long = (KEY_READ)

'private Member der Klasse
Private mCurrentKey  As hKey  'Handle auf den Aktuellen Key
Private mCurrentPath As String 'Pfad zum aktuellen Key als String
Private mLazyWrite   As Boolean
Private mRootKey     As hKey
Private mAccess      As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal RtKey As hKey, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal RtKey As hKey, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Any) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal RtKey As hKey) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal RtKey As hKey, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal RtKey As hKey, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal RtKey As hKey, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_DWord Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal RtKey As hKey) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal RtKey As hKey, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal RtKey As hKey, ByVal lpValueName As String) As Long

Public Sub Init()
    mRootKey = HKEY_CURRENT_USER
    mCurrentKey = mRootKey
    mCurrentPath = ""
    mLazyWrite = True
    mAccess = KEY_ALL_ACCESS
End Sub

'registriert Dateiverknüpfung, nur wenn Ini-Eintrag gesetzt
'was passiert wenn nicht Adminzugriff ????????????????????? --> testen
Public Sub RegisterShellFileTypes(ByVal FileExtension As String, _
                                  ByVal sAppReg As String, _
                                  ByVal sAppName As String, _
                                  ByVal aPFN As String, _
                                  ByVal lngIconId As Long)
Try: On Error GoTo Finally
    Init
    'Generiert die Assoziation mit der Endung
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey "." & FileExtension
    Registry.WriteString vbNullString, sAppReg
    Registry.CloseKey
    
    'Generiert den neuen Eintrag strAppReg
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg
    Registry.WriteString vbNullString, sAppReg
    Registry.CloseKey
    
    'Speichert Verknüpfung zum Icon
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg & "\DefaultIcon"
    Registry.WriteString vbNullString, """" & aPFN & """" & "," & CStr(lngIconId)
    Registry.CloseKey
    
    'Setzt den ausführenden Pfad für die Anwendung
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg & "\shell\open\command"
    Dim StrKeyVal As String: StrKeyVal = """" & aPFN & """" & " %" & CStr(lngIconId) 'soll man hier quotes dazumachen ?
    Registry.WriteString vbNullString, StrKeyVal

Finally:
    Registry.CloseKey
    'lautlos beenden?
    If Err Then
        MsgBox Err.Description
    Else
        If Err.LastDllError Then
            MsgBox Err.Description
        End If
    End If
End Sub

'registriert Dateiverknüpfung, nur wenn Ini-Eintrag gesetzt
'was passiert wenn nicht Adminzugriff ????????????????????? --> testen
Public Sub UnRegisterShellFileTypes(ByVal FileExtension As String, _
                                     ByVal sAppReg As String) ', _
                                     'ByVal strAppName As String, _
                                     'ByVal aPFN As String, _
                                     'ByVal lngIconId As Long)
Try: On Error GoTo Finally
    Init
    'löscht den ausführenden Pfad für die Anwendung
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg & "\shell\open\command"
    'Dim StrKeyVal As String: StrKeyVal = aPFN.Quoted & " %" & CStr(lngIconId) 'soll man hier quotes dazumachen ?
    'Registry.WriteString vbNullString, StrKeyVal
    Registry.CloseKey
    
    'löscht die Verknüpfung zum Icon
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg & "\DefaultIcon"
    'Registry.WriteString vbNullString, aPFN.Quoted & "," & CStr(lngIconId)
    Registry.CloseKey
    
    'löscht den Eintrag strAppReg
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg
    'Registry.WriteString vbNullString, strAppReg
    Registry.CloseKey
    
    'löscht die Assoziation mit der Endung
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey "." & FileExtension
    'Registry.WriteString vbNullString, strAppReg
    'Registry.CloseKey
    
Finally:
    Registry.CloseKey
    'lautlos beenden
    '
End Sub



'vv##########################  Properties  ##########################vv
Public Property Get CurrentKey() As hKey    ' nur lesen
  CurrentKey = mCurrentKey
End Property

Public Property Get CurrentPath() As String ' nur lesen
  CurrentPath = mCurrentPath
End Property

Public Property Get LazyWrite() As Boolean
  LazyWrite = mLazyWrite
End Property
Public Property Let LazyWrite(BolVal As Boolean)
  mLazyWrite = BolVal
End Property

Public Property Get RootKey() As hKey
  RootKey = mRootKey
End Property
Public Property Let RootKey(Key As hKey)
  mRootKey = Key
End Property

Public Property Get Access() As Long
  Access = mAccess
End Property
Public Property Let Access(c As Long)
  mAccess = c
End Property
'^^##########################  Properties  ##########################^^

'vv######################  Subs und Functions  ######################vv
'########################### Keys betreffend ##########################
Public Function KeyExists(Key As String) As Boolean
'Frägt ab ob der Key existiert und schließt den Key gleich wieder
  KeyExists = KeyExistsNoClose(Key) ', HandleKey)
  If KeyExists Then
    Call RegCloseKey(mCurrentKey)
  End If
End Function
Private Function KeyExistsNoClose(Key As String) As Boolean ' , ByRef HndKey As Long) As Boolean
'Frägt ab ob der Key existiert ohne den Key zu schließen ', gibt den Handle zurück nö
Dim lResult As Long, HandleKey As Long
  lResult = RegOpenKeyEx(mRootKey, Key, 0&, KEY_READ, HandleKey)
  KeyExistsNoClose = (lResult = ERROR_SUCCESS)
  If KeyExistsNoClose Then
    mCurrentKey = HandleKey
    mCurrentPath = Key
  End If
End Function
Public Function CreateKey(Key As String) As Boolean
'Key kann ein absoluter oder ein relativer Schlüsselname sein.
'Ein absoluter Schlüssel beginnt mit einem Backslash und setzt direkt auf den Hauptschlüssel auf.
'Ein relativer Schlüssel ist ein Unterschlüssel des aktuellen. (ohne BackSlash am Anfang)
Dim lResult As Long, HandleKey As Long, lAction As Long, Class As String '
TryE: On Error GoTo CatchE
  lResult = RegCreateKeyEx(mRootKey, Key, 0, Class, REG_OPTION_NON_VOLATILE, mAccess, 0&, HandleKey, lAction)
  If lResult = ERROR_SUCCESS Then
    mCurrentPath = Key
    mCurrentKey = HandleKey
    If RegFlushKey(mCurrentKey) = ERROR_SUCCESS Then
      Call RegCloseKey(mCurrentKey)
    End If
    CreateKey = (lAction = REG_CREATED_NEW_KEY)
  Else
    CreateKey = False
  End If
  Exit Function
CatchE:
  MsgBox "Key: """ & Key & """ Konnte nicht erzeugt werden."
End Function
Public Function OpenKey(Key As String, CanCreate As Boolean) As Boolean
'öffnet den Key, wenn er nicht da ist, dann wird mit cancreate entschieden ob er erstellt werden soll
  OpenKey = KeyExistsNoClose(Key)
  If Not OpenKey Then
    If CanCreate Then
      OpenKey = CreateKey(Key)
    End If
  End If
End Function
'Private Function KeyExistsNoClose(key As String) As Boolean ' , ByRef HndKey As Long) As Boolean
'Frägt ab ob der Key existiert ohne den Key zu schließen, gibt den Handle zurück
'Dim lResult As Long, HandleKey As Long
'  lResult = RegOpenKeyEx(mRootKey, key, 0, KEY_READ, HandleKey)
'  KeyExistsNoClose = (lResult = ERROR_SUCCESS)
'  If KeyExistsNoClose Then
'    mCurrentKey = HandleKey
'    mCurrentPath = key
'  End If
'End Function

Public Function OpenKeyReadOnly(Key As String) As Boolean
Dim mA As Long
Try: On Error GoTo Catch
  mA = mAccess
  mAccess = KEY_READ
  OpenKeyReadOnly = KeyExistsNoClose(Key)
  mAccess = mA
  Exit Function
Catch:
  MsgBox "Key: """ & Key & """ Konnte nicht mit ReadOnly geöffnet werden."
End Function
Public Sub CloseKey()
'Diese Methode schreibt den aktuellen Schlüssel in die Registrierdatenbank und schließt ihn.
  'Key
  Call RegCloseKey(mCurrentKey)
  Call RegCloseKey(mRootKey)
End Sub
Public Sub DeleteKey(Key As String)
Dim lResult As Long
  lResult = RegDeleteKey(mRootKey, Key)
  'DeleteKey = (lResult = ERROR_SUCCESS)
End Sub
Public Sub MoveKey(OldName As String, NewName As String, Delete As Boolean)
  'ToDo
End Sub
'vv########################  Für Hive-File  ###################################
Public Function LoadKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function
Public Function ReplaceKey(Key As String, FileName As String, BackUpFileName As String) As Boolean
  'ToDo
End Function
Public Function RestoreKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function
Public Function SaveKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function

Public Function UnLoadKey(Key As String) As Boolean
  'ToDo
End Function
Public Function HasSubKeys() As Boolean
  'ToDo
End Function

'########################### Values betreffend ##########################
Public Function ValueExists(Name As String) As Boolean
Dim HandleVal As Long
  ValueExists = ValueExistsNoClose(Name) ', HandleVal)
  If ValueExists Then Call RegCloseKey(HandleVal)
End Function
Private Function ValueExistsNoClose(Name As String) As Boolean ', HandleVal As Long) As Boolean
Dim lResult As Long, HandleVal As Long, dwType As Long, puffergröße As Long
  lResult = RegQueryValueEx(HandleVal, Name, 0&, dwType, ByVal 0&, puffergröße)
  ValueExistsNoClose = (lResult = ERROR_SUCCESS)
  If ValueExistsNoClose Then
    mCurrentKey = HandleVal
  End If
End Function
Public Function DeleteValue(Name As String) As Boolean
Dim lResult As Long, HandleKey As Long
  DeleteValue = ValueExistsNoClose(Name)
  If DeleteValue Then
    lResult = RegDeleteValue(mCurrentKey, Name)
    DeleteValue = (lResult = ERROR_SUCCESS)
  End If
  Call RegCloseKey(mCurrentKey)
End Function

'End Function
Public Sub RenameValue(OldName As String, NewName As String)
  'ToDo
End Sub
'######################### Get- Subs und Functions ########################
'Public Function GetDataInfo(ValueName As String, value As RegDataInfo) As Boolean
'
'End Function
Public Function GetDataSize(ValueName As String) As Integer
  'ToDo
End Function
Public Sub GetDataType(ValueName As String)
  'ToDo
End Sub
'Public Function GetKeyInfo(value As RegKeyInfo) As Boolean
'
'End Function
Public Sub GetKeyNames(StrCol As Collection)
  'ToDo
End Sub
Public Sub GetValueNames(StrCol As Collection)
  'ToDo
End Sub

'########################### Spezial #####################################
Public Function RegistryConnect(UNCName As String) As Boolean
'Die Methode richtet eine Verbindung zur Registrierdatenbank auf einem anderen Computer ein.
  'ToDo
End Function

'vv############################### ReadFunctions und WriteSubs ########################vv
Public Function ReadCurrency(Name As String) As Currency
  'ToDo
End Function
Public Sub WriteCurrency(Name As String, Value As Currency)
  'ToDo
End Sub
Public Function ReadBinaryData(Name As String, Buffer As Variant, BufSize As Integer) As Integer
  'ToDo
End Function
Public Sub WriteBinaryData(Name As String, Buffer As Variant, BufSize As Integer)
  'ToDo
End Sub
Public Function ReadBool(Name As String) As Boolean
  'ToDo
End Function
Public Sub WriteBool(Name As String, Value As Boolean)
  'ToDo
End Sub

Public Function ReadDate(Name As String) As Date
  'ToDo
End Function
Public Sub WriteDate(Name As String, Value As Date)
  'ToDo
End Sub
Public Function ReadDateTime(Name As String) As Date
  'ToDo
End Function
Public Sub WriteDateTime(Name As String, Value As Date)
  'ToDo
End Sub
Public Function ReadTime(Name As String) As Date
  'ToDo
End Function
Public Sub WriteTime(Name As String, Value As Date)
  'ToDo
End Sub
Public Function ReadFloat(Name As String) As Double
  'ToDo
End Function
Public Sub WriteFloat(Name As String, Value As Double)
  'ToDo
End Sub
Public Function ReadInteger(Name As String) As Long
Dim LngVal As Long
  If GetValue(mCurrentPath, Name, LngVal) Then
    ReadInteger = LngVal
  Else
    MsgBox "Wert: """ & Name & """ konnte nicht gelesen werden"
  End If
End Function
Public Sub WriteInteger(Name As String, Value As Long)
Dim LngVal As Long
  LngVal = Value
  If Not SetValue(mRootKey, mCurrentPath, Name, LngVal) Then
    MsgBox "Wert: """ & Name & """ konnte nicht geschrieben werden"
  End If
End Sub
Public Function ReadString(Name As String) As String
Dim StrVal As String
  If GetValue(mCurrentPath, Name, StrVal) Then
    ReadString = StrVal
  Else
    MsgBox "Wert: """ & Name & """ konnte nicht gelesen werden"
  End If
End Function
Public Sub WriteString(Name As String, Value As String)
Dim StrVal As String
  StrVal = Value
  If Not SetValue(mRootKey, mCurrentPath, Name, StrVal) Then
    MsgBox "Wert: """ & Name & """ konnte nicht geschrieben werden"
  End If
End Sub
Public Sub WriteExpandString(Name As String, Value As String)
  'ToDo
End Sub

Private Function GetValue(Key As String, ValNam As String, VarVal As Variant) As Boolean
Dim lResult As Long, dwType As Long
Dim zw As Long, puffergröße As Long
Dim puffer As String
  'GetValue =  KeyExistsNoClose(Key) ', HandleKey)
  If Not KeyExistsNoClose(Key) Then
    Exit Function
  End If
  lResult = RegQueryValueEx(mCurrentKey, ValNam, 0&, dwType, ByVal 0&, puffergröße)
  GetValue = (lResult = ERROR_SUCCESS)
  If lResult <> ERROR_SUCCESS Then Exit Function ' Feld existiert nicht
  Select Case dwType
    Case REG_SZ       ' nullterminierter String
      puffer = Space$(puffergröße + 1)
      lResult = RegQueryValueEx(mCurrentKey, ValNam, 0&, dwType, ByVal puffer, puffergröße)
      GetValue = (lResult = ERROR_SUCCESS)
      If lResult <> ERROR_SUCCESS Then Exit Function ' Fehler beim auslesen des Feldes
      Dim plen As Long
      plen = InStr(1, puffer, vbNullChar) - 1
      If plen > 0 Then
        VarVal = Left$(puffer, plen)
      End If
    Case REG_DWORD     ' 32-Bit Number   !!!! Word
      puffergröße = 4      ' = 32 Bit
      lResult = RegQueryValueEx(mCurrentKey, ValNam, 0&, dwType, zw, puffergröße)
      GetValue = (lResult = ERROR_SUCCESS)
      If lResult <> ERROR_SUCCESS Then Exit Function ' Fehler beim auslesen des Feldes
      VarVal = zw
      ' Hier könnten auch die weiteren Datentypen behandelt werden, soweit dies sinnvoll ist
  End Select
  Call RegCloseKey(mCurrentKey)
  GetValue = (lResult = ERROR_SUCCESS)
End Function

'Private Function SetValue(key As String, ValNam As String, VarVal As Variant) As Boolean
'Dim lResult As Long, l As Long, HandleKey As Long
'Dim s As String
'  If Not KeyExistsNoClose(key) Then
'    Exit Function
'  End If
'  Select Case VarType(VarVal)
'    Case vbInteger, vbLong
'      l = CLng(VarVal)
'      HandleKey = mCurrentKey
'      lResult = RegSetValueEx_DWd(HandleKey, ValNam, 0&, REG_DWORD, l, 4)
'    Case vbString&
'      s = CStr(VarVal)
'      lResult = RegSetValueEx_Str(mCurrentKey, ValNam, 0&, REG_SZ, s, Len(s) + 1)    ' +1 für die Null am Ende
'    ' Hier können noch weitere Datentypen umgewandelt bzw. gespeichert werden
'  End Select
'  Call RegCloseKey(mCurrentKey)
'  SetValue = (lResult = ERROR_SUCCESS)
'End Function

Private Function SetValue(root As Long, Key As String, field As String, Value As Variant) As Boolean
Dim lResult As Long, keyhandle As Long
Dim s As String, L As Long
    lResult = RegOpenKeyEx(root, Key, 0, KEY_ALL_ACCESS, keyhandle)
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
        Exit Function
    End If
    Select Case VarType(Value)
        Case vbInteger, vbLong
            L = CLng(Value)
            lResult = RegSetValueEx_DWord(keyhandle, field, 0, REG_DWORD, L, 4)
        Case vbString
            s = CStr(Value)
            lResult = RegSetValueEx_String(keyhandle, field, 0, REG_SZ, s, Len(s) + 1)    ' +1 für die Null am Ende
        ' Hier können noch weitere Datentypen umgewandelt bzw. gespeichert werden
    End Select
    RegCloseKey keyhandle
    SetValue = (lResult = ERROR_SUCCESS)
End Function



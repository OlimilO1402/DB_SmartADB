Attribute VB_Name = "MRegistry"
Option Explicit
 
' zunächst alle benötigten API-Deklarationen
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
 
Private Const KEY_ALL_ACCESS     As Long = &H3F&
Private Const KEY_SET_VALUE      As Long = &H2&
Private Const KEY_CREATE_SUB_KEY As Long = &H4&
 
Private Const REG_PRIMARY_KEY    As String = "Software\Classes\"
Private Const REG_SHELL_KEY      As String = "Shell\"
Private Const REG_SHELL_OPEN_KEY As String = "Open\"
Private Const REG_SHELL_OPEN_COMMAND_KEY As String = "Command"
 
Private Const REG_SZ             As Long = 1&
Private Const REG_OPTION_NON_VOLATILE As Long = 0&
 
Private Const ERROR_SUCCESS      As Long = 0&
 
Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Boolean
End Type
 
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
    ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" ( _
    ByVal hKey As Long, ByVal lpSubKey As Any, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
 
Private Function OpenKey(lhKey As Long, SubKey As String, ulOptions As Long) As Long
    Dim lhKeyOpen As Long
    Dim lResult As Long: lResult = RegOpenKeyEx(lhKey, SubKey, 0, ulOptions, lhKeyOpen)
    OpenKey = IIf(lResult <> ERROR_SUCCESS, 0, lhKeyOpen)
End Function
 
Private Function CreateKey(lhKey As Long, SubKey As String, NewSubKey As String) As Boolean
    Dim lhKeyOpen As Long: lhKeyOpen = OpenKey(lhKey, SubKey, KEY_CREATE_SUB_KEY)
    Dim Security  As SECURITY_ATTRIBUTES, lhKeyNew As Long, lDisposition As Long
    Dim lResult   As Long:   lResult = RegCreateKeyEx(lhKeyOpen, NewSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew, lDisposition)
    CreateKey = lResult = ERROR_SUCCESS
    If CreateKey Then RegCloseKey lhKeyNew
    RegCloseKey lhKeyOpen
End Function
 
Private Function SetValue(lhKey As Long, SubKey As String, sValue As String) As Boolean
    Dim lByte     As Long: lByte = Len(sValue)
    Dim lTyp      As Long: lTyp = REG_SZ
    Dim lhKeyOpen As Long: lhKeyOpen = OpenKey(lhKey, SubKey, KEY_SET_VALUE)
    Dim lResult   As Long: lResult = RegSetValue(lhKey, SubKey, lTyp, sValue, lByte)
    SetValue = lResult = ERROR_SUCCESS
    If Not SetValue Then RegCloseKey lhKeyOpen
End Function
 
' Datei-Verknüpfung in der Registry speichern
' sFileExt = Dateiendung (z.B. .txt)
' sFileDescr = Beschreibung (z.B. Textdokument)
' sAppID = Programm-Kennung (z.B. Mein Texteditor)
' sOpenCmd = vollständiger Dateiname der Anwendung
' inkl. Parameter %1
' (z.B. App.Path & "\" & App.EXEName & ".exe %1"
Public Function RegisterFile(sFileExt As String, sFileDescr As String, sAppID As String, sOpenCmd As String) As Boolean
    Dim bSuccess As Boolean
    'bSuccess = False
    Dim hKey As Long: hKey = HKEY_LOCAL_MACHINE
    ' File-Extension
    If CreateKey(hKey, REG_PRIMARY_KEY, sFileExt) Then
        If SetValue(hKey, REG_PRIMARY_KEY & sFileExt, sAppID) Then
            ' AppID
            If CreateKey(hKey, REG_PRIMARY_KEY, sAppID) Then
                ' AppDescription
                If SetValue(hKey, REG_PRIMARY_KEY & sAppID, sFileDescr) Then
                    ' OpenCommand
                    If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY) Then
                        bSuccess = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY, sOpenCmd)
                    End If
                End If
            End If
        End If
    End If
    RegisterFile = bSuccess
End Function

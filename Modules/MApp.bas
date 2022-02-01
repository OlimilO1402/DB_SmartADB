Attribute VB_Name = "Application"
Option Explicit

Private Const GW_OWNER As Long = 4

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetInstanceEx Lib "DirectCOM" (StrPtr_FName As Long, StrPtr_ClassName As Long, ByVal UseAlteredSearchPath As Boolean) As Object

Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Const MyExt   As String = ".icxadb"
Public Const AppExtReg As String = "SmartADB"
Public Const AppName As String = "MBO-Ing.com\SmartADB"
Public Const FIconId As Byte = 1 'for file-icon, prog-Icon-Id = 0
Public Const DefaultFileName As String = "Addresses." & MyExt

Public IsInIDE As Boolean

Public Sub Main()
    
    IsInIDE = IsRunningInIDE(0)
    
    PrepareCOMCTL32dll
    
    'Einstellungen lesen
    Settings.LoadSettings
    
    'Einstellungen anwenden
    If Settings.VisualStylesEnabled Then
        InitCommonControls
    End If
    
    If Settings.SplashScreenEnabled Then
        frmSplash.Show vbModeless, FMain
    End If
    
    If Settings.StartWithHelpTipps Then
        frmHelp.Show vbModeless, FMain
    End If
    
    FMain.Show
End Sub

Public Sub Terminate() 'called from FMain.Form_Unload
    'we also have to save all opened and changed files!
    'how to track data "changed"?
    'maybe the Undo&Redo could help, we must track file savings with UndoRedo
    'in some apps Undo will be cleared by file save, maybe we do not have to do this
    'simply by tracking a version-variable or a changedsincelastsaving-variable in Undo&Redo
    
    Settings.SaveSettings
End Sub

Public Function IsValidFileExt(pfn As PathFileName) As Boolean
    IsValidFileExt = StrComp(pfn.Extension, MyExt, vbTextCompare) = 0
    If IsValidFileExt Then Exit Function
    'hier evtl weitere Dateiformate prüfen falls true dann gleich raus
End Function

Private Sub PrepareCOMCTL32dll()
    If IsInIDE Then
        'Set Constructor = New vbrichclient5.cConstructor
    Else
        Dim pfn As PathFileName: Set pfn = MNew.PathFileName(App.path, "DirectCOM.dll")
        'sVBRC5pfn = App.path
        'If Right(sVBRC5pfn, 1) <> "\" Then sVBRC5pfn = sVBRC5pfn & "\"
        'Debug.Print sVBRC5pfn
        If Not pfn.Exists() Then MsgBox "konnte die Datei nicht finden:" & vbCrLf & pfn.Value
        pfn.FileName = "COMCTL32.ocx"
        If Not pfn.Exists() Then MsgBox "konnte die Datei nicht finden:" & vbCrLf & pfn.Value
        'Set Constructor = GetInstanceEx(StrPtr(sVBRC5pfn), StrPtr("cConstructor"), True)
        Dim pnl As ComctlLib.Panel: Set pnl = GetInstanceEx(StrPtr(pfn.Value), StrPtr("Panel"), True)
    End If
End Sub
Public Sub RegisterExt()
    RegisterShellFileTypes MyExt, AppExtReg, AppName, App.path & "\" & AppName & ".exe", FIconId
End Sub

Public Sub UnRegisterExt()
    UnRegisterShellFileTypes MyExt, AppExtReg
End Sub

Public Function GetFilter() As String
    GetFilter = MyExt & "-Dateien [*" & MyExt & "]|*" & MyExt & "|ifc-Dateien [*.ifc]|*.ifc|Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*"
End Function

Public Function OpenFileName_ShowDlg(ByRef pfn_inout As PathFileName) As VbMsgBoxResult
'Try: On Error GoTo Catch
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    With OFD
        .InitialDirectory = App.path
        If Not pfn_inout Is Nothing Then
            .FileName = pfn_inout.Value
        End If
        .Filter = GetFilter
        OpenFileName_ShowDlg = .ShowDialog
        Set pfn_inout = MNew.PathFileName(.FileName)
    End With
'    With FMain.FileDlg
'        .InitDir = App.Path
'        If Not pfn_inout Is Nothing Then
'            .FileName = pfn_inout.Name
'        End If
'        .Filter = GetFilter
'        .CancelError = True
'        .ShowOpen
'        Set pfn_inout = MNew.PathFileName(.FileName)
'    End With
    'so wie bringen wir jetzt den Dateinamen in die MRUList?
    'die MRUlist ist das eine normale List, ein Stack oder ein Queue
    'die Datei an erster Stelle, soll die da bleiben oder soll die nach unten rutschen
    'wenn eine neue Dateoi geöffnet wird?
    'die vorherige Datei soll nach unten rutschen (bzw von der Zahl her nach oben)
    'd.h. es ist eine Queue?
    'wenn eine Datei geöffnet wird die schon in der MRUlist enthalten
    'die datei an der stelle aus der liste löschen und an erster Stelle wieder einfügen
    'dafür müßte es eigentlich eine eigene Funktion in der List-klasse geben
    'Eintrag an die erste Stelle bzw eine beliebige Stelle verschieben, alle folgenden nach unten schieben
    'wie sieht das mit Queue aus?
    '
    'halt das sollte nicht hier passoieren sondern gleich in Settings!!
    'wird erst in der Form gemacht!
    'Settings.MRUFiles_Add pfn_inout
    
    
    'OpenFileName_ShowDlg = vbOK
'    Exit Function
'Catch: OpenFileName_ShowDlg = vbCancel
End Function

Public Function SaveFileName_ShowDlg(ByRef pfn_inout As PathFileName) As VbMsgBoxResult
'Try: On Error GoTo Catch
    Dim SFD As SaveFileDialog: Set SFD = New SaveFileDialog
    With SFD
        .InitialDirectory = App.path
        If Not pfn_inout Is Nothing Then
            .FileName = pfn_inout.Value
        End If
        .Filter = GetFilter
        SaveFileName_ShowDlg = .ShowDialog
        
        'so
        'Set pfn_inout = MNew.PathFileName(.FileName)
        'oder so?
        If pfn_inout Is Nothing Then
            Set pfn_inout = MNew.PathFileName(.FileName)
        Else
            pfn_inout.New_ .FileName
        End If
    End With
'    With FMain.FileDlg
'        .InitDir = App.Path
'        If Not pfn_inout Is Nothing Then
'            .FileName = pfn_inout.Value
'        End If
'        .Filter = GetFilter
'        .CancelError = True
'        .ShowSave
'        Set pfn_inout = MNew.PathFileName(.FileName)
'    End With
    
    Settings.MRUFiles_Add pfn_inout
    
'    SaveFileName_ShowDlg = vbOK
'    Exit Function
'Catch: SaveFileName_ShowDlg = vbCancel
End Function


'Public Function IsInIDE() As Boolean
'Try: On Error GoTo Catch
'    Debug.Print 1 / 0
'    Exit Function
'Catch: IsInIDE = True
'End Function
Public Function IsRunningInIDE(ByVal ahWnd As Long) As Boolean
'Private Function IsInIDE() As Boolean
    'evtl hier die Länge des Srings vergrößern:
    Dim Buffer     As String:     Buffer = Space(128)
    Dim ParenthWnd As Long:   ParenthWnd = GetWindow(ahWnd, GW_OWNER)
    Call GetClassName(ParenthWnd, Buffer, Len(Buffer))

    IsRunningInIDE = (LCase(Left(Buffer, 11)) = "thundermain")
    'If LCase(Left(Buffer, 11)) = "thundermain" Then
    '    IsRunningInIDE = True
    'Else
    '    IsRunningInIDE = False
    'End If
End Function


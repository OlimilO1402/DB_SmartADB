Attribute VB_Name = "Settings"
Option Explicit
Private m_SplashScreenEnabled  As Boolean
Private m_StartWithHelpTipps   As Boolean
Private m_VisualStylesEnabled  As Boolean
Private m_FileIconIsRegistered As Boolean
Private m_MonitorDiagonal      As Double  'in cm
Private m_FMainWindowState     As FormWindowStateConstants
Private m_MaxMRUFiles          As Byte 'max 255
Private m_MRUFiles             As List
Private Const MinMRUFiles      As Byte = 4

Public Property Let SplashScreenEnabled(ByVal Value As Boolean)
    m_SplashScreenEnabled = Value
End Property
Public Property Get SplashScreenEnabled() As Boolean
    SplashScreenEnabled = m_SplashScreenEnabled
End Property

Public Property Let StartWithHelpTipps(ByVal Value As Boolean)
    m_StartWithHelpTipps = Value
End Property
Public Property Get StartWithHelpTipps() As Boolean
    StartWithHelpTipps = m_StartWithHelpTipps
End Property

Public Property Let VisualStylesEnabled(ByVal Value As Boolean)
    m_VisualStylesEnabled = Value
End Property
Public Property Get VisualStylesEnabled() As Boolean
    VisualStylesEnabled = m_VisualStylesEnabled
End Property

Public Property Let FileIconIsRegistered(ByVal Value As Boolean)
    m_FileIconIsRegistered = Value
End Property
Public Property Get FileIconIsRegistered() As Boolean
    FileIconIsRegistered = m_FileIconIsRegistered
End Property

Public Property Let MonitorDiagonal(ByVal Value As Double)
    m_MonitorDiagonal = Value
End Property
Public Property Get MonitorDiagonal() As Double
    MonitorDiagonal = m_MonitorDiagonal
End Property

Public Property Let FMainWindowState(ByVal Value As FormWindowStateConstants)
    m_FMainWindowState = Value
End Property
Public Property Get FMainWindowState() As FormWindowStateConstants
    FMainWindowState = m_FMainWindowState
End Property

Public Property Let MaxMRUFiles(ByVal Value As Byte)
    If Value < MinMRUFiles Then Value = MinMRUFiles
    m_MaxMRUFiles = Value
End Property
Public Property Get MaxMRUFiles() As Byte
    If m_MaxMRUFiles <= 0 Then m_MaxMRUFiles = MinMRUFiles
    MaxMRUFiles = m_MaxMRUFiles
End Property

Public Property Get MRUFiles() As List
    Set MRUFiles = m_MRUFiles
End Property

Sub MRUFiles_Add(pfn As PathFileName)
    '
    'OK da die MRUlist eine begrenzte liste sein sollte, (oder?)
    'ist es eigentlich egal ob man entweder
    '* am Anfang einfügt und am Ende entfernt oder ob man
    '* am Ende hinzufügt und am Anfang entfernt
    'OK wer soltle dafür verantwortlich sein dass die MRUList ins Menü eingefügt wird
    'ist das eine Sache der Settings oder der MRUList oder liegt die verantwortung beim MainForm?
    'OK de Anzeige ist eine Sache des Views, also vom MainForm
    
    'Dim pfn As PathFileName: Set pfn = MNew.PathFileName(FNm)
    'If m_MRUFiles.Contains(pfn, True) Then
    m_MRUFiles.RemoveObj pfn
    m_MRUFiles.Insert 0, pfn
    If MaxMRUFiles < m_MRUFiles.Count Then
        'jetzt die letzte Datei löschen
        m_MRUFiles.Remove m_MRUFiles.Count - 1
    End If
    'so und jetzt auch anzeigen
    FMain.MRUFiles_FillMenu m_MRUFiles
End Sub

Function MRUFiles_Load() As List
    Dim apnam As String: apnam = Application.AppName
    Dim scnam As String: scnam = "Appsettings"

    Dim i As Long, n As Long
    n = GetSetting(apnam, scnam, "nMRUFiles", 0)
    Dim mru As List: Set mru = MNew.List(vbObject) ', , True)
    For i = 0 To n - 1
        Dim pfn As String: pfn = GetSetting(apnam, scnam, "MRUFile" & i, "")
        If Len(pfn) Then
            mru.Add MNew.PathFileName(pfn)
        End If
    Next
    Set MRUFiles_Load = mru
End Function

Sub MRUFiles_Save(mru As List)
    Dim apnam As String: apnam = Application.AppName
    Dim scnam As String: scnam = "Appsettings"
    'the list of the "most recently used" files
    Dim i As Long, n As Long: n = m_MRUFiles.Count
    SaveSetting apnam, scnam, "nMRUFiles", n
    For i = 0 To n - 1
        Dim pfn As PathFileName: Set pfn = m_MRUFiles.Item(i)
        SaveSetting apnam, scnam, "MRUFile" & i, pfn.Name
    Next
End Sub

Public Sub LoadSettings()
    
    Dim apnam As String: apnam = Application.AppName
    Dim scnam As String: scnam = "Appsettings"
    m_SplashScreenEnabled = GetSetting(apnam, scnam, "SplashScreenEnabled", True)
    m_StartWithHelpTipps = GetSetting(apnam, scnam, "StartWithHelpTipps", True)
    m_VisualStylesEnabled = GetSetting(apnam, scnam, "VisualStylesEnabled", True)
    m_FileIconIsRegistered = GetSetting(apnam, scnam, "FileIconIsRegistered", False)
    m_MonitorDiagonal = GetSetting(apnam, scnam, "MonitorDiagonal", 0)
    m_FMainWindowState = GetSetting(apnam, scnam, "FMainWindowState", FormWindowStateConstants.vbNormal)
    m_MaxMRUFiles = GetSetting(apnam, scnam, "MaxMRUFiles", MinMRUFiles)
    
    Set m_MRUFiles = MRUFiles_Load
    FMain.MRUFiles_FillMenu m_MRUFiles
    
End Sub

Public Sub SaveSettings()

    Dim apnam As String: apnam = Application.AppName
    Dim scnam As String: scnam = "Appsettings"
    SaveSetting apnam, scnam, "SplashScreenEnabled", m_SplashScreenEnabled
    SaveSetting apnam, scnam, "StartWithHelpTipps", m_StartWithHelpTipps
    SaveSetting apnam, scnam, "VisualStylesEnabled", m_VisualStylesEnabled
    SaveSetting apnam, scnam, "FileIconIsRegistered", m_FileIconIsRegistered
    SaveSetting apnam, scnam, "FMainWindowState", m_FMainWindowState
    SaveSetting apnam, scnam, "MonitorDiagonal", m_MonitorDiagonal
    
    SaveSetting apnam, scnam, "MaxMRUFiles", m_MaxMRUFiles
    
    MRUFiles_Save m_MRUFiles
    
End Sub

Public Sub DeleteAllSettings()
'
End Sub

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Optionen"
   ClientHeight    =   4935
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox PnlTab1 
      Appearance      =   0  '2D
      BackColor       =   &H80000014&
      BorderStyle     =   0  'Kein
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3615
      ScaleWidth      =   5775
      TabIndex        =   10
      Top             =   600
      Width           =   5775
      Begin VB.TextBox TxtIconId 
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox TxtAppPath 
         Height          =   330
         Left            =   1440
         TabIndex        =   22
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox TxtAppName 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox TxtAppReg 
         Height          =   330
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox TxtFileExt 
         Height          =   330
         Left            =   1440
         TabIndex        =   17
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton BtnUnregisterFileExt 
         Caption         =   "Datei-Endung mit Icon aus Registry löschen"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   5295
      End
      Begin VB.CommandButton BtnRegisterFileExt 
         Caption         =   "Datei-Endung mit Icon registrieren"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   5295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IconId in Res:"
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "App Path:"
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "App Reg:"
         Height          =   225
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Ext:"
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "App Name:"
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start as Administrator!"
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.PictureBox PnlTab2 
      Appearance      =   0  '2D
      BackColor       =   &H80000014&
      BorderStyle     =   0  'Kein
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   6240
      ScaleHeight     =   3615
      ScaleWidth      =   5775
      TabIndex        =   13
      Top             =   600
      Width           =   5775
      Begin VB.TextBox TxtMonitorCM 
         Alignment       =   1  'Rechts
         Height          =   330
         Left            =   4320
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TxtMonitorInch 
         Alignment       =   1  'Rechts
         Height          =   330
         Left            =   3120
         TabIndex        =   30
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TxtMaxMRUFiles 
         Alignment       =   1  'Rechts
         Height          =   330
         Left            =   3840
         TabIndex        =   29
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox ChkVisualStylesEnabled 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Visual Styles aktivieren"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CheckBox ChkShowHelpTippsAtStartup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hilfe (Tipps) bei Programmstart anzeigen."
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox ChkShowSplashAtStartup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Splashscreen bei Programmstart anzeigen"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label LblMonitor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Größe der Monitordiagonale          ""           cm"
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   5250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximale Anzahl Letzter Dateien:"
         Height          =   225
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   3360
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Beispiel 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Beispiel 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Beispiel 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton BtnApply 
      Caption         =   "Übernehmen"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TSOptions 
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Datei-Icon"
            Key             =   "FileIcon"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Einstellungen"
            Key             =   "Settings"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'um Tabs zum Tabstrip hinzuzufügen:
'=> Selektiere das TabStrip und im Eigenschaften-Explorer (ganz oben) wähle "Benutzerdefiniert"
Private MonitorCm As Double

Public Sub FormShow(owner As Form, ByVal TabI As Byte)
    TSOptions.Tabs.Item(TabI).Selected = True
    Me.Show vbModeless, owner
    MoveMe owner
End Sub

Sub MoveMe(owner As Form)
    Dim L As Single: L = owner.Left + owner.Width / 2 - Me.Width / 2
    Dim T As Single: T = owner.Top + owner.Height / 2 - Me.Height / 2
    Me.Move L, T
End Sub

Private Sub Form_Load()
    PnlTab1.ZOrder 0
    
    ChkShowSplashAtStartup.Value = IIf(Settings.SplashScreenEnabled, vbChecked, vbUnchecked)
    ChkShowHelpTippsAtStartup.Value = IIf(Settings.StartWithHelpTipps, vbChecked, vbUnchecked)
    ChkVisualStylesEnabled.Value = IIf(Settings.VisualStylesEnabled, vbChecked, vbUnchecked)
    TxtMaxMRUFiles.Text = Settings.MaxMRUFiles
    
    MonitorCm = Settings.MonitorDiagonal
    LblMonitor.Caption = CalcMonitorDiagonal
    
    TxtMonitorCM.Text = Format(MonitorCm, "0.00")
    TxtMonitorInch.Text = Format(MonitorCm / 2.54, "0.000")
    
    
    TxtFileExt.Text = Application.MyExt
    TxtAppReg.Text = Application.AppReg
    TxtAppName.Text = Application.AppName
    TxtAppPath.Text = App.Path
    TxtIconId.Text = Application.FIconId
End Sub

Function CalcMonitorDiagonal() As String
    Dim s As String
    Dim pw As Double: pw = Screen.Width \ Screen.TwipsPerPixelX
    Dim ph As Double: ph = Screen.Height \ Screen.TwipsPerPixelY
    Dim pd As Double: pd = Math.Sqr(pw * pw + ph * ph)
    
    Dim dpi As Double, dpcm As Double
    
    s = "Auflösung: " & pw & "x" & ph
    Dim r As Double
    Dim c As Double
    If MonitorCm = 0 Then
        MonitorCm = 10
        Do Until (90 < dpi) And (dpi < 110)
            dpcm = pd / MonitorCm
            dpi = dpcm * 2.54
            MonitorCm = MonitorCm + 1
            If MonitorCm >= 1000 Then Exit Do
        Loop
    Else
        dpcm = pd / MonitorCm
        dpi = dpcm * 2.54
    End If
    CalcMonitorDiagonal = s & vbCrLf & "Dpi: " & Format(dpi, "0.00") & vbCrLf & "Dpcm: " & Format(dpcm, "0.00")
End Function
Private Sub Form_Resize()
    'Formular zentrieren
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, 6345
    Me.Move Me.Left, Me.Top, 6345
    
    'Me.ScaleWidth = 417  ' Width = 6345
    'Me.ScaleHeight = 329 'Height = 5370
    'alle TabPages ausrichten
    'PnlTab2.BackColor = Me.BackColor
    PnlTab2.Move PnlTab1.Left, PnlTab2.Top
    
End Sub

Private Sub BtnRegisterFileExt_Click()
    Application.RegisterExt
    Settings.FileIconIsRegistered = True
End Sub

Private Sub BtnUnregisterFileExt_Click()
    Application.UnRegisterExt
    Settings.FileIconIsRegistered = False
End Sub

Private Sub BtnOK_Click()
    'MsgBox "Hier Code eingeben, um Optionen zu setzen und das Dialogfeld zu schließen!"
    BtnApply_Click
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnApply_Click()
    'MsgBox "Hier Code eingeben, um Optionen zu setzen, ohne das Dialogfeld zu schließen!"
    Settings.SplashScreenEnabled = ChkShowSplashAtStartup.Value = vbChecked
    Settings.StartWithHelpTipps = ChkShowHelpTippsAtStartup.Value = vbChecked
    Settings.VisualStylesEnabled = ChkVisualStylesEnabled.Value = vbChecked
    Settings.MaxMRUFiles = TxtMaxMRUFiles.Text
    Settings.MonitorDiagonal = MonitorCm
End Sub

Private Sub TSOptions_Click()
    Select Case TSOptions.SelectedItem.Index
    Case 1: PnlTab1.ZOrder 0
    Case 2: PnlTab2.ZOrder 0
    End Select
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim i As Integer
'    'STRG+TAB-Zugriffsnummer, um zur nächsten Registerkarte zu wechseln
'    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
'        i = tbsOptions.SelectedItem.Index
'        If i = tbsOptions.Tabs.Count Then
'            'Letzte Registerkarte, also wieder mit der ersten beginnen
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
'        Else
'            'Registerkarte hochzählen
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
'        End If
'    End If
'End Sub
'
'Private Sub tbsOptions_Click()
'
'    Dim i As Integer
'    'Steuerelemente der ausgewählten Registerkarte anzeigen und aktivieren
'    'und alle anderen ausblenden und deaktivieren
'    For i = 0 To tbsOptions.Tabs.Count - 1
'        If i = tbsOptions.SelectedItem.Index - 1 Then
'            picOptions(i).Left = 210
'            picOptions(i).Enabled = True
'        Else
'            picOptions(i).Left = -20000
'            picOptions(i).Enabled = False
'        End If
'    Next
'
'End Sub

Private Sub TxtMonitorCM_LostFocus()
    Dim mInch As Double
    If Double_TryParse(TxtMonitorCM.Text, MonitorCm) Then
        mInch = MonitorCm / 2.54
        TxtMonitorCM.Text = Format(MonitorCm, "0.00")
        TxtMonitorInch.Text = Format(mInch, "0.000")
        LblMonitor.Caption = CalcMonitorDiagonal
    End If
End Sub

Private Sub TxtMonitorInch_LostFocus()
    Dim mInch As Double
    If Double_TryParse(TxtMonitorInch.Text, mInch) Then
        MonitorCm = mInch * 2.54
        TxtMonitorCM.Text = Format(MonitorCm, "0.00")
        TxtMonitorInch.Text = Format(mInch, "0.000")
        LblMonitor.Caption = CalcMonitorDiagonal
    End If
End Sub

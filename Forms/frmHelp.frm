VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Hilfe"
   ClientHeight    =   4575
   ClientLeft      =   2295
   ClientTop       =   2325
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4350
      Left            =   2280
      Picture         =   "frmHelp.frx":0000
      ScaleHeight     =   4290
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   3555
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'' Tip-Datenbank im Speicher.
'Dim Tips As New Collection
'
'' Name der Tip-Datei
'Const TIP_FILE = "TIPOFDAY.TXT"
'
'' Index in der Tip-Auflistung, die momentan angezeigt wird.
'Dim CurrentTip As Long
'
'
'Private Sub DoNextTip()
'
'    ' Einen Tip willkürlich auswählen.
'    CurrentTip = Int((Tips.Count * Rnd) + 1)
'
'    ' Oder die Tips der Reihenfolge nach durchgehen.
'
''    CurrentTip = CurrentTip + 1
''    If Tips.Count < CurrentTip Then
''        CurrentTip = 1
''    End If
'
'    ' Tip anzeigen.
'    DisplayCurrentTip
'
'End Sub
'
'Function LoadTips(sFile As String) As Boolean
'    Dim NextTip As String   ' Jeder Tip wird aus der Datei eingelesen.
'    Dim InFile As Integer   ' Descriptor für Datei.
'
'    ' Nächsten freien Datei-Descriptor abrufen.
'    InFile = FreeFile
'
'    ' Sicherstellen, daß eine Datei angegeben wurde.
'    If sFile = "" Then
'        LoadTips = False
'        Exit Function
'    End If
'
'    ' Sicherstellen, daß die Datei vorhanden ist, bevor sie geöffnet wird.
'    If Dir(sFile) = "" Then
'        LoadTips = False
'        Exit Function
'    End If
'
'    ' Auflistung aus einer Text-Datei lesen.
'    Open sFile For Input As InFile
'    While Not EOF(InFile)
'        Line Input #InFile, NextTip
'        Tips.Add NextTip
'    Wend
'    Close InFile
'
'    ' Tips willkürlich anzeigen.
'    DoNextTip
'
'    LoadTips = True
'
'End Function
'
'Private Sub chkLoadTipsAtStartup_Click()
'    ' Speichern, ob dieses Formular beim Start angezeigt werden soll oder nicht
'    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
'End Sub
'
'Private Sub cmdNextTip_Click()
'    DoNextTip
'End Sub
'
'Private Sub cmdOK_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim ShowAtStartup As Long
'
'    ' Feststellen, ob das Dialogfeld beim Start angezeigt werden soll
'    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
'    If ShowAtStartup = 0 Then
'        Unload Me
'        Exit Sub
'    End If
'
'    ' Kontrollkästchen festlegen. Hierdurch wird der Wert in die Registrierung geschrieben
'    Me.chkLoadTipsAtStartup.Value = vbChecked
'
'    ' Randomisieren beginnen
'    Randomize
'
'    ' Tip-Datei lesen und einen Tip willkürlich anzeigen.
'    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
'        lblTipText.Caption = "Die Datei " & TIP_FILE & " wurde nicht gefunden? " & vbCrLf & vbCrLf & _
'           "Textdatei mit dem Namen " & TIP_FILE & " unter Verwendung von NotePad mit 1 Tip pro Zeile erstellen. " & _
'           "Dann im selben Verzeichnis wie die Anwendung ablegen. "
'    End If
'
'
'End Sub
'
'Public Sub DisplayCurrentTip()
'    If Tips.Count > 0 Then
'        lblTipText.Caption = Tips.Item(CurrentTip)
'    End If
'End Sub
Private Sub Form_Load()

End Sub

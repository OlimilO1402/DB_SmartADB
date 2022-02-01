VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   9735
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows-Standard
   Begin ComctlLib.Toolbar TlbMain 
      Height          =   390
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Start a new database"
            Object.Tag             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Open an existing database"
            Object.Tag             =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Save the current database"
            Object.Tag             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Print the database"
            Object.Tag             =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Load the database in text-editor"
            Object.Tag             =   "Edit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cut to clipboard"
            Object.Tag             =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Copy to Clipboard"
            Object.Tag             =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Paste data from clipboard"
            Object.Tag             =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Undo the last edit step"
            Object.Tag             =   "Undo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Redo the last undo step"
            Object.Tag             =   "Redo"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.CommandButton BtnShowCCAT 
         Caption         =   "Land+Stadt+Adresse+TelefonNr"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   0
         Width           =   3255
      End
      Begin VB.CommandButton BtnShowPersons 
         Caption         =   "Personen"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox PnlPersons 
      Appearance      =   0  '2D
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   9225
      ScaleWidth      =   11745
      TabIndex        =   1
      Top             =   360
      Width           =   11775
      Begin VB.PictureBox PnlPersonList 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   9135
         Left            =   0
         ScaleHeight     =   9135
         ScaleWidth      =   5655
         TabIndex        =   24
         Top             =   0
         Width           =   5655
         Begin ComctlLib.Toolbar TlbPersons 
            Height          =   390
            Left            =   120
            TabIndex        =   25
            Top             =   450
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            ImageList       =   "ILDataTlb"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   11
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "New"
                  Object.Tag             =   "AddNew"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Insert"
                  Object.Tag             =   "InsertNew"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Edit"
                  Object.Tag             =   "EditSave"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Delete"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move up"
                  Object.Tag             =   "MoveUp"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move down"
                  Object.Tag             =   "MoveDown"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort up"
                  Object.Tag             =   "SortUp"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort down"
                  Object.Tag             =   "SortDown"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  Object.Width           =   2150
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Suchen"
                  Object.Tag             =   "Search"
                  ImageIndex      =   9
               EndProperty
            EndProperty
            Begin VB.TextBox TxtSearchPerson 
               Height          =   330
               Left            =   3480
               TabIndex        =   53
               Text            =   "Suche"
               Top             =   30
               Width           =   2010
            End
         End
         Begin VB.ListBox LstPerson 
            Height          =   8160
            ItemData        =   "Form1.frx":163F2
            Left            =   120
            List            =   "Form1.frx":163F4
            OLEDragMode     =   1  'Automatisch
            OLEDropMode     =   1  'Manuell
            TabIndex        =   26
            Top             =   840
            Width           =   5520
         End
         Begin VB.Label LblPersons 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Personen"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.PictureBox PnlPersonDetail 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   5760
         ScaleHeight     =   9015
         ScaleWidth      =   5895
         TabIndex        =   5
         Top             =   0
         Width           =   5895
         Begin VB.PictureBox PnlTabChildren 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3975
            Left            =   120
            ScaleHeight     =   3945
            ScaleWidth      =   5385
            TabIndex        =   47
            Top             =   5040
            Width           =   5415
            Begin VB.ListBox LstChildren 
               Height          =   3660
               ItemData        =   "Form1.frx":163F6
               Left            =   0
               List            =   "Form1.frx":163F8
               TabIndex        =   52
               Top             =   0
               Width           =   5520
            End
         End
         Begin VB.PictureBox PnlTabFriends 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   120
            ScaleHeight     =   3825
            ScaleWidth      =   5505
            TabIndex        =   48
            Top             =   5040
            Width           =   5535
            Begin ComctlLib.Toolbar TlbPersonFriends 
               Height          =   390
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   688
               ButtonWidth     =   635
               ButtonHeight    =   582
               ImageList       =   "ILDataTlb"
               _Version        =   327682
               BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
                  NumButtons      =   11
                  BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "New"
                     Object.Tag             =   "AddNew"
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Insert"
                     Object.Tag             =   "InsertNew"
                     ImageIndex      =   2
                  EndProperty
                  BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Edit"
                     Object.Tag             =   "EditSave"
                     ImageIndex      =   3
                  EndProperty
                  BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Delete"
                     Object.Tag             =   "Delete"
                     ImageIndex      =   4
                  EndProperty
                  BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.Tag             =   ""
                     Style           =   3
                     MixedState      =   -1  'True
                  EndProperty
                  BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Move up"
                     Object.Tag             =   "MoveUp"
                     ImageIndex      =   5
                  EndProperty
                  BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Move down"
                     Object.Tag             =   "MoveDown"
                     ImageIndex      =   6
                  EndProperty
                  BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Sort up"
                     Object.Tag             =   "SortUp"
                     ImageIndex      =   7
                  EndProperty
                  BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.ToolTipText     =   "Sort down"
                     Object.Tag             =   "SortDown"
                     ImageIndex      =   8
                  EndProperty
                  BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Object.Tag             =   ""
                     Style           =   3
                     Object.Width           =   2150
                     MixedState      =   -1  'True
                  EndProperty
                  BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Key             =   "Search"
                     Object.ToolTipText     =   "Suchen"
                     Object.Tag             =   "Search"
                     ImageIndex      =   9
                  EndProperty
               EndProperty
               Begin VB.TextBox TxtSearchFriends 
                  Height          =   330
                  Left            =   3480
                  TabIndex        =   51
                  Text            =   "Suche"
                  Top             =   30
                  Width           =   2010
               End
            End
            Begin VB.ListBox LstFriends 
               Height          =   3210
               ItemData        =   "Form1.frx":163FA
               Left            =   0
               List            =   "Form1.frx":163FC
               TabIndex        =   49
               Top             =   390
               Width           =   5520
            End
         End
         Begin ComctlLib.TabStrip TSFamFrnds 
            Height          =   4335
            Left            =   0
            TabIndex        =   46
            Top             =   4560
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   7646
            TabWidthStyle   =   2
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Kinder"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Freunde"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox CmbTelefonNr 
            Height          =   345
            Left            =   1680
            TabIndex        =   14
            Top             =   4080
            Width           =   4095
         End
         Begin VB.ComboBox CmbAddress 
            Height          =   345
            Left            =   1680
            TabIndex        =   13
            Top             =   3600
            Width           =   4095
         End
         Begin VB.ComboBox CmbFather 
            Height          =   345
            Left            =   1680
            TabIndex        =   12
            Top             =   3120
            Width           =   4095
         End
         Begin VB.TextBox TxtBirthD 
            Height          =   345
            Left            =   1680
            TabIndex        =   11
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox TxtFamName 
            Height          =   345
            Left            =   1680
            TabIndex        =   10
            Top             =   1200
            Width           =   4095
         End
         Begin VB.TextBox TxtName2 
            Height          =   345
            Left            =   1680
            TabIndex        =   9
            Top             =   720
            Width           =   4095
         End
         Begin VB.ComboBox CmbMother 
            Height          =   345
            Left            =   1680
            TabIndex        =   8
            Top             =   2640
            Width           =   4095
         End
         Begin VB.TextBox TxtName 
            Height          =   345
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   4095
         End
         Begin VB.ComboBox CmbPersonGender 
            Height          =   345
            Left            =   1680
            TabIndex        =   6
            Top             =   2160
            Width           =   4095
         End
         Begin VB.Label LblTelefonNr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefonnummer"
            Height          =   225
            Left            =   120
            TabIndex        =   23
            Top             =   4080
            Width           =   1365
         End
         Begin VB.Label LblAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adresse"
            Height          =   225
            Left            =   120
            TabIndex        =   22
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label LblFather 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vater"
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   3120
            Width           =   525
         End
         Begin VB.Label LblBirthD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Geburtsdatum"
            Height          =   225
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label LblFamName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Familienname"
            Height          =   225
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label LblName2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "zus. Vornamen"
            Height          =   225
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label LblMother 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mutter"
            Height          =   225
            Left            =   120
            TabIndex        =   17
            Top             =   2640
            Width           =   630
         End
         Begin VB.Label LblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vorname"
            Height          =   225
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label LblGender 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   225
            Left            =   120
            TabIndex        =   15
            Top             =   2160
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox PnlCCAT 
      Appearance      =   0  '2D
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8985
      ScaleWidth      =   11865
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      Begin VB.PictureBox PnlCityTelNr 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   5880
         ScaleHeight     =   9015
         ScaleWidth      =   5775
         TabIndex        =   35
         Top             =   0
         Width           =   5775
         Begin ComctlLib.Toolbar TlbTelNr 
            Height          =   390
            Left            =   120
            TabIndex        =   37
            Top             =   4770
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            ImageList       =   "ILDataTlb"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   11
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "New"
                  Object.Tag             =   "AddNew"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Insert"
                  Object.Tag             =   "InsertNew"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Edit"
                  Object.Tag             =   "EditSave"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Delete"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move up"
                  Object.Tag             =   "MoveUp"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move down"
                  Object.Tag             =   "MoveDown"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort up"
                  Object.Tag             =   "SortUp"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort down"
                  Object.Tag             =   "SortDown"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Suchen"
                  Object.Tag             =   "Search"
                  ImageIndex      =   9
               EndProperty
            EndProperty
            Begin VB.TextBox Text4 
               Height          =   330
               Left            =   3480
               TabIndex        =   45
               Text            =   "Suche"
               Top             =   30
               Width           =   2010
            End
         End
         Begin ComctlLib.Toolbar TlbCity 
            Height          =   390
            Left            =   120
            TabIndex        =   36
            Top             =   450
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            ImageList       =   "ILDataTlb"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   11
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "New"
                  Object.Tag             =   "AddNew"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Insert"
                  Object.Tag             =   "InsertNew"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Edit"
                  Object.Tag             =   "EditSave"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Delete"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move up"
                  Object.Tag             =   "MoveUp"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move down"
                  Object.Tag             =   "MoveDown"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort up"
                  Object.Tag             =   "SortUp"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort down"
                  Object.Tag             =   "SortDown"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Suchen"
                  Object.Tag             =   "Search"
                  ImageIndex      =   9
               EndProperty
            EndProperty
            Begin VB.TextBox Text2 
               Height          =   330
               Left            =   3480
               TabIndex        =   43
               Text            =   "Suche"
               Top             =   30
               Width           =   2010
            End
         End
         Begin VB.ListBox LstCity 
            Height          =   3435
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   5520
         End
         Begin VB.ListBox LstTelefonNr 
            Height          =   3435
            Left            =   120
            TabIndex        =   38
            Top             =   5160
            Width           =   5520
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Städte"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefonnummern"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   40
            Top             =   4440
            Width           =   2100
         End
      End
      Begin VB.PictureBox PnlCountryAddress 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   0
         ScaleHeight     =   9015
         ScaleWidth      =   5775
         TabIndex        =   28
         Top             =   0
         Width           =   5775
         Begin ComctlLib.Toolbar TlbAddress 
            Height          =   390
            Left            =   120
            TabIndex        =   30
            Top             =   4770
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            ImageList       =   "ILDataTlb"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   11
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "New"
                  Object.Tag             =   "AddNew"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Insert"
                  Object.Tag             =   "InsertNew"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Edit"
                  Object.Tag             =   "EditSave"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Delete"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move up"
                  Object.Tag             =   "MoveUp"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move down"
                  Object.Tag             =   "MoveDown"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort up"
                  Object.Tag             =   "SortUp"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort down"
                  Object.Tag             =   "SortDown"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Suchen"
                  Object.Tag             =   "Search"
                  ImageIndex      =   9
               EndProperty
            EndProperty
            Begin VB.TextBox Text3 
               Height          =   330
               Left            =   3480
               TabIndex        =   44
               Text            =   "Suche"
               Top             =   30
               Width           =   2010
            End
         End
         Begin ComctlLib.Toolbar TlbCountry 
            Height          =   390
            Left            =   120
            TabIndex        =   29
            Top             =   450
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            ImageList       =   "ILDataTlb"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   11
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "New"
                  Object.Tag             =   "AddNew"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Insert"
                  Object.Tag             =   "InsertNew"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Edit"
                  Object.Tag             =   "EditSave"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Delete"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move up"
                  Object.Tag             =   "MoveUp"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Move down"
                  Object.Tag             =   "MoveDown"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort up"
                  Object.Tag             =   "SortUp"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Sort down"
                  Object.Tag             =   "SortDown"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.ToolTipText     =   "Suchen"
                  Object.Tag             =   "Search"
                  ImageIndex      =   9
               EndProperty
            EndProperty
            Begin VB.TextBox Text1 
               Height          =   330
               Left            =   3480
               TabIndex        =   42
               Text            =   "Suche"
               Top             =   30
               Width           =   2010
            End
         End
         Begin VB.ListBox LstCountry 
            Height          =   3435
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   5520
         End
         Begin VB.ListBox LstAddress 
            Height          =   3435
            Left            =   120
            TabIndex        =   31
            Top             =   5160
            Width           =   5520
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Länder"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adressen"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   33
            Top             =   4440
            Width           =   1200
         End
      End
   End
   Begin ComctlLib.ImageList ILDataTlb 
      Left            =   9960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":163FE
            Key             =   "AddNew"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":16750
            Key             =   "InsertNew"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":16AA2
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":16DF4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17146
            Key             =   "MoveUp"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17698
            Key             =   "MoveDown"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17BEA
            Key             =   "SortUp"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17F3C
            Key             =   "SortDown"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1828E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":185E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":18932
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":18C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":18FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":19328
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1967A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":199CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":19D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1A070
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1A3C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Datei"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Neu"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Öffnen"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Speichern"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Speichern &unter..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecentFiles 
         Caption         =   "Letzte Dateien"
         Begin VB.Menu mnuFileRecentFile 
            Caption         =   "1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "B&eenden"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "B&earbeiten"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Rückgängig"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Wiederherstellen"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Auss&chneiden"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Kopieren"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "E&infügen"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEditor 
         Caption         =   "&Editor"
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuExtraRegisterFileIcon 
         Caption         =   "Datei-Icon registrieren"
      End
      Begin VB.Menu mnuExtraOptions 
         Caption         =   "Optionen"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpShow 
         Caption         =   "&Hilfe"
      End
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Doc As Document
Private m_obj As Object 'current selected object ready for copy to clipboard
'
Private m_SplPersons As Splitter

'Hja wie machen wir das jetzt?
'blöd Lightswitch läßt sich nicht mehr starten?
'mal sehen ob man eine LS-Anwendung übertratgen kann
'OK haben LS zum Laufen gebracht mit Visual Studio 2012 Premium, hat scheinbar sofort meine liz erkannt (?)
'lief sofort

' ############################## '         Form          ' ############################## '
Private Sub Form_Load()
    Set m_SplPersons = New Splitter
    m_SplPersons.New_ False, Me, PnlPersons, "Splitter1", PnlPersonList, PnlPersonDetail
    m_SplPersons.LeftTopPos = PnlPersonList.Width
    m_SplPersons.BorderStyle = bsXPStyl
    
    'Me.Caption = App.ProductName & " - " & Application.DefaultFileName '& "]"
    Dim pfn As PathFileName
    If Len(Command$) Then Set pfn = MNew.PathFileName(Command$)
'        if application.IsValidFileExt(pfn)
'        Set m_Doc = MNew.Document()
'        UpdateView
'        UpdateFMainCaption
'    Else
'        Set m_Doc = MNew.Document
'        UpdateFMainCaption Application.DefaultFileName
'    End If
    NewDocument pfn
'    PnlPerson_Enabled = False
'    TlbCountry_Enabled = False
'    TlbAddress_Enabled = False
'    TlbTelNr_Enabled = False
'    TlbCity_Enabled = False
    Me.WindowState = Settings.FMainWindowState
    
    BtnShowPersons_Click
    EGender_ToListBox CmbPersonGender
    PnlPersons.BorderStyle = 0
    PnlCCAT.BorderStyle = 0
    
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = PnlPersons.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight
    If W > 0 And H > 0 Then
        PnlPersons.Move L, T, W, H
        PnlCCAT.Move L, T, W, H
    End If
End Sub

Private Sub PnlPersons_Resize()
    PnlPersonList.Height = PnlPersons.ScaleHeight
    PnlPersonDetail.Height = PnlPersons.ScaleHeight
End Sub

Private Sub PnlPersonList_Resize()
    Dim brdr As Single: brdr = Screen.TwipsPerPixelX * 8
    Dim L  As Single, T  As Single, W  As Single, H  As Single
    Dim L1 As Single, T1 As Single, W1 As Single, H1 As Single
    L = brdr: T = brdr
    LblPersons.Move L, T
    T = T + LblPersons.Height
    W = PnlPersonList.ScaleWidth - L
    If W > 0 Then
        TlbPersons.Move L, T, W
        Dim btn As Button: Set btn = TlbPersons.Buttons.Item(11)
        L1 = btn.Left + btn.Width
        T1 = btn.Top
        W1 = TlbPersons.Width - L1
        H1 = btn.Height
        If W1 > 0 Then
            TxtSearchPerson.Move L1, T1, W1, H1
        End If
    End If
    'L = brdr
    T = T + TlbPersons.Height
    'Debug.Print TlbPersons.Height
    H = PnlPersonList.ScaleHeight - T - 3 * brdr 'wieso 3 * brdr?
    If W > 0 And H > 0 Then
        LstPerson.Move L, T, W, H
    End If
End Sub
Private Sub PnlPersonDetail_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single, T As Single, W As Single, H As Single
    L = brdr
    W = Max(W, LblName.Width)
    W = Max(W, LblName2.Width)
    W = Max(W, LblFamName.Width)
    W = Max(W, LblBirthD.Width)
    W = Max(W, LblGender.Width)
    W = Max(W, LblMother.Width)
    W = Max(W, LblFather.Width)
    W = Max(W, LblAddress.Width)
    W = Max(W, LblTelefonNr.Width)
    'W = 1365
    L = brdr + W + brdr
    H = 345
    W = PnlPersonDetail.Width - L - brdr
    T = LblName.Top:      If W > 0 And H > 0 Then TxtName.Move L, T, W ', H
    T = LblName2.Top:     If W > 0 And H > 0 Then TxtName2.Move L, T, W ', H
    T = LblFamName.Top:   If W > 0 And H > 0 Then TxtFamName.Move L, T, W ', H
    T = LblBirthD.Top:    If W > 0 And H > 0 Then TxtBirthD.Move L, T, W ', H
    T = LblGender.Top:    If W > 0 And H > 0 Then CmbPersonGender.Move L, T, W ', H
    T = LblMother.Top:    If W > 0 And H > 0 Then CmbMother.Move L, T, W ', H
    T = LblFather.Top:    If W > 0 And H > 0 Then CmbFather.Move L, T, W ', H
    T = LblAddress.Top:   If W > 0 And H > 0 Then CmbAddress.Move L, T, W ', H
    T = LblTelefonNr.Top: If W > 0 And H > 0 Then CmbTelefonNr.Move L, T, W ', H
    T = T + 480
    L = brdr
    W = PnlPersonDetail.ScaleWidth - L - brdr
    H = PnlPersonDetail.Height - T
    If W > 0 And H > 0 Then
        TSFamFrnds.Move L, T, W, H
        PnlTabChildren.BorderStyle = 0 ' Kein
        PnlTabFriends.BorderStyle = 0  ' Kein
        T = T + 480
        H = H - 480
        'Debug.Print L; T; W; H
        PnlTabChildren.Move L, T, W, H
        PnlTabFriends.Move L, T, W, H
        
    End If
    
End Sub
Private Sub PnlTabChildren_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single, T As Single, W As Single, H As Single
    L = 0 'brdr
    T = 0
    W = PnlTabChildren.Width
    H = PnlTabChildren.ScaleHeight - 3 * brdr
    If W > 0 And H > 0 Then
        LstChildren.Move L, T, W, H
    End If
End Sub
Private Sub PnlTabFriends_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L  As Single, T  As Single, W  As Single, H  As Single
    Dim L1 As Single, T1 As Single, W1 As Single, H1 As Single
    L = 0 'brdr
    T = 0
    W = PnlTabChildren.Width
    H = PnlTabChildren.ScaleHeight - 3 * brdr
    If W > 0 And H > 0 Then
        TlbPersonFriends.Move L, T, W ', H
        Dim btn As Button: Set btn = TlbPersonFriends.Buttons.Item(11)
        L1 = btn.Left + btn.Width
        T1 = btn.Top
        W1 = TlbPersonFriends.Width - L1
        H1 = btn.Height
        If W1 > 0 Then
            TxtSearchFriends.Move L1, T1, W1, H1
        End If
    End If
    T = T + TlbPersonFriends.Height
    H = H - T
    If W > 0 And H > 0 Then
        LstFriends.Move L, T, W, H
    End If
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'hier zuerst abfragen ob Datei schon gespeichert wurde
    'z.B. so: wenn UndoRedoManager leer dann keine Änderung wenn nicht, dann nohc nicht gespeichert?
    'OK oder vielleicht nicht, weil dann die UndoRedo immer gelöscht werden muss.
    
    Settings.FMainWindowState = Me.WindowState
    Application.Terminate
End Sub

' ############################## '         Menue         ' ############################## '
' ############################## '         File          ' ############################## '
Private Sub mnuFileNew_Click()
    NewDocument Nothing
    'Set m_Doc = New Document
    'UpdateView
End Sub

'OK hier brauchen wir eine Vereinheitlichung
'eine Dokument kann über 6 verschiedene Wege angelegt werden:
'* bei Programmstart NewDocument
'* im Laufenden Betrieb über mnuFileNew
'* über Datei-Doppelklick im Explorer
'* über den DateiÖffnen-Dialog
'* über die MRU-List
'* über Datei drag'n'dop
'was ist dann zu tun:
'* die MRU-List muss aktualisiert werden
'* die FormCaption muss aktualisiert werden
'* die Ansicht muss aktualisiert werden
'

Private Sub LstPerson_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    MyDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    MyDragDrop Data, Effect, Button, Shift, X, y
'    If Data.Files.Count = 0 Then Exit Sub
'    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
'    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(Data.Files(1))
'    If Not Application.IsValidFileExt(pfn) Then
'        MsgBox "Dieses Dateiformat wird momentan nicht unterstützt: " & vbCrLf & pfn.Extension & vbCrLf & pfn.Value
'        Exit Sub
'    End If
'    NewDocument pfn
'    Set m_Doc = MNew.Document(pfn)
'    Settings.MRUFiles_Add pfn
'    UpdateFMainCaption
'    UpdateView
End Sub
Private Sub MyDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Data.Files.Count = 0 Then Exit Sub
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(Data.Files(1))
    NewDocument pfn
End Sub

Private Sub mnuFileOpen_Click()
    Dim pfn As PathFileName 'String ': pfn = App.Path & "\" & Application.DefaultFileName
    If Application.OpenFileName_ShowDlg(pfn) = vbCancel Then Exit Sub
    'in Application.OpenFile wird die Datei schon in MRUFiles gesetzt. nein, jetzt hier in der Form
    NewDocument pfn
'    Set m_Doc = MNew.Document(pfn)
'    UpdateFMainCaption
'    UpdateView
End Sub

Private Sub mnuFileRecentFile_Click(Index As Integer)
    Dim pfn As PathFileName: Set pfn = Settings.MRUFiles.Item(Index)
    NewDocument pfn
'    Set m_Doc = MNew.Document(pfn)
'    Settings.MRUFiles_Add pfn
'    UpdateFMainCaption
'    UpdateView
End Sub

Private Sub NewDocument(pfn As PathFileName)
    'If pfn Is Nothing Then wird im Document erledigt
    'bei DragDrop und wenn über command eine Datei übergeben wird:
    If Not pfn Is Nothing Then
        If Not Application.IsValidFileExt(pfn) Then
            MsgBox "Dieses Dateiformat wird momentan nicht unterstützt: " & vbCrLf & pfn.Extension & vbCrLf & pfn.Value
            Exit Sub
        End If
        Settings.MRUFiles_Add pfn
    End If
    Set m_Doc = MNew.Document(pfn)
    UpdateFMainCaption
    UpdateView
End Sub
Private Sub mnuFileSave_Click()
    If m_Doc.Exists Then
        m_Doc.SaveFile
        UpdateFMainCaption
    Else
        mnuFileSaveAs_Click
    End If
End Sub
Private Sub mnuFileSaveAs_Click()
    'Dim sfnm As String: sfnm = IIf(m_Doc.Exists, m_Doc.Name, App.Path & "\" & Application.DefaultFileName)
    Dim pfn As PathFileName: Set pfn = m_Doc.PathFileName
    If Application.SaveFileName_ShowDlg(pfn) = vbCancel Then Exit Sub
    m_Doc.SaveFile pfn
    UpdateFMainCaption
    UpdateView
End Sub
Private Sub mnuFilePrint_Click()
    '
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Public Sub MRUFiles_FillMenu(mru As List)
'Man könnte auf die Idee kommen, das Menü gleich als Liste zu benützen
'dies ist jedoch keine gute Idee, falls der Dateiname zu lang ist, muss
'der Dateinamen gekürzt angezeigt werden.
'Besser ist eine extra MRU-List mit PathFileName-Objekten die den Namen
'gekürzt anzeigen können

    Dim n As Long: n = mru.Count
    If n <= 0 Then
        mnuFileRecentFiles.Visible = False
    Else
        mnuFileRecentFiles.Visible = True
        Dim pfn As PathFileName: Set pfn = mru.Item(0)
        mnuFileRecentFile(0).Caption = "&" & 1 & " " & pfn.Shorted
        If n > 1 Then
            Dim i As Long
            For i = 1 To n - 1
                If i >= mnuFileRecentFile.Count Then
                    Load mnuFileRecentFile(i)
                End If
                Set pfn = mru.Item(i)
                If Not pfn Is Nothing Then
                    mnuFileRecentFile(i).Caption = "&" & i + 1 & " " & pfn.Shorted
                End If
            Next
        End If
    End If
End Sub

'Sub MRUFiles_FillMenu(mru As Collection)
'    Dim n As Long: n = mru.Count
'    If n = 0 Then
'        mnuFileRecentFiles.Visible = False
'    Else
'        mnuFileRecentFiles.Visible = True
'        mnuFileRecentFile(0).Caption = mru(1)
'        If n > 1 Then
'            Dim i As Long
'            For i = 1 To n - 1
'                If i >= mnuFileRecentFile.Count Then
'                    Load mnuFileRecentFile(i)
'                End If
'                mnuFileRecentFile(i).Caption = mru(i + 1)
'            Next
'        End If
'    End If
'End Sub

' ############################## '         Menue         ' ############################## '
' ############################## '         Edit          ' ############################## '

Private Sub mnuEditUndo_Click()
    'zuerst gesamte datei serialisieren und in String-Liste packen
    'zwischen Undo-Redo wechseln indem man ListIndex setzt.
End Sub

Private Sub mnuEditRedo_Click()
    '
End Sub

Private Sub mnuEditCut_Click()
    If m_obj Is Nothing Then Exit Sub
    Serializer.Init
    m_obj.Serial
    'MsgBox Serializer.Str.ToStr
    Clipboard.SetText Serializer.LStr.ToStr, ClipBoardConstants.vbCFText
    m_Doc.Remove m_obj
    UpdateView
End Sub

Private Sub mnuEditCopy_Click()
'    If m_obj Is Nothing Then Exit Sub
'    Serializer.Init
'    m_obj.Serial
'    'MsgBox Serializer.Str.ToStr
'    Clipboard.SetText Serializer.LStr.ToStr ', ClipBoardConstants.vbCFText

    
    Dim s0 As String: s0 = Clipboard.GetText
    s0 = InputBox("Gib einen Text ein", "ClipBoard.SetText", s0)
    If Len(s0) = 0 Then Exit Sub
    
    
    Clipboard.SetText s0 '"Hello ClipBoard dude?"
    '???????????????????
    Dim s As String
    If Clipboard.GetFormat(ClipBoardConstants.vbCFText) Then
        s = Clipboard.GetText
    End If
    MsgBox s
End Sub


Private Sub mnuEditPaste_Click()
    'Halt nicht einfach nur auf GLeichheit überprüfen
    'wenn es zwei gleiche Objekte gibt dann müssen auch die Objekte
    'in den neu einzufügenden Objekten durch die vorhandenen Objekte ersetzt werden.
    'oder einfach zuerst alle einfügen und dann die Duplikate rausschmeissen bzw n der Tiefe für jedes Objekt ersetzen.
    '
    Dim sCont As String: sCont = Clipboard.GetText
    If Left(sCont, 1) <> "#" Then Exit Sub
    Dim aNewDoc As New Document
    aNewDoc.Parse sCont
    'OK wie gehen wir vor beim Pasten?
    m_Doc.Paste aNewDoc
    UpdateView
End Sub

Private Sub mnuEditEditor_Click()
    m_Doc.Serial
    Dim sCont As String: sCont = Serializer.LStr.ToStr
    If FrmEditor.ShowDialog(Me, sCont) = vbCancel Then Exit Sub
    'sonst neu laden
    Dim aNewDoc As Document
    Set aNewDoc = New Document
    aNewDoc.Parse sCont
    Set m_Doc = aNewDoc
    UpdateView
End Sub

' ############################## '        Extras         ' ############################## '
Private Sub mnuExtraRegisterFileIcon_Click()
    frmOptions.FormShow Me, 1
End Sub

Private Sub mnuExtraUnRegisterFileIcon_Click()
    'Application.UnRegisterExt
End Sub

Private Sub mnuExtraOptions_Click()
    frmOptions.FormShow Me, 2
End Sub


' ############################## '         Help          ' ############################## '
Private Sub mnuHelpShow_Click()
    'frmAbout.Show vbModeless, Me
    frmHelp.Show vbModeless, Me
End Sub
Private Sub mnuHelpInfo_Click()
    'frmAbout.Show vbModeless, Me
    frmAbout.Show vbModal, Me
End Sub

' ############################## '        Toolbar        ' ############################## '
' ############################## '        Buttons        ' ############################## '

Private Sub TlbMain_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    'Case 1: 'sep
    Case "New":   mnuFileNew_Click
    Case "Open":  mnuFileOpen_Click
    Case "Save":  mnuFileSave_Click
    'Case 5: 'sep
    Case "Print": mnuFilePrint_Click
    Case "Edit":  mnuEditEditor_Click
    'Case 7: 'sep
    Case "Cut":   mnuEditCut_Click
    Case "Copy":  mnuEditCopy_Click
    Case "Paste": mnuEditPaste_Click
    'Case 11: 'sep
    Case "Undo":  mnuEditUndo_Click
    Case "Redo":  mnuEditRedo_Click
    End Select
End Sub

Private Sub BtnShowCCAT_Click()
    PnlCCAT.Enabled = True
    PnlCCAT.Visible = True
    PnlPersons.Enabled = False
    PnlPersons.Visible = False
    
    PnlCCAT.ZOrder 0
End Sub

Private Sub BtnShowPersons_Click()
    PnlCCAT.Enabled = False
    PnlCCAT.Visible = False
    PnlPersons.Enabled = True
    PnlPersons.Visible = True
        
    PnlPersons.ZOrder 0
End Sub

'Private Sub BtnFileNew_Click()
'    mnuFileNew_Click
'End Sub
'Private Sub BtnFileOpen_Click()
'    mnuFileOpen_Click
'End Sub
'Private Sub BtnFileSave_Click()
'    mnuFileSave_Click
'End Sub
'
'Private Sub BtnEditFile_Click()
'    mnuEditEditor_Click
'End Sub
'
'Private Sub BtnCopy_Click()
'    mnuEditCopy_Click
'End Sub
'
'Private Sub BtnPaste_Click()
'    mnuEditPaste_Click
'End Sub


' ############################## '        EditArea       ' ############################## '
' ############################## '        Buttons        ' ############################## '

Private Sub TlbPersons_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    Case "AddNew":    BtnPersonAdd_Click
                      PnlPerson_Enabled = m_Doc.Persons.Count > 0
    Case "InsertNew": BtnPersonInsert_Click
    Case "EditSave":  LstPerson_DblClick
    Case "Delete":    BtnPersonDel_Click
    Case "MoveUp":    BtnPersonMoveUp_Click
    Case "MoveDown":  BtnPersonMoveDown_Click
    Case "SortUp":    BtnPersonSortUp_Click
    Case "SortDown":  BtnPersonSortDown_Click
    End Select
End Sub
Private Sub BtnPersonAdd_Click()
    Dim obj As New Person
    If FrmPerson.ShowDialog(Me, obj, m_Doc.Persons, m_Doc.Addresses, m_Doc.TelefonNrs) = vbCancel Then Exit Sub
    m_Doc.Add obj
    Dim u As Long: u = m_Doc.Persons.Count - 1
    m_Doc.Persons.ListIndex = u
    'UpdateView_Person obj
    m_Doc.Persons.ToListbox LstPerson ', True
    UpdateView_PersonTBs obj
End Sub
Private Sub BtnPersonInsert_Click()
    Dim i As Long: i = LstPerson.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As New Person
    If FrmPerson.ShowDialog(Me, obj, m_Doc.Persons, m_Doc.Addresses, m_Doc.TelefonNrs) = vbCancel Then Exit Sub
    m_Doc.Persons.Insert i, obj
    m_Doc.Persons.ToListbox LstPerson
    LstPerson.ListIndex = i
End Sub
'Private Sub BtnPersonEdit_Click()
''ist das gleiche wie LstPerson_Dblclick
'    Dim li As Long: li = LstPerson.ListIndex
'    If li < 0 Then Exit Sub
'    Dim obj As Person: Set obj = m_Doc.Persons.Item(li)
'    If obj Is Nothing Then Exit Sub
'    With obj
'        .PreName1 = Me.TxtName.Text
'        .PreName2 = Me.TxtName2.Text
'        .FamName = Me.TxtFamName.Text
'        .BirthD = ParseParam(.BirthD, Me.TxtBirthD.Text, obj, "BirthD")
'        .Gender = EGender_Parse(CmbPersonGender.Text)
'        Dim i As Long
'        i = CmbMother.ListIndex
'        If i >= 0 Then Set .Mother = m_Doc.Persons.Item(i) Else If CmbMother.Text = "" Then Set .Mother = Nothing
'        i = CmbFather.ListIndex
'        If i >= 0 Then Set .Father = m_Doc.Persons.Item(i) Else If CmbFather.Text = "" Then Set .Father = Nothing
'        i = CmbAddress.ListIndex
'        If i >= 0 Then Set .Address = m_Doc.Addresses.Item(i) Else If CmbAddress.Text = "" Then Set .Address = Nothing
'        i = CmbTelefonNr.ListIndex
'        If i >= 0 Then Set .TelNumber = m_Doc.TelefonNrs.Item(i) Else If CmbTelefonNr.Text = "" Then Set .TelNumber = Nothing
'    End With
'    'UpdateView_Person obj
'End Sub
Private Sub BtnPersonDel_Click()
    Dim i As Long: i = LstPerson.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As Person: Set obj = m_Doc.Persons.Item(i)
    Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
    If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
    m_Doc.Persons.Remove i
    UpdateView
End Sub
Private Sub BtnPersonMoveUp_Click()
    Dim i As Long: i = LstPerson.ListIndex
    If i <= 0 Then Exit Sub
    m_Doc.Persons.Swap i, i - 1
    UpdateView
    LstPerson.ListIndex = i - 1
End Sub
Private Sub BtnPersonMoveDown_Click()
    Dim i As Long: i = LstPerson.ListIndex
    If LstPerson.ListCount - 1 <= i Then Exit Sub
    m_Doc.Persons.Swap i, i + 1
    UpdateView
    LstPerson.ListIndex = i + 1
End Sub
Private Sub BtnPersonSortUp_Click()
    If m_Doc.Persons.EEmpty Then Exit Sub
    m_Doc.Persons.Sort
    UpdateView_PersonLV
End Sub
Private Sub BtnPersonSortDown_Click()
    If m_Doc.Persons.EEmpty Then Exit Sub
    m_Doc.Persons.SortRev
    UpdateView
End Sub

Private Sub TlbCountry_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    Case "AddNew":    BtnCountryAdd_Click
                      TlbCountry_Enabled = m_Doc.Countries.Count > 0
    Case "InsertNew": BtnCountryInsert_Click
    Case "EditSave":  BtnCountryEdit_Click
    Case "Delete":    BtnCountryDel_Click
    Case "MoveUp":    BtnCountryMoveUp_Click
    Case "MoveDown":  BtnCountryMoveDown_Click
    Case "SortUp":    BtnCountrySortUp_Click
    Case "SortDown":  BtnCountrySortDown_Click
    Case "Search":    BtnCountrySearch_Click
    End Select
End Sub
Private Sub BtnCountryAdd_Click()
    Dim obj As New Country
    If FrmCountry.ShowDialog(Me, obj) = vbCancel Then Exit Sub
    m_Doc.Add obj
    m_Doc.Countries.ToListbox LstCountry
End Sub
Private Sub BtnCountryInsert_Click()
    Dim i As Long: i = LstCountry.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As New Country
    If FrmCountry.ShowDialog(Me, obj) = vbCancel Then Exit Sub
    m_Doc.Countries.Insert i, obj
    m_Doc.Countries.ToListbox LstPerson
    LstCountry.ListIndex = i
End Sub
Private Sub BtnCountryEdit_Click()
    LstCountry_DblClick
End Sub
Private Sub BtnCountryDel_Click()
    Dim i As Long: i = LstCountry.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As Country: Set obj = m_Doc.Countries.Item(i)
    Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
    If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
    m_Doc.Countries.Remove i
    UpdateView
End Sub
Private Sub BtnCountryMoveUp_Click()
    Dim i As Long: i = LstCountry.ListIndex
    If i <= 0 Then Exit Sub
    m_Doc.Countries.Swap i, i - 1
    UpdateView
    LstCountry.ListIndex = i - 1
End Sub
Private Sub BtnCountryMoveDown_Click()
    Dim i As Long: i = LstCountry.ListIndex
    If LstCountry.ListCount - 1 <= i Then Exit Sub
    m_Doc.Countries.Swap i, i + 1
    UpdateView
    LstCountry.ListIndex = i + 1
End Sub
Private Sub BtnCountrySortUp_Click()
    If m_Doc.Countries.EEmpty Then Exit Sub
    m_Doc.Countries.Sort
    UpdateView
End Sub
Private Sub BtnCountrySortDown_Click()
    If m_Doc.Countries.EEmpty Then Exit Sub
    m_Doc.Countries.SortRev
    UpdateView
End Sub
Private Sub BtnCountrySearch_Click()
    '
End Sub

Private Sub TlbCity_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    Case "AddNew":   BtnCityAdd_Click
                     TlbCity_Enabled = m_Doc.Cities.Count > 0
    Case "EditSave": BtnCityEdit_Click
    Case "Delete":   BtnCityDel_Click
    Case "MoveUp":   BtnCityMoveUp_Click
    Case "MoveDown": BtnCityMoveDown_Click
    Case "SortUp":   BtnCitySortUp_Click
    Case "SortDown": BtnCitySortDown_Click
    Case "Search":   BtnCitySearch_Click
    End Select
End Sub
Private Sub BtnCityAdd_Click()
    Dim obj As New City
    If FrmCity.ShowDialog(Me, obj, m_Doc.Countries) = vbCancel Then Exit Sub
    m_Doc.Add obj
    m_Doc.Cities.ToListbox LstCity
End Sub
Private Sub BtnCityEdit_Click()
    Dim i As Long: i = LstCity.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As City: Set obj = m_Doc.Cities.Item(i)
End Sub
Private Sub BtnCityDel_Click()
    Dim i As Long: i = LstCity.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As City: Set obj = m_Doc.Cities.Item(i)
    Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
    If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
    m_Doc.Cities.Remove i
    UpdateView
End Sub
Private Sub BtnCityMoveUp_Click()
    Dim i As Long: i = LstCity.ListIndex
    If i <= 0 Then Exit Sub
    m_Doc.Cities.Swap i, i - 1
    UpdateView
    LstCity.ListIndex = i - 1
End Sub
Private Sub BtnCityMoveDown_Click()
    Dim i As Long: i = LstCity.ListIndex
    If LstCity.ListCount - 1 <= i Then Exit Sub
    m_Doc.Cities.Swap i, i + 1
    UpdateView
    LstCity.ListIndex = i + 1
End Sub
Private Sub BtnCitySortUp_Click()
    If m_Doc.Cities.EEmpty Then Exit Sub
    m_Doc.Cities.Sort
    UpdateView
End Sub
Private Sub BtnCitySortDown_Click()
    If m_Doc.Cities.EEmpty Then Exit Sub
    m_Doc.Cities.SortRev
    UpdateView
End Sub
Private Sub BtnCitySearch_Click()
    '
End Sub

Private Sub TlbTelNr_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    Case "AddNew":   BtnTelNrAdd_Click
                     TlbTelNr_Enabled = m_Doc.TelefonNrs.Count > 0
    Case "Insert":   BtnTelNrInsert_Click
    Case "EditSave": BtnTelNrEdit_Click
    Case "Delete":   BtnTelNrDel_Click
    Case "MoveUp":   BtnTelNrMoveUp_Click
    Case "MoveDown": BtnTelNrMoveDown_Click
    Case "SortUp":   BtnTelNrSortUp_Click
    Case "SortDown": BtnTelNrSortDown_Click
    Case "Search":   BtnTelNrSearch_Click
    End Select
End Sub
Private Sub BtnTelNrAdd_Click()
    Dim obj As New TelefonNr
    If FrmTelefonNr.ShowDialog(Me, obj, m_Doc.Countries) = vbCancel Then Exit Sub
    m_Doc.Add obj
    m_Doc.TelefonNrs.ToListbox LstTelefonNr
End Sub
Private Sub BtnTelNrInsert_Click()
    '
End Sub
Private Sub BtnTelNrEdit_Click()
    '
End Sub
Private Sub BtnTelNrDel_Click()
    '
End Sub
Private Sub BtnTelNrMoveUp_Click()
    '
End Sub
Private Sub BtnTelNrMoveDown_Click()
    '
End Sub
Private Sub BtnTelNrSortUp_Click()
    If m_Doc.TelefonNrs.EEmpty Then Exit Sub
    m_Doc.TelefonNrs.SortRev
    UpdateView
End Sub
Private Sub BtnTelNrSortDown_Click()
    If m_Doc.TelefonNrs.EEmpty Then Exit Sub
    m_Doc.TelefonNrs.SortRev
    UpdateView
End Sub
Private Sub BtnTelNrSearch_Click()
    '
End Sub

Private Sub TlbAddress_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
    Case "AddNew":   BtnAddressAdd_Click
                     TlbAddress_Enabled = m_Doc.Addresses.Count > 0
    Case "Insert":   BtnAddressInsert_Click
    Case "EditSave": BtnAddressEdit_Click
    Case "Delete":   BtnAddressDel_Click
    Case "MoveUp":   BtnAddressMoveUp_Click
    Case "MoveDown": BtnAddressMoveDown_Click
    Case "SortUp":   BtnAddressSortUp_Click
    Case "SortDown": BtnAddressSortDown_Click
    Case "Search":   BtnAddressSearch_Click
    End Select
End Sub

Private Sub BtnAddressAdd_Click()
    Dim obj As New Address
    If FrmAddress.ShowDialog(Me, obj, m_Doc.Cities) = vbCancel Then Exit Sub
    m_Doc.Add obj
    m_Doc.Addresses.ToListbox LstAddress
End Sub
Private Sub BtnAddressInsert_Click()
    '
End Sub
Private Sub BtnAddressEdit_Click()
    '
End Sub
Private Sub BtnAddressDel_Click()
    '
End Sub
Private Sub BtnAddressMoveUp_Click()
    '
End Sub
Private Sub BtnAddressMoveDown_Click()
    '
End Sub
Private Sub BtnAddressSortUp_Click()
    '
End Sub
Private Sub BtnAddressSortDown_Click()
    '
End Sub
Private Sub BtnAddressSearch_Click()
    '
End Sub

'Private Sub BtnAddTest_Click()
'    Dim obj As New Test
'    If FrmTest.ShowDialog(me, obj) = vbCancel Then Exit Sub
'    m_Doc.Add obj
'    m_Doc.Tests.ToListbox LstTest
'End Sub

' ############################## '        ListBoxen        ' ############################## '

Private Function SelectedPerson() As Person
    Dim i As Long: i = LstPerson.ListIndex
    m_Doc.Persons.ListIndex = i
    Set m_obj = m_Doc.Persons.Item(i)
    Set SelectedPerson = m_obj
End Function
Private Sub LstPerson_Click()
    If LstPerson.ListIndex < 0 Then Exit Sub
    UpdateView_PersonTBs SelectedPerson
End Sub
Private Sub LstPerson_DblClick()
    Dim i As Long: i = LstPerson.ListIndex
    If i < 0 Then Exit Sub
    Dim obj As Person: Set obj = SelectedPerson
    If FrmPerson.ShowDialog(Me, obj, m_Doc.Persons, m_Doc.Addresses, m_Doc.TelefonNrs) = vbCancel Then Exit Sub
    'nur update dieses einzigen listitems
    LstPerson.List(i) = obj.Key
    'dann noch alle textboxen updaten
    UpdateView_PersonTBs obj
End Sub
Private Sub LstPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If LstPerson.ListIndex < 0 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        BtnPersonDel_Click
    End If
End Sub

Private Function SelectedCountry() As Country
    Dim i As Long: i = LstCountry.ListIndex
    If i < 0 Then Exit Function
    m_Doc.Countries.ListIndex = i
    Set m_obj = m_Doc.Countries.Item(i)
    Set SelectedCountry = m_obj
End Function
Private Sub LstCountry_Click()
    m_Doc.Countries.ListIndex = LstCountry.ListIndex
End Sub
Private Sub LstCountry_DblClick()
    Dim obj As Country: Set obj = SelectedCountry
    If obj Is Nothing Then Exit Sub
    If FrmCountry.ShowDialog(Me, obj) = vbCancel Then Exit Sub
    LstCountry.List(LstCountry.ListIndex) = obj.Key
End Sub
Private Sub LstCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Dim i As Long: i = LstCountry.ListIndex
        Dim obj As Country: Set obj = m_Doc.Countries.Item(i)
        Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
        If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
        m_Doc.Countries.Remove i
        LstCountry.RemoveItem i
    End If
End Sub

Private Function SelectedCity() As City
    Dim i As Long: i = LstCity.ListIndex
    m_Doc.Cities.ListIndex = i
    Set m_obj = m_Doc.Cities.Item(i)
    Set SelectedCity = m_obj
End Function
Private Sub LstCity_Click()
    m_Doc.Cities.ListIndex = LstCity.ListIndex
End Sub
Private Sub LstCity_DblClick()
    Dim obj As City: Set obj = SelectedCity
    If FrmCity.ShowDialog(Me, obj, m_Doc.Countries) = vbCancel Then Exit Sub
    LstCity.List(LstCity.ListIndex) = obj.Key
End Sub
Private Sub LstCity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Dim i As Long: i = LstCity.ListIndex
        Dim obj As City: Set obj = m_Doc.Cities.Item(i)
        Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
        If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
        m_Doc.Cities.Remove i
        LstCity.RemoveItem i
    End If
    UpdateView_CCAT
End Sub

Private Function SelectedAddress() As Address
    Dim i As Long: i = LstAddress.ListIndex
    m_Doc.Addresses.ListIndex = i
    Set m_obj = m_Doc.Addresses.Item(i)
    Set SelectedAddress = m_obj
End Function
Private Sub LstAddress_Click()
    m_Doc.Addresses.ListIndex = LstAddress.ListIndex
End Sub
Private Sub LstAddress_DblClick()
    Dim obj As Address: Set obj = SelectedAddress
    If FrmAddress.ShowDialog(Me, obj, m_Doc.Cities) = vbCancel Then Exit Sub
    LstAddress.List(LstAddress.ListIndex) = obj.Key
End Sub
Private Sub LstAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Dim i As Long: i = LstAddress.ListIndex
        Dim obj As Address: Set obj = m_Doc.Addresses.Item(i)
        Dim msg As String: msg = "Soll das Object wirklich gelöscht werden?" & vbCrLf & obj.Key
        If MsgBox(msg, vbOKCancel) = vbCancel Then Exit Sub
        m_Doc.Addresses.Remove i
        LstAddress.RemoveItem i
    End If
End Sub

Private Function SelectedTelefonNr() As TelefonNr
    Dim i As Long: i = LstTelefonNr.ListIndex
    m_Doc.TelefonNrs.ListIndex = i
    Set m_obj = m_Doc.TelefonNrs.Item(i)
    Set SelectedTelefonNr = m_obj
End Function
Private Sub LstTelefonNr_Click()
    m_Doc.TelefonNrs.ListIndex = LstTelefonNr.ListIndex
End Sub
Private Sub LstTelefonNr_DblClick()
    Dim obj As TelefonNr: Set obj = SelectedTelefonNr
    If FrmTelefonNr.ShowDialog(Me, obj, m_Doc.Countries) = vbCancel Then Exit Sub
    LstTelefonNr.List(LstTelefonNr.ListIndex) = obj.Key
End Sub

Property Let PnlPerson_Enabled(ByVal Value As Boolean)
    'die Textboxen und die Toolbar-Buttons disablen bis auf [ + ]
    
    TlbButtons_Enabled(TlbPersons) = Value
    Dim cEnabled As Long: cEnabled = SystemColorConstants.vbWindowBackground
    'Dim cDisabld As Long: cDisabld = SystemColorConstants.vbGrayText
    Dim cDisabld As Long: cDisabld = RGB(240, 240, 240)
    Dim g As Long: g = IIf(Value, cEnabled, cDisabld)
    Me.TxtName.Enabled = Value:         Me.TxtName.BackColor = g
    Me.TxtName2.Enabled = Value:        Me.TxtName2.BackColor = g
    Me.TxtFamName.Enabled = Value:      Me.TxtFamName.BackColor = g
    Me.TxtBirthD.Enabled = Value:       Me.TxtBirthD.BackColor = g
    Me.CmbPersonGender.Enabled = Value: Me.CmbPersonGender.BackColor = g
    Me.CmbMother.Enabled = Value:       Me.CmbMother.BackColor = g
    Me.CmbFather.Enabled = Value:       Me.CmbFather.BackColor = g
    Me.CmbTelefonNr.Enabled = Value:    Me.CmbTelefonNr.BackColor = g
    Me.CmbAddress.Enabled = Value:      Me.CmbAddress.BackColor = g
    
    TlbButtons_Enabled(TlbPersonFriends) = Value
    
    Me.TSFamFrnds.Enabled = Value
End Property
Property Let TlbPerson_Enabled(ByVal Value As Boolean)
    TlbButtons_Enabled(TlbPersons) = Value
End Property
Property Let TlbCountry_Enabled(ByVal Value As Boolean)
    TlbButtons_Enabled(TlbCountry) = Value
End Property
Property Let TlbCity_Enabled(ByVal Value As Boolean)
    TlbButtons_Enabled(TlbCity) = Value
End Property
Property Let TlbAddress_Enabled(ByVal Value As Boolean)
    TlbButtons_Enabled(TlbAddress) = Value
End Property
Property Let TlbTelNr_Enabled(ByVal Value As Boolean)
    TlbButtons_Enabled(TlbTelNr) = Value
End Property
Property Let TlbButtons_Enabled(aTlb As Toolbar, ByVal Value As Boolean)
    Dim i As Long
    With aTlb
        .Buttons(1).Enabled = True
        For i = 2 To .Buttons.Count '- 1
            .Buttons(i).Enabled = Value
        Next
    End With
End Property


'Private Function SelectedTest() As Test
'    Dim i As Long: i = LstTest.ListIndex
'    m_Doc.Tests.ListIndex = i
'    Set m_Obj = m_Doc.Tests.Item(i)
'    Set SelectedTest = m_Obj
'End Function
'Private Sub LstTest_Click()
'    m_Doc.Tests.ListIndex = LstTest.ListIndex
'End Sub
'Private Sub LstTest_DblClick()
'    Dim obj As Test: Set obj = SelectedTest
'    If FrmTest.ShowDialog(me, obj) = vbCancel Then Exit Sub
'    LstTest.List(LstTest.ListIndex) = obj.Key
'End Sub
Public Sub UpdateFMainCaption(Optional ByVal fnam As String = "")
    If Len(fnam) = 0 Then fnam = m_Doc.PathFileName.Value
    Me.Caption = App.ProductName & " - [" & fnam & "]"
End Sub

Sub UpdateView()
    TlbPerson_Enabled = m_Doc.Persons.Count > 0
    PnlPerson_Enabled = m_Doc.Persons.Count > 0
    'tlbfriends_enable = m_doc.Persons.ListIndex
    TlbCountry_Enabled = m_Doc.Countries.Count > 0
    TlbCity_Enabled = m_Doc.Cities.Count > 0
    TlbAddress_Enabled = m_Doc.Addresses.Count > 0
    TlbTelNr_Enabled = m_Doc.TelefonNrs.Count > 0
    
    UpdateView_Person m_obj
    UpdateView_CCAT
    
End Sub

Sub UpdateView_Person(obj As Person)
    UpdateView_PersonTBs obj
    UpdateView_PersonLV
End Sub
Sub UpdateView_PersonTBs(obj As Person)
    If obj Is Nothing Then Exit Sub
    With obj
        Me.TxtName.Text = .PreName1
        Me.TxtName2.Text = .PreName2
        Me.TxtFamName.Text = .FamName
        Me.TxtBirthD.Text = IIf(.BirthD = 0, "", Format(.BirthD, "dd.mmm.yyyy"))
        Me.CmbPersonGender.Text = EGender_ToStr(.Gender)
        
        Dim p As Person
        Set p = .Mother
        If Not p Is Nothing Then Me.CmbMother.Text = p.Key Else Me.CmbMother.Text = ""
        Set p = .Father
        If Not p Is Nothing Then Me.CmbFather.Text = p.Key Else Me.CmbFather.Text = ""
        
        Dim a As Address: Set a = .Address
        If Not a Is Nothing Then Me.CmbAddress.Text = a.Key Else Me.CmbAddress.Text = ""
        
        Dim T As TelefonNr: Set T = .TelNumber
        If Not T Is Nothing Then Me.CmbTelefonNr.Text = T.Key Else Me.CmbTelefonNr.Text = ""
    End With
End Sub
Sub UpdateView_PersonLV()
    Dim i As Long: i = Me.LstPerson.ListIndex
    m_Doc.Persons.ToListbox Me.LstPerson, True ', i >= 0
    'If i >= 0 Then Me.LstPerson.ListIndex = i
End Sub

Sub UpdateView_CCAT()
    Dim i As Long
    With m_Doc
        
        i = Me.LstAddress.ListIndex
        .Addresses.ToListbox Me.LstAddress
        If i >= 0 And Me.LstAddress.ListCount > 0 Then Me.LstAddress.ListIndex = i
        
        i = Me.LstCity.ListIndex
        .Cities.ToListbox Me.LstCity
        If i >= 0 And Me.LstCity.ListCount > 0 Then Me.LstCity.ListIndex = i
        
        i = Me.LstCountry.ListIndex
        .Countries.ToListbox Me.LstCountry
        If i >= 0 And Me.LstCountry.ListCount > 0 Then Me.LstCountry.ListIndex = i
        
        i = Me.LstTelefonNr.ListIndex
        .TelefonNrs.ToListbox Me.LstTelefonNr
        If i >= 0 And Me.LstTelefonNr.ListCount > 0 Then Me.LstTelefonNr.ListIndex = i
        
    End With
End Sub

Private Sub TSFamFrnds_Click()
    Select Case TSFamFrnds.SelectedItem.Index
    Case 1: PnlTabChildren.ZOrder 0
    Case 2: PnlTabFriends.ZOrder 0
    End Select
End Sub

Private Sub TxtName_LostFocus()
    If m_obj Is Nothing Then Exit Sub
    Dim prs As Person: Set prs = m_obj: prs.PreName1 = TxtName.Text
End Sub
Private Sub TxtName2_LostFocus()
    If m_obj Is Nothing Then Exit Sub
    Dim prs As Person: Set prs = m_obj: prs.PreName2 = TxtName2.Text
End Sub
Private Sub TxtFamName_LostFocus()
    If m_obj Is Nothing Then Exit Sub
    Dim prs As Person: Set prs = m_obj: prs.FamName = TxtFamName.Text
End Sub
Private Sub TxtBirthD_LostFocus()
    If m_obj Is Nothing Then Exit Sub
    Dim s As String: s = TxtBirthD.Text: If Len(s) = 0 Then Exit Sub
    Dim prs As Person: Set prs = m_obj
    Dim dat As Date: If Date_TryParse(s, dat) Then prs.BirthD = dat
End Sub
Function Date_TryParse(s As String, ByRef out_date As Date) As Boolean
Try: On Error GoTo Catch
    out_date = CDate(s)
    Date_TryParse = True
Catch:
End Function
Private Sub CmbPersonGender_LostFocus()
    If m_obj Is Nothing Then Exit Sub
    Dim prs As Person: Set prs = m_obj: prs.Gender = EGender_Parse(Me.CmbPersonGender.Text)
End Sub
'Private Sub CmbPersonGender_Click()
'    Dim prs As Person: Set prs = m_obj: prs.Gender = EGender_Parse(Me.CmbPersonGender.Text)
'End Sub
Private Sub CmbMother_Click()
    If m_obj Is Nothing Then Exit Sub
    Dim i As Long: i = CmbMother.ListIndex:
    Dim mot As Person: Set mot = m_Doc.Persons.Item(i)
    Dim prs As Person: Set prs = m_obj: Set prs.Mother = mot
End Sub
Private Sub CmbFather_Click()
    If m_obj Is Nothing Then Exit Sub
    Dim i As Long: i = CmbFather.ListIndex:
    Dim fat As Person: Set fat = m_Doc.Persons.Item(i)
    Dim prs As Person: Set prs = m_obj: Set prs.Father = fat
End Sub




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//2005_09_01
'//390 Zeilen
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As mRect, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const DFC_BUTTON As Long = 4: Private Const DFCS_BUTTONPUSH As Long = &H10: Private Const DFCS_PUSHED As Long = &H200
Private mName As String
Private WithEvents mMDIPar As MDIForm
Attribute mMDIPar.VB_VarHelpID = -1
Private WithEvents mParent As Form
Attribute mParent.VB_VarHelpID = -1
Private WithEvents mContainer As PictureBox
Attribute mContainer.VB_VarHelpID = -1
Private WithEvents mPicSplit As PictureBox
Attribute mPicSplit.VB_VarHelpID = -1
Private WithEvents mXPStyleC As CommandButton 'PictureBox 'Jo, CommandButton sieht richtig schick aus!
Attribute mXPStyleC.VB_VarHelpID = -1
Private mWidth As Integer  'die Breite des Splitters in Pixel
Private mHorVer As Boolean 'True = horizontaler Splitter False = Vertikaler Splitter
Private m12LTRB As Long '(01)1: LeftCtrl RightSplit; (10)2: LeftSplit RightCtrl; (11)3: LeftCtrl MidSplit RightCtrl
Private mLeftTopCtrl As Control
Private mRghtBotCtrl As Control
Private OldX As Long 'X bei MouseDown
Private OldY As Long 'Y bei MouseDown
Private CurX As Long 'X bei MouseMove
Private CurY As Long 'Y bei MouseMove
Private DX  As Long  'Abstand zw. OldX und CurX
Private DY  As Long  'Abstand zw. OldY und CurY
Private mMinLT As Long 'Minimale Breite/Höhe des Left-Top-Controls
Private mMinRB As Long 'Minimale Breite/Höhe des Right-Bot-Controls
Private mLeft As Long  'merkt sich die letzte Left-Pos des Splitters (Umschalten zw hor und vert)
Private mTop As Long   'merkt sich die letzte Top-Pos des Splitters (Umschalten zw hor und vert)
Private mLeftTopPos As Long 'entweder Left oder top
Private mBolMove As Boolean 'die Maus wird gezogen
Private STPX As Long
Private STPY As Long
Private Type mRect
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Public Enum ESplitBorderstyle
    bsNone = 0
    bsSunken = 1
    bsRaised = 2
    bsXPStyl = 3
End Enum
Private mBorderStyle As ESplitBorderstyle
Private mBolMDI As Boolean
Public Event OnMove(Sender As Splitter)
Public Event OnDoubleClick(Sender As Splitter)
Private mInitializing As Boolean

Private Sub Class_Initialize()
    '
    STPX = Screen.TwipsPerPixelX 'is kürzer
    STPY = Screen.TwipsPerPixelY 'is kürzer
    mWidth = 5
    mMinLT = 10 * STPX 'Minimale Breite/Höhe des Left-Top-Controls (10Pixel)
    mMinRB = 10 * STPX 'Minimale Breite/Höhe des Right-Bot-Controls(10Pixel)
    mBorderStyle = bsRaised
End Sub

Friend Sub New_(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control)
    mInitializing = True
    mName = Name
    mBolMDI = BolMDI
    Set mContainer = MyContainer
    If mBolMDI Then
        Set mMDIPar = MyOwner
        If mPicSplit Is Nothing Then Set mPicSplit = mMDIPar.Controls.Add("VB.PictureBox", mName, mContainer)
    Else
        Set mParent = MyOwner
        If mPicSplit Is Nothing Then Set mPicSplit = mParent.Controls.Add("VB.PictureBox", mName, mContainer)
    End If
    
    'man könnte auch einen CommandButton verwenden, Vorteil bei XPStyles
    'MContainer.BorderStyle = 0
    'mContainer.BackColor = &HE0E0E0
    mPicSplit.Refresh
    mPicSplit.Visible = True
    mPicSplit.BorderStyle = 0 ' kein Border
    'mPicSplit.Appearance = 0  '-Flat2D
    mPicSplit.BackColor = mContainer.BackColor
    
    m12LTRB = 0
    If Not LeftTop Is Nothing Then
        Set mLeftTopCtrl = LeftTop
        mLeft = mLeftTopCtrl.Width
        mTop = mLeftTopCtrl.Height
        m12LTRB = m12LTRB + 1
    End If
    If Not RghtBot Is Nothing Then
        Set mRghtBotCtrl = RghtBot
        mLeft = mRghtBotCtrl.Width
        mTop = mRghtBotCtrl.Height
        m12LTRB = m12LTRB + 2
    End If
    BolHorizontal = True
    
    'BolHorizontal = True d.H. der Splitter hat eine senkrechte Form,
    'und trennt zwei benachbarte controls links-rechts also horizontal voneinander ab.
    'die controls werden horizontal resized
    mInitializing = False
End Sub

Public Property Get Width() As Integer
    Width = mWidth * STPX
End Property
Public Property Let Width(iVal As Integer)
    mWidth = iVal
End Property

Public Property Get LeftTopPos() As Long
    LeftTopPos = mLeftTopPos
End Property
Public Property Let LeftTopPos(LngValPos As Long)
    mLeftTopPos = LngValPos
    Call ResizeControls
End Property

Public Property Get RghtBotPos() As Long
    RghtBotPos = mLeftTopPos + mWidth
End Property
Public Property Let RghtBotPos(LngValPos As Long)
    mLeftTopPos = LngValPos - mWidth
    Call ResizeControls
End Property

Public Property Get BolHorizontal() As Boolean
    BolHorizontal = mHorVer 'True = Hori. Spl.; False = Vert. Spl.
End Property
Public Property Let BolHorizontal(BolVal As Boolean)
    If BolVal <> mHorVer Then
        mHorVer = BolVal
        If mHorVer Then
            mTop = mLeftTopPos
            mLeftTopPos = mLeft
            mPicSplit.MousePointer = 9 'Größenänderung W O
        Else
            mLeft = mLeftTopPos
            mLeftTopPos = mTop
            mPicSplit.MousePointer = 7 'Größenänderung N S
        End If
        Call ResizeControls
    End If
End Property

Public Property Get BorderStyle() As ESplitBorderstyle
    BorderStyle = mBorderStyle
End Property
Public Property Let BorderStyle(LngVal As ESplitBorderstyle)
    mBorderStyle = LngVal
    If mBorderStyle = bsXPStyl Then
        If mXPStyleC Is Nothing Then
            If mBolMDI Then
                Set mXPStyleC = mMDIPar.Controls.Add("VB.CommandButton", mName & "XPStyleC", mPicSplit)
            Else
                Set mXPStyleC = mParent.Controls.Add("VB.CommandButton", mName & "XPStyleC", mPicSplit)
            End If
            mXPStyleC.Visible = True
        End If
    Else
        If Not mXPStyleC Is Nothing Then
            'Set mXPStyleC = Nothing
            mXPStyleC.Visible = False
        End If
    End If
End Property

Public Property Get LeftTopCtrl() As Control
    Set LeftTopCtrl = mLeftTopCtrl
End Property
Public Property Get RghtBotCtrl() As Control
    Set RghtBotCtrl = mRghtBotCtrl
End Property
Public Sub Refresh()
    If Not mPicSplit Is Nothing Then mPicSplit.Refresh
    If Not mLeftTopCtrl Is Nothing Then mLeftTopCtrl.Refresh
    If Not mRghtBotCtrl Is Nothing Then mRghtBotCtrl.Refresh
    ResizeControls
End Sub

Private Sub mXPStyleC_DblClick()
    RaiseEvent OnDoubleClick(Me)
End Sub


Private Sub mContainer_Resize()
    'MDIForm1.Text1.Text = "mCResize: L=" & CStr(mContainer.Left) & " T=" & CStr(mContainer.Top) & " W=" & CStr(mContainer.Width) & " H=" & CStr(mContainer.Height)
    'If Not mBolResizeSelf Then
    If Not mBolMove Then
        Call ResizeControls
    End If
End Sub

Private Sub mPicSplit_DblClick()
    RaiseEvent OnDoubleClick(Me)
End Sub

Private Sub mXPStyleC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mPicSplit_KeyDown(KeyCode, 0)
End Sub
Private Sub mPicSplit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 33:         DX = -1 * STPX
    Case 34:         DX = 1 * STPX
    Case vbKeyLeft:  DX = -1 * STPX
    Case vbKeyUp:    DX = -1 * STPX
    Case vbKeyRight: DX = 1 * STPX
    Case vbKeyDown:  DX = 1 * STPX
    End Select
    DY = DX
    ResizeControls
End Sub

Private Sub mXPStyleC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call mPicSplit_MouseDown(Button, Shift, X, y)
End Sub
Private Sub mPicSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        mBolMove = True
        OldX = X
        OldY = y
    End If
End Sub
Private Sub mXPStyleC_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call mPicSplit_MouseMove(Button, Shift, X, y)
End Sub
Private Sub mPicSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If mBolMove Then
        DX = X - OldX
        DY = y - OldY
        If mHorVer Then
            If DX <> 0 Then
                ResizeControls
            End If
        Else
            If DY <> 0 Then
                ResizeControls
            End If
        End If
        RaiseEvent OnMove(Me)
    End If
    'Form1.Label1.Caption = mLeftTopCtrl.Name
End Sub
Private Sub mXPStyleC_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call mPicSplit_MouseUp(Button, Shift, X, y)
End Sub
Private Sub mPicSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        DX = X - OldX
        DY = y - OldY
        Call ResizeControls
        If mBolMove Then mBolMove = False
    End If
End Sub

Public Sub ResizeControls()
    If mInitializing Then Exit Sub
    Dim Lm As Long, Tm As Long, Wm As Long, Hm As Long
    Dim L1 As Long, T1 As Long, W1 As Long, H1 As Long 'LeftTop-Control
    Dim L2 As Long, T2 As Long, W2 As Long, H2 As Long 'RightBottom-Control
    Dim LC As Long, TC As Long, WC As Long, HC As Long 'mContainer-Control für einseitige Splitter

    'Es wird davon ausgegangen, daß alle controls auf dem mContainer liegen
    If mHorVer Then 'Horizontaler Splitter
        Select Case m12LTRB
        Case 1 '(01)1: LeftCtrl RghtSplt;
            L1 = 0:                     T1 = 0
            If mLeftTopPos = -1 Then mLeftTopPos = mLeftTopCtrl.Width
            W1 = mLeftTopPos + DX:      H1 = mContainer.Height
            W1 = mMax(W1, mMinLT)
            Wm = mWidth * STPX:         Hm = H1
            'W1 = mMin(W1, mContainer.Width - mMinRB - Wm)
            mLeftTopPos = W1
            Lm = W1:                    Tm = 0
            
            LC = mContainer.Left:       TC = mContainer.Top
            WC = W1 + Wm:               HC = H1
        Case 2 '(10)2: LeftSplt RghtCtrl;
            'If mLeftTopPos = -1 Then mLeftTopPos = 0
            LC = mContainer.Left + DX:  TC = mContainer.Top
            WC = mContainer.Width - DX: HC = mContainer.Height
                  
            Lm = 0:                     Tm = 0
            Wm = mWidth * STPX:         Hm = HC
            
            L2 = Wm:                    T2 = 0
            W2 = WC - Wm:               H2 = HC
            'W2 = mMax(W1, mMinLT)
            
            'W1 = mMin(W1, mContainer.Width - mMinRB - Wm)
            'mLeftTopPos = W1
            
            'If mLeftTopPos = -1 Then mLeftTopPos = mLeftTopCtrl.Width
            
        Case 3 '(11)3: LeftTopCtrl MidSplit RightBotCtrl
            L1 = 0:                     T1 = 0
            If mLeftTopPos = -1 Then mLeftTopPos = mLeftTopCtrl.Width
            W1 = mLeftTopPos + DX:      H1 = mContainer.Height
            W1 = mMax(W1, mMinLT)
            Wm = mWidth * STPX:         Hm = H1
            W1 = mMin(W1, mContainer.Width - mMinRB - Wm)
            mLeftTopPos = W1
            Lm = W1:                    Tm = 0
            
            L2 = W1 + Wm:               T1 = 0
            W2 = mContainer.Width - L2: H2 = H1
        End Select
    Else 'Vertikaler Splitter
        Select Case m12LTRB
        Case 1 '(01)1: TopCtrl BotSplt;
            L1 = 0:                     T1 = 0
            If mLeftTopPos = -1 Then mLeftTopPos = mLeftTopCtrl.Height
            W1 = mContainer.Width:      H1 = mLeftTopPos + DY
            H1 = mMax(H1, mMinLT)
            Wm = W1:                    Hm = mWidth * STPY
            'H1 = mMin(H1, mContainer.Height - mMinRB - Hm)
            mLeftTopPos = H1
            Lm = 0:                     Tm = H1
            
            LC = mContainer.Left:       TC = mContainer.Top
            WC = W1:                    HC = H1 + Hm
            
        Case 2 '(10)2: TopSplt BotCtrl;
        
        Case 3 '(11)3: LeftTopCtrl MidSplit RightBotCtrl
            
            L1 = 0:                     T1 = 0
            If mLeftTopPos = -1 Then mLeftTopPos = mLeftTopCtrl.Height
            W1 = mContainer.Width:      H1 = mLeftTopPos + DY
            H1 = mMax(H1, mMinLT)
            Wm = W1:                    Hm = mWidth * STPY
            H1 = mMin(H1, mContainer.Height - mMinRB - Hm)
            mLeftTopPos = H1
            Lm = 0:                     Tm = H1
            
            L2 = 0:                     T2 = H1 + mWidth * STPY
            W2 = W1:                    H2 = mContainer.Height - T2
        End Select
    End If
    DX = 0: DY = 0
      
    If Not mLeftTopCtrl Is Nothing Then
        If W1 >= 0 And H1 >= 0 Then
            mLeftTopCtrl.Move L1, T1, W1, H1
        End If
    Else
        If WC >= 0 And HC >= 0 Then
            mContainer.Move LC, TC, WC, HC
        End If
    End If
    
    If Not mPicSplit Is Nothing Then
        If Wm >= 0 And Hm >= 0 Then
            mPicSplit.Move Lm, Tm, Wm, Hm
        End If
    End If
    
    If Not mRghtBotCtrl Is Nothing Then
        If W2 >= 0 And H2 >= 0 Then
            mRghtBotCtrl.Move L2, T2, W2, H2
        End If
    Else
        If WC >= 0 And HC >= 0 Then
            mContainer.Move LC, TC, WC, HC
        End If
    End If
End Sub

Private Function mMin(V1, V2)
    If V1 < V2 Then mMin = V1 Else mMin = V2
End Function
Private Function mMax(V1, V2)
    If V1 > V2 Then mMax = V1 Else mMax = V2
End Function

Private Sub mPicSplit_Paint()
    Dim mR As mRect
    'mPicSplit.Borderstyle = 0
    Select Case mBorderStyle
    Case bsNone
        'mPicSplit.Borderstyle = 0
        'sonst nichts
    Case bsSunken
        mR = MakeBiggerRect(0, 0, mPicSplit.ScaleWidth / STPX, mPicSplit.ScaleHeight / STPY)
        Call DrawFrameControl(mPicSplit.hDC, mR, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
    Case bsRaised
        mR = MakeBiggerRect(0, 0, mPicSplit.ScaleWidth / STPX, mPicSplit.ScaleHeight / STPY)
        Call DrawFrameControl(mPicSplit.hDC, mR, DFC_BUTTON, DFCS_BUTTONPUSH)
    Case bsXPStyl
        mR = MakeBiggerRect(0, 0, mPicSplit.ScaleWidth / STPX, mPicSplit.ScaleHeight / STPY)
        mXPStyleC.Move (mR.Left) * STPX, mR.Top * STPY, (mR.Right - mR.Left) * STPX, (mR.Bottom - mR.Top) * STPY
    End Select
End Sub
Private Function MakeBiggerRect(l As Long, T As Long, W As Long, H As Long) As mRect
    Dim Bigger As Long, LW As Long, TH As Long
    Bigger = 2 * mWidth
    If mHorVer Then
        TH = Bigger: LW = 0
    Else
        TH = 0: LW = Bigger
    End If
    MakeBiggerRect.Left = l - LW - 1
    MakeBiggerRect.Top = T - TH
    MakeBiggerRect.Right = l + W + 2 * LW + 1
    MakeBiggerRect.Bottom = T + H + 2 * TH
End Function

'Private Sub Class_Terminate()
'    Set mMDIPar = Nothing
'    Set mParent = Nothing
'    Set mContainer = Nothing
'    Set mPicSplit = Nothing
'    Set mXPStyleC = Nothing
'    Set mLeftTopCtrl = Nothing
'    Set mRghtBotCtrl = Nothing
'End Sub



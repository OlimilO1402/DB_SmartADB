VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse erstellt einen Datenbank-Navigator
'
Private mStrName As String
Private mLastI As Long  'um festzustellen, ob nach oben oder nach unten gescrollt wird
'Private mCount As Long
'Buttons zum reinen Navigieren in den Datensätzen
Private WithEvents mBtnFirst As CommandButton
Attribute mBtnFirst.VB_VarHelpID = -1
Private WithEvents mWalker As HScrollBar
Attribute mWalker.VB_VarHelpID = -1
Private WithEvents mBtnLast As CommandButton
Attribute mBtnLast.VB_VarHelpID = -1
Private mPnlLblCnt As PictureBox
Private mLblCounter As Label
'Buttons zum Anfügen und Einfügen von Neuen und Löschen von vorhandenen Datensätzen
Private WithEvents mBtnAddNew As CommandButton 'Neuen Datensatz hintenanfügen
Attribute mBtnAddNew.VB_VarHelpID = -1
Private WithEvents mBtnInsNew As CommandButton 'Neuen Datensatz an akt. Pos. einfügen
Attribute mBtnInsNew.VB_VarHelpID = -1
Private WithEvents mBtnDelete As CommandButton 'akt. Datensatz löschen
Attribute mBtnDelete.VB_VarHelpID = -1
'Buttons für Bearbeitung von Datensätzen
Private WithEvents mBtnEdit   As CommandButton 'akt. Datensatz bearbeiten
Attribute mBtnEdit.VB_VarHelpID = -1
Private WithEvents mBtnCancel As CommandButton 'die akt. Datensatzbearbeitung abbrechen
Attribute mBtnCancel.VB_VarHelpID = -1
Private WithEvents mBtnUpdate As CommandButton 'die akt. Datensatzbearbeitung übernehmen
Attribute mBtnUpdate.VB_VarHelpID = -1
Public Enum EDataNavButton
  'Navigation
  nbFirst
  nbPrior
  nbNext
  nbLast
  'Edit
  nbAddNew
  nbInsNew
  nbDelete
  nbEdit
  nbUpdate
  nbCancel
End Enum
Public Event Click(Sender As Object, Button As EDataNavButton)

Public Sub NewC(aParentFrm As Form, aPanel As PictureBox, StrName As String, Optional MinIndex As Long, Optional CurIndex As Long, Optional MaxIndex As Long)
  'MsgBox "3"
  mStrName = StrName
  Set mBtnFirst = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnFirst", aPanel)
  Set mWalker = aParentFrm.Controls.Add("VB.HScrollBar", mStrName & "mWalker", aPanel)
  Set mBtnLast = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnLast", aPanel)
  Set mPnlLblCnt = aParentFrm.Controls.Add("VB.PictureBox", mStrName & "mPnlLblCnt", aPanel)
  Set mLblCounter = aParentFrm.Controls.Add("VB.Label", mStrName & "mLblCounter", mPnlLblCnt)
'Buttons zum Anfügen und Einfügen von Neuen und Löschen von vorhandenen Datensätzen
  Set mBtnAddNew = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnAddNew", aPanel)
  Set mBtnInsNew = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnInsNew", aPanel)
  Set mBtnDelete = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnDelete", aPanel)
'Buttons für Bearbeitung von Datensätzen
  Set mBtnEdit = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnEdit", aPanel)
  Set mBtnUpdate = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnPost", aPanel)
  Set mBtnCancel = aParentFrm.Controls.Add("VB.CommandButton", mStrName & "mBtnCancel", aPanel)
  'mCurIndex = CurIndex: mCount = Count

'Buttons zum Navigieren
  mBtnFirst.Visible = True
  mWalker.Visible = True
  mBtnLast.Visible = True
  mPnlLblCnt.Visible = True: mPnlLblCnt.BorderStyle = 0 'Kein
  mLblCounter.Visible = True
'Buttons zum Bearbeiten der Tabelle
  mBtnAddNew.Visible = True
  mBtnInsNew.Visible = True
  mBtnDelete.Visible = True
'Buttons zum Bearbeiten enes Datensatzes
  mBtnEdit.Visible = True
  mBtnUpdate.Visible = True
  mBtnCancel.Visible = True

'Buttons zum Navigieren
  mBtnFirst.Caption = "|<"
  
  If MaxIndex Then
    mWalker.Min = MinIndex
    mWalker.Value = CurIndex
    mLastI = CurIndex
    mWalker.Max = MaxIndex
  Else
    mWalker.Min = 0
    mWalker.Value = 0
    mWalker.Max = 0
  End If
  UpdateLblCounter
  mBtnLast.Caption = ">|"
'Buttons zum Anfügen, Einfügen, Löschen von Datensätzen
  mBtnAddNew.Caption = "New"
  mBtnInsNew.Caption = "Insert"
  mBtnDelete.Caption = "Delete"
'Buttons für Bearbeitung von Datensätzen
  mBtnEdit.Caption = "Edit"
  mBtnUpdate.Caption = "Update"
  mBtnCancel.Caption = "Cancel"
  Call MoveControls(aPanel)
End Sub
Private Sub MoveControls(aPanel As PictureBox)
Dim L As Single, T As Single, W As Single, H As Single
Dim W1 As Single, W2 As Single
Dim Brdr As Single
Dim Gap As Single
Dim BW1 As Single
Dim BW2 As Single
  'Select Case mPnl.BorderStyle
  'Case mPnl.BorderStyle =
  Select Case aPanel.BorderStyle
  Case 0: Brdr = 0 * 15 'Kein
  Case 1: Brdr = 2 * 15 'Fest Einfach
  End Select
  Gap = -15
  BW1 = 675
  BW2 = 675
  L = 0: T = 0:  W1 = 285: H = aPanel.Height - 2 * Brdr
  Call mBtnFirst.Move(L, T, W1, H): mBtnFirst.ZOrder 0
  L = L + W1 + Gap: W2 = 1575 '2535
  Call mWalker.Move(L, T, W2, H): mWalker.ZOrder 0
  L = L + W2 + Gap: W = 285
  Call mBtnLast.Move(L, T, W, H): mBtnLast.ZOrder 0
  
  L = L + W + Gap: W = BW1
  Call mBtnAddNew.Move(L, T, W, H): mBtnAddNew.ZOrder 0
  L = L + W + Gap: W = BW1
  Call mBtnInsNew.Move(L, T, W, H): mBtnInsNew.ZOrder 0
  L = L + W + Gap: W = BW1
  Call mBtnDelete.Move(L, T, W, H): mBtnDelete.ZOrder 0
  
  L = L + W + Gap: W = BW2
  Call mBtnEdit.Move(L, T, W, H): mBtnEdit.ZOrder 0
  L = L + W + Gap: W = BW2
  Call mBtnUpdate.Move(L, T, W, H): mBtnUpdate.ZOrder 0
  L = L + W + Gap: W = BW2
  Call mBtnCancel.Move(L, T, W, H): mBtnCancel.ZOrder 0
  
  W = L + W - 4 * Gap
  Call aPanel.Move(aPanel.Left, aPanel.Top, W)
  
  L = W1 + 255 + 2 * Gap
  W = W2 - 2 * 255 - 1 * Gap
  Call mPnlLblCnt.Move(L, T, W, H): mPnlLblCnt.ZOrder 0
  
  T = (H - 195) \ 2
  Call mLblCounter.Move(30, T)
  
End Sub
Private Sub UpdateLblCounter()
  mLblCounter.Caption = CStr(mWalker.Value) & " / " & CStr(mWalker.Max)
End Sub

'##########' Buttons zum Navigieren '##########'
Private Sub mBtnFirst_Click()
  'mCurIndex = mWalker.Min
  mWalker.Value = mWalker.Min 'mCurIndex
  UpdateLblCounter
  RaiseEvent Click(Me, nbFirst)
End Sub
Private Sub mWalker_Change()
  If (mWalker.Min <= mWalker.Value) And (mWalker.Value < mLastI) Then
    mLastI = mWalker.Value
    UpdateLblCounter
    If (mWalker.Value = mWalker.Min) Then
      RaiseEvent Click(Me, nbFirst)
    Else
      RaiseEvent Click(Me, nbPrior)
    End If
  ElseIf (mLastI < mWalker.Value) And (mWalker.Value <= mWalker.Max) Then
    mLastI = mWalker.Value
    UpdateLblCounter
    If (mWalker.Value = mWalker.Max) Then
      RaiseEvent Click(Me, nbLast)
    Else
      RaiseEvent Click(Me, nbNext)
    End If
  End If
End Sub
Private Sub mBtnLast_Click()
  'mCurIndex = mWalker.Max
  mWalker.Value = mWalker.Max 'mCurIndex
  UpdateLblCounter
  RaiseEvent Click(Me, nbLast)
End Sub

'##########'  Buttons zum Bearbeiten der Tabelle   '##########'
Private Sub mBtnAddNew_Click()
  RaiseEvent Click(Me, nbAddNew)
End Sub
Private Sub mBtnInsNew_Click()
  RaiseEvent Click(Me, nbInsNew)
End Sub
Private Sub mBtnDelete_Click()
  RaiseEvent Click(Me, nbDelete)
End Sub

'##########'  Buttons zum Bearbeiten eines Datensatzes   '##########'
Private Sub mBtnEdit_Click()
  RaiseEvent Click(Me, nbEdit)
End Sub
Private Sub mBtnCancel_Click()
  RaiseEvent Click(Me, nbCancel)
End Sub
Private Sub mBtnUpdate_Click()
  RaiseEvent Click(Me, nbUpdate)
End Sub

'##########'   Properties   '##########'
Public Property Let MinIndex(LngVal As Long)
  mWalker.Min = LngVal
End Property
Public Property Get MinIndex() As Long
  MinIndex = mWalker.Min
End Property

Public Property Let CurIndex(LngVal As Long)
  mWalker.Value = LngVal
End Property
Public Property Get CurIndex() As Long
  CurIndex = mWalker.Min
End Property

Public Property Let MaxIndex(LngVal As Long)
  mWalker.Max = LngVal
End Property
Public Property Get MaxIndex() As Long
  MaxIndex = mWalker.Max
End Property






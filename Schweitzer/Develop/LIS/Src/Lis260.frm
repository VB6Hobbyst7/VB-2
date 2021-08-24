VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm260MDefAnti 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   ClientHeight    =   6165
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00F4F0F2&
      Caption         =   "결 정"
      Height          =   510
      Left            =   2550
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   5625
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "취 소"
      Height          =   510
      Left            =   1215
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   5625
      Width           =   1320
   End
   Begin VB.CommandButton cmdUnselAll 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전체소거"
      Height          =   510
      Left            =   3900
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton cmdSelAll 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전체선택"
      Height          =   510
      Left            =   3030
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton cmdIns 
      BackColor       =   &H00F4F0F2&
      Caption         =   "추가"
      Height          =   510
      Left            =   2415
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   4875
      Width           =   600
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "제거"
      Height          =   510
      Left            =   1095
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   4875
      Width           =   600
   End
   Begin VB.ComboBox cboSpecies 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2910
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   600
      Width           =   2130
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00F4F0F2&
      Caption         =   "정렬"
      Height          =   510
      Left            =   495
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   4875
      Width           =   600
   End
   Begin VB.ListBox lstAnti 
      BackColor       =   &H00F1F5F4&
      Columns         =   3
      Height          =   3840
      Left            =   1920
      Style           =   1  '확인란
      TabIndex        =   1
      Top             =   960
      Width           =   3090
   End
   Begin FPSpread.vaSpread ssAnti 
      Height          =   4215
      Left            =   105
      TabIndex        =   0
      Top             =   585
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   7435
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   3
      MaxRows         =   100
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "Lis260.frx":0000
      UserResize      =   0
   End
   Begin MedControls1.LisLabel lblMicNm 
      Height          =   405
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   714
      BackColor       =   -2147483635
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   10
      Left            =   1905
      TabIndex        =   11
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "검     체"
      Appearance      =   0
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   4905
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frm260MDefAnti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fRIndex As Integer
Private fParentForm As Form

Private cRed As Long
Private cBlack As Long

Private objMicRst As New clsLISMicResult


Public Sub SetCurAnti(ByVal pForm As Form, ByVal pIdx As Integer, ByVal pMicNm As String, ByVal pAntiCnt As Integer, ByVal pBuf As String)
    
    Dim sTmp As String
    
    Set fParentForm = pForm
    
    fRIndex = (pIdx) * 3 + 2
    lblMicNm.Caption = pIdx + 1 & ". " & pMicNm
    
    Dim i As Integer
    ssAnti.MaxRows = pAntiCnt
    For i = 1 To pAntiCnt
        sTmp = medShift(pBuf, ";")
        ssAnti.Col = 1: ssAnti.Row = i: ssAnti.Text = medGetP(sTmp, 1, ":")
        ssAnti.Col = 2: ssAnti.Row = i: ssAnti.Text = medGetP(sTmp, 2, ":")
        ssAnti.Col = 3: ssAnti.Row = i: ssAnti.Text = medGetP(sTmp, 3, ":")
    Next i
    
End Sub

Private Sub cboSpecies_Click()
     Call objMicRst.LoadAnti(cboSpecies.List(cboSpecies.ListIndex), lstAnti)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frm260MDefAnti = Nothing
End Sub

Private Sub cmdDel_Click()
    
    Dim i As Integer
    
    ssAnti.Col = 1
    For i = ssAnti.MaxRows To 1 Step -1
        ssAnti.Row = i
        If ssAnti.ForeColor = cRed Then Call DelAnti(i)
    Next i

End Sub

Private Sub DelAnti(ByVal pIdx As Integer)
    
    ssAnti.Row = pIdx
    ssAnti.Action = ActionDeleteRow
    ssAnti.MaxRows = ssAnti.MaxRows - 1

End Sub

Private Sub cmdIns_Click()

    Dim i As Integer
    
    For i = 0 To lstAnti.ListCount - 1
        If lstAnti.Selected(i) And Not ExistAnti(lstAnti.List(i)) Then
            Call InsAnti(lstAnti.List(i))
        End If
        lstAnti.Selected(i) = False
    Next i

End Sub

Private Function ExistAnti(ByVal pAnti As String) As Boolean

    Dim i As Integer
    
    ssAnti.Col = 1
    For i = 1 To ssAnti.MaxRows
        ssAnti.Row = i
        If ssAnti.Text = pAnti Then ExistAnti = True: Exit Function
    Next i

    ExistAnti = False

End Function

Private Sub InsAnti(ByVal pAnti As String)
    
    Dim tRow As String

    tRow = ssAnti.MaxRows + 1
    ssAnti.MaxRows = tRow
    
    ssAnti.Col = 1: ssAnti.Row = tRow
    ssAnti.Text = pAnti

End Sub

Private Sub cmdOk_Click()
    
    Dim sCnt As Integer, sTmp As String, sBuf As String

    Dim i As Integer
    
    sBuf = ""
    sCnt = ssAnti.MaxRows
    For i = 1 To sCnt
        ssAnti.Col = 1: ssAnti.Row = i: sTmp = ssAnti.Text
        ssAnti.Col = 2: ssAnti.Row = i: sTmp = sTmp & ":" & ssAnti.Text
        ssAnti.Col = 3: ssAnti.Row = i: sTmp = sTmp & ":" & ssAnti.Text
        sBuf = sBuf & sTmp & ";"
    Next i

    Call fParentForm.ApplyDefAnti(fRIndex, sCnt, sBuf)
    
    Unload Me
    Set frm260MDefAnti = Nothing

End Sub

Private Sub cmdSelAll_Click()
    
    Dim i As Integer
    
    For i = 0 To lstAnti.ListCount - 1
        lstAnti.Selected(i) = True
    Next i

End Sub

Private Sub cmdSort_Click()
    
    ssAnti.Col = -1: ssAnti.Row = -1
    ssAnti.ForeColor = cBlack
        
    ssAnti.SortBy = SortByRow
    ssAnti.SortKey(1) = 1
    ssAnti.SortKeyOrder(1) = SortKeyOrderAscending
    ssAnti.Col = 1
    ssAnti.COL2 = ssAnti.MaxCols
    ssAnti.Row = 1
    ssAnti.Row2 = ssAnti.MaxRows
    ssAnti.Action = ActionSort

End Sub


Private Sub cmdUnselAll_Click()
    
    Dim i As Integer
    
    For i = 0 To lstAnti.ListCount - 1
        lstAnti.Selected(i) = False
    Next i

End Sub

Private Sub Form_Load()
    
    Me.Top = 4500
    Me.Left = 9500

    cRed = RGB(255, 0, 0)
    cBlack = RGB(0, 0, 0)

    Call objMicRst.LoadMicSpecies(cboSpecies)
    If cboSpecies.ListCount > 0 Then cboSpecies.ListIndex = 0
    
    Call objMicRst.LoadAnti(cboSpecies.List(cboSpecies.ListIndex), lstAnti)

End Sub


Private Sub ssAnti_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpcolor As Long
    
    If Col >= 0 And Row > 0 Then
    
        ssAnti.Col = -1: ssAnti.Row = Row
        tmpcolor = ssAnti.ForeColor
        
        If tmpcolor = cRed Then
            ssAnti.ForeColor = cBlack
        Else
            ssAnti.ForeColor = cRed
        End If
        
    End If

End Sub


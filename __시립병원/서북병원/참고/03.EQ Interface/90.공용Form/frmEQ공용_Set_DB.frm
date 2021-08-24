VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEQ공용_Set_DB 
   BorderStyle     =   1  '단일 고정
   Caption         =   "DB 접속정보"
   ClientHeight    =   2925
   ClientLeft      =   7575
   ClientTop       =   1035
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_Set_DB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4320
   Begin VB.TextBox txtDB_ID 
      Height          =   330
      IMEMode         =   10  '한글 
      Left            =   1260
      TabIndex        =   1
      Text            =   "txtDB_ID"
      Top             =   1140
      Width           =   2955
   End
   Begin VB.TextBox txtDB_PW 
      Height          =   330
      IMEMode         =   10  '한글 
      Left            =   1260
      TabIndex        =   2
      Text            =   "txtDB_PW"
      Top             =   1500
      Width           =   2955
   End
   Begin VB.TextBox txtDB_SERVER 
      Height          =   330
      IMEMode         =   10  '한글 
      Left            =   1260
      TabIndex        =   3
      Text            =   "txtDB_SERVER"
      Top             =   1860
      Width           =   2955
   End
   Begin VB.TextBox txtDB_NM 
      Height          =   330
      IMEMode         =   10  '한글 
      Left            =   1260
      TabIndex        =   4
      Text            =   "txtDB_NM"
      Top             =   2220
      Width           =   2955
   End
   Begin VB.ComboBox cboDBTYPE 
      Height          =   300
      ItemData        =   "frmEQ공용_Set_DB.frx":263A
      Left            =   1260
      List            =   "frmEQ공용_Set_DB.frx":263C
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   780
      Width           =   2955
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   9
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "계정 ID"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "계정 PW"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "DB SERVER"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "DB NAME"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "※ DB NAME은 SQL 서버를 사용할 경우 입력"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   2640
      Width           =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "DB TYPE"
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   630
   End
   Begin VB.Label lblDBSET 
      BackStyle       =   0  '투명
      Caption         =   "DB Connection Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "frmEQ공용_Set_DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SUB_MM_CANCEL()
    barStatus.Max = 100
    barStatus.Value = 100
    
    Call SUB_MM_KEY_CLEAR
End Sub

Private Sub SUB_MM_INITIAL()
    lblDBSET.Caption = gstrArgTemp1 & " DB Connection Information"
    
    Me.Height = 3405
    Me.Width = 4440
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    GoSub ADD_ITEM
    
    Call SUB_MM_CANCEL
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    '/DB TYPE Setting
    cboDBTYPE.AddItem ""
    cboDBTYPE.AddItem "Oracle 8i" & Space(20) & "01"
    cboDBTYPE.AddItem "Oracle 9i" & Space(20) & "02"
    cboDBTYPE.AddItem "Oracle 10g" & Space(20) & "03"
    cboDBTYPE.AddItem "Oracle 11g" & Space(20) & "04"
    cboDBTYPE.AddItem "SQL Server 2000" & Space(20) & "11"
    cboDBTYPE.AddItem "SQL Server 2005" & Space(20) & "12"
    cboDBTYPE.AddItem "SQL Server 2008" & Space(20) & "13"
    cboDBTYPE.AddItem "Sybase" & Space(20) & "21"
Return
End Sub

Private Sub SUB_MM_KEY_CLEAR()
    cboDBTYPE.ListIndex = -1
    txtDB_ID = ""
    txtDB_PW = ""
    txtDB_SERVER = ""
    txtDB_NM = ""
End Sub

Public Function FUNC_MM_VIEW() As Boolean
    FUNC_MM_VIEW = False
    
On Error GoTo RTN_ERR

    Call SUB_MM_KEY_CLEAR
    
    If ConnDB_LOC = False Then End
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM CUS_MST "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        If Mid(lblDBSET, 1, 3) = "HIS" Then                                 '/HISDB 셋팅일때
            Call SET_CBO_DT_R(Trim(ADR_LOC!HISDB_TYPE & ""), cboDBTYPE)     '/DB TYPE
            txtDB_ID = Trim(ADR_LOC!HISDB_ID & "")                          '/계정 ID
            txtDB_PW = Trim(ADR_LOC!HISDB_PW & "")                          '/계정 PW
            txtDB_SERVER = Trim(ADR_LOC!HISDB_SERVER & "")                  '/계정 SERVER
            txtDB_NM = Trim(ADR_LOC!HISDB_DBNM & "")                        '/계정 NAME
            
        Else                                                                '/ETCDB 셋팅일때
            Call SET_CBO_DT_R(Trim(ADR_LOC!ETCDB_TYPE & ""), cboDBTYPE)     '/DB TYPE
            txtDB_ID = Trim(ADR_LOC!ETCDB_ID & "")                          '/계정 ID
            txtDB_PW = Trim(ADR_LOC!ETCDB_PW & "")                          '/계정 PW
            txtDB_SERVER = Trim(ADR_LOC!ETCDB_SERVER & "")                  '/계정 SERVER
            txtDB_NM = Trim(ADR_LOC!ETCDB_DBNM & "")                        '/계정 NAME
        End If
        
        ADR_LOC.Close: Set ADR_LOC = Nothing
    End If
 
    Call CloseDB_LOC
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_SAVE() As Boolean
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM CUS_MST "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        gstrQuy = "UPDATE CUS_MST SET "
        If Mid(lblDBSET, 1, 3) = "HIS" Then                                                                 '/HISDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_TYPE    = '" & Trim(Right(cboDBTYPE, 2)) & "', "     '/DB TYPE
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_ID      = '" & Trim(txtDB_ID) & "', "                '/계정 ID
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_PW      = '" & Trim(txtDB_PW) & "', "                '/계정 PW
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_SERVER  = '" & Trim(txtDB_SERVER) & "', "            '/계정 SERVER
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_DBNM    = '" & Trim(txtDB_NM) & "' "                 '/계정 NAME
        Else                                                                                                '/ETCDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_TYPE    = '" & Trim(Right(cboDBTYPE, 2)) & "', "     '/DB TYPE
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_ID      = '" & Trim(txtDB_ID) & "', "                '/계정 ID
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_PW      = '" & Trim(txtDB_PW) & "', "                '/계정 PW
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_SERVER  = '" & Trim(txtDB_SERVER) & "', "            '/계정 SERVER
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_DBNM    = '" & Trim(txtDB_NM) & "' "                 '/계정 NAME
        End If
    Else
        gstrQuy = "INSERT INTO CUS_MST "
        If Mid(lblDBSET, 1, 3) = "HIS" Then                                                             '/HISDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & " (HISDB_TYPE, "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_ID, "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_PW, "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_SERVER, "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_DBNM) "
            gstrQuy = gstrQuy & vbCrLf & " VALUES "
            gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(Right(cboDBTYPE, 2)) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_ID) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_PW) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_SERVER) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_NM) & "') "
        Else                                                                                            '/ETCDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & " (ETCDB_TYPE,  "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_ID, "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_PW, "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_SERVER, "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_DBNM) "
            gstrQuy = gstrQuy & vbCrLf & " VALUES "
            gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(Right(cboDBTYPE, 2)) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_ID) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_PW) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_SERVER) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_NM) & "') "
        End If
    End If
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_SAVE = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboDBTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장질의") = vbCancel Then Exit Sub
        
    If FUNC_MM_SAVE = True Then
        MsgBox "저장되었습니다!", vbInformation, "확인"
        Call SUB_MM_CANCEL
        Call FUNC_MM_VIEW
        
        MsgBox "환경설정이 변경되었습니다." & vbCrLf & _
               "프로그램을 재 실행하십시오!", vbInformation, "프로그램 종료"
        End
    Else
        MsgBox "저장오류!", vbCritical, "확인"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdQuit_Click
    End Select
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    Call FUNC_MM_VIEW
End Sub

Private Sub txtDB_ID_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDB_ID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDB_NM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDB_NM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDB_PW_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDB_PW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDB_SERVER_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDB_SERVER_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

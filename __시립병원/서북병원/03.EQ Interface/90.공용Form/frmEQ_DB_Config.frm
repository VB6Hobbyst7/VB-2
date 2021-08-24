VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEQ_DB_Config 
   Caption         =   "DB Setting"
   ClientHeight    =   4215
   ClientLeft      =   10275
   ClientTop       =   3225
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      Caption         =   "[접속정보]"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
      Begin VB.TextBox txtDB_NM 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1860
         TabIndex        =   12
         Text            =   "txtDB_NM"
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox txtDB_SERVER 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1860
         TabIndex        =   10
         Text            =   "txtDB_SERVER"
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox txtDB_PW 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1860
         TabIndex        =   8
         Text            =   "txtDB_PW"
         Top             =   660
         Width           =   1995
      End
      Begin VB.TextBox txtDB_ID 
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1860
         TabIndex        =   6
         Text            =   "txtDB_ID"
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "※ SQL 서버를 사용할 경우에 사용"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   1020
         TabIndex        =   14
         Top             =   1560
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "DB NAME"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "DB SERVER"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "계정 PW"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "계정 ID"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.ComboBox cboDBTYPE 
      Height          =   300
      ItemData        =   "frmEQ_DB_Config.frx":0000
      Left            =   1965
      List            =   "frmEQ_DB_Config.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   1380
      Width           =   1995
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
      TabIndex        =   2
      Top             =   660
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
      TabIndex        =   1
      Top             =   660
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   15
      Top             =   1200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "DB TYPE"
      Height          =   195
      Index           =   8
      Left            =   300
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDBSET 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   " DB Connection Information"
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
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   555
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4155
   End
End
Attribute VB_Name = "frmEQ_DB_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function MM_CANCEL() As Boolean
    barStatus.Max = 100
    barStatus.Value = 100

    
    Call MM_KEY_CLEAR
End Function

Private Sub MM_INITIAL()
    
    lblDBSET.Caption = gstrArgTemp1 & "  DB Connection Information "
    Me.Height = 4875
    Me.Width = 4440
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    
    '/DB TYPE Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/DB TYPE Setting----------------------------------------------------------------------------------------------------/
    
    
ADD_ITEM:
    '/DB,TYPE 설정
    cboDBTYPE.AddItem ""
    cboDBTYPE.AddItem "Oracle 8i" & Space(20) & "01"
    cboDBTYPE.AddItem "Oracle 9i" & Space(20) & "02"
    cboDBTYPE.AddItem "Oracle 10g" & Space(20) & "03"
    cboDBTYPE.AddItem "Oracle 11g" & Space(20) & "04"
    cboDBTYPE.AddItem "SQL Server 2000" & Space(20) & "11"
    cboDBTYPE.AddItem "SQL Server 2005" & Space(20) & "12"
    cboDBTYPE.AddItem "SQL Server 2008" & Space(20) & "13"
    cboDBTYPE.AddItem "Sybase" & Space(20) & "21"
    
    Call MM_CANCEL
End Sub

Private Sub MM_KEY_CLEAR()
    
    txtDB_ID = ""
    txtDB_PW = ""
    txtDB_SERVER = ""
    txtDB_NM = ""
    
End Sub

Public Function MM_VIEW() As Boolean
    
    
    Call MM_KEY_CLEAR
    
    If ConnDB_LOC = False Then Exit Function
    
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
End Function

Public Function FUNC_MM_SAVE() As Boolean
    
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장질의") = vbCancel Then Exit Function
    
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then End
    
    ADC_LOC.BeginTrans
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM CUS_MST "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        gstrQuy = "UPDATE CUS_MST SET "
        
        If Mid(lblDBSET, 1, 3) = "HIS" Then                                                                 '/HISDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_TYPE    =  '" & Trim(Right(cboDBTYPE, 2)) & "', "    '/DB TYPE
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_ID      =  '" & Trim(txtDB_ID) & "', "               '/계정 ID
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_PW      =  '" & Trim(txtDB_PW) & "', "               '/계정 PW
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_SERVER  =  '" & Trim(txtDB_SERVER) & "', "           '/계정 SERVER
            gstrQuy = gstrQuy & vbCrLf & "       HISDB_DBNM    =  '" & Trim(txtDB_NM) & "' "                '/계정 NAME
            
        Else                                                                                                '/ETCDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_TYPE    =  '" & Trim(Right(cboDBTYPE, 2)) & "', "    '/DB TYPE
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_ID      =  '" & Trim(txtDB_ID) & "', "               '/계정 ID
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_PW      =  '" & Trim(txtDB_PW) & "', "               '/계정 PW
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_SERVER  =  '" & Trim(txtDB_SERVER) & "', "           '/계정 SERVER
            gstrQuy = gstrQuy & vbCrLf & "       ETCDB_DBNM    =  '" & Trim(txtDB_NM) & "' "                '/계정 NAME
            
        End If
    Else
        gstrQuy = "INSERT INTO CUS_MST "
        
        If Mid(lblDBSET, 1, 3) = "HIS" Then                                                             '/HISDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & " (HISDB_TYPE,  "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_ID, HISDB_PW, "
            gstrQuy = gstrQuy & vbCrLf & "  HISDB_SERVER, HISDB_DBNM) "
            gstrQuy = gstrQuy & vbCrLf & " VALUES "
            gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(Right(cboDBTYPE, 2)) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_ID) & "', '" & Trim(txtDB_PW) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtDB_SERVER) & "', '" & Trim(txtDB_NM) & "') "
            
        Else                                                                                            '/ETCDB 셋팅일때
            gstrQuy = gstrQuy & vbCrLf & " (ETCDB_TYPE,  "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_ID, ETCDB_PW, "
            gstrQuy = gstrQuy & vbCrLf & "  ETCDB_SERVER, ETCDB_DBNM) "
            gstrQuy = gstrQuy & vbCrLf & " VALUES "
            gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(Val(Right(cboDBTYPE, 2))) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Val(txtDB_ID)) & "', '" & Trim(Val(txtDB_PW)) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Val(txtDB_SERVER)) & "', '" & Trim(Val(txtDB_NM)) & "') "
            
        End If
    End If
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_SAVE = True

    
    MsgBox "저장되었습니다!", vbInformation, "확인"
    
    Call MM_CANCEL
    
    Call MM_VIEW
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call FUNC_MM_SAVE
End Sub

Private Sub Form_Load()
    Call MM_INITIAL
    
    Call MM_VIEW
End Sub


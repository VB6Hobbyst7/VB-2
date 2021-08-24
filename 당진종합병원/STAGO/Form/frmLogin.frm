VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   3435
   ClientLeft      =   2805
   ClientTop       =   3060
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2029.513
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5450.582
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtInstCd 
      Height          =   270
      Left            =   4080
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   270
      Left            =   2850
      TabIndex        =   14
      Top             =   750
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chkPW 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "아이디저장"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4110
      TabIndex        =   12
      Top             =   2490
      Width           =   1425
   End
   Begin VB.Timer Timer1 
      Left            =   2250
      Top             =   1410
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   8  '영문
      Left            =   2715
      TabIndex        =   2
      Top             =   2190
      Width           =   1245
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1665
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2775
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   2715
      TabIndex        =   0
      Top             =   2490
      Width           =   1245
   End
   Begin HSCotrol.CButton cmdOK 
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      BackColor       =   16777215
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      BackColor       =   16777215
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblSvr 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▒ 사용처 : 강남성심병원"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 명:"
      Height          =   180
      Index           =   2
      Left            =   1770
      TabIndex        =   13
      Top             =   2550
      Width           =   840
   End
   Begin VB.Label lblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▒ 사용처 : 강남성심병원"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   90
      TabIndex        =   11
      Top             =   1770
      Width           =   2835
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "▒ BIOFLEX 2200 Interface ▒"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   105
      Width           =   4245
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   9
      Top             =   930
      Width           =   405
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  '투명
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID(&U):"
      Height          =   180
      Index           =   0
      Left            =   1515
      TabIndex        =   7
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "암호(&P):"
      Height          =   180
      Index           =   1
      Left            =   870
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "사용자 ID를 입력 하십시오."
      Height          =   180
      Left            =   390
      TabIndex        =   5
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":030A
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":0454
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":059E
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet4 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":06E8
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   2010
      Left            =   30
      Picture         =   "frmLogin.frx":0832
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5745
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   30
      Picture         =   "frmLogin.frx":2E6F
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   5745
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldUid          As String
Private OldPwd          As String
Private MsgFg           As Boolean
Private OldUser         As UserInfo

Public CancelIsEnd      As Boolean
Public LoginSucceeded   As Boolean

Private adoRS As ADODB.Recordset
Dim gwTmp1 As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Call ReleaseCapture
   Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub cmdCancel_Click()
    If MainForm Is Nothing Then
        Call Unload(Me)
        Set frmLogin = Nothing
        End
    Else
        CurrUser = OldUser
        Call Unload(Me)
        Set frmLogin = Nothing
    End If

End Sub

Private Sub cmdOk_Click()
    Dim ShowAtStartup As Variant
    Dim objUserInf As clsCommon
    Dim strUID As String
    Dim strPWD As String
    Dim strSvrcCnt  As Variant
    Dim strSvrcData As Variant
        
    Timer1.Enabled = False
    imgNet4.ZOrder
    
'    strSvrcData = getSvrcInfo("SL_USERM_L1", "020487", "", "", "", "", "", "", "", "")
'    strSvrcData = getSvrcInfo("AP_PATBA_S7", "030480287", "", "", "", "", "", "", "", "")
    
    
    Set objUserInf = New clsCommon
    
    strUID = Trim(txtUserID.text)
    strPWD = Trim(txtPassword.text)
    
    
    If CurrUser.CuUserPW = strPWD Then
        If CurrUser.CuPower = Authority.ELVEL_NOT Then
            MsgBox "실행 권한이 없읍니다. 관리자에게 문의 하세요. ", vbOKOnly + vbExclamation
            Exit Sub
        End If
        
        If chkPW.Value = 1 Then
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_SAVEPW, "1")
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_UID, strUID)
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_PWD, strPWD)
        Else
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_SAVEPW, "")
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_UID, "")
            Call SaveString(HKEY_CURRENT_USER, REG_JETDB, REG_PWD, "")
        End If
        
        Call Unload(Me)
        
        If MainForm Is Nothing Then
            Set MainForm = New MDIMain
            MainForm.Show
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        Else
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        End If
      
      Else
         MsgBox "비밀번호가 틀립니다. 비밀번호를 확인하세요. ", vbOKOnly + vbExclamation
         txtPassword.SetFocus
         txtPassword.SelStart = 0
         txtPassword.SelLength = Len(txtPassword)
      End If

End Sub

Private Sub Form_Activate()
    txtUserID.SetFocus
End Sub

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "URL", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtURL.text = Trim(db_tmp)

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "INSTCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtInstCd.text = Trim(db_tmp)

    GetSetup = True

End Function

Private Sub Form_Load()

Dim strSvrcData As String

    imgNet1.ZOrder 0
    Timer1.interval = 500
    Timer1.Enabled = True
    
    Call GetSetup
    
    lblTitle.Caption = App.Title
    lblVersion.Caption = "Ver. " & App.major & "." & App.minor & "." & App.Revision
    lblSite.Caption = " ▒ 사용처 : " & App.CompanyName
    lblSvr.Caption = txtURL.text
    
    
    If Not MainForm Is Nothing Then
        OldUser = CurrUser
    End If
    
    
    
    'txtUserID.text = "020487" park770728
    
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_SAVEPW)) > 0 Then
        txtUserID.text = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_UID)
'        Call txtUserID_LostFocus
        txtPassword.text = GetString(HKEY_CURRENT_USER, REG_JETDB, REG_PWD)
        
        chkPW.Value = "1"
    
        If Trim(txtUserID.text) <> "" Then
            'strSvrcData = getSvrcInfo("SL_USERM_L1", Trim(txtUserID.text))
            
            Call txtUserID_LostFocus
'            If strSvrcData = "" Then
'                MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation
'                With txtUserID
'                    .SetFocus
'                    .SelStart = 0
'                    .SelLength = Len(.text)
'                End With
'            End If
        End If
    End If
    
End Sub

Private Sub Timer1_Timer()
    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If

End Sub

Private Sub txtPassword_GotFocus()
   With txtPassword
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserID_Change()
   txtUserName.text = ""
End Sub

Private Sub txtUserID_GotFocus()
   With txtUserID
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
        Call txtUserID_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtUserID_LostFocus()
    Dim ret As Boolean

    Dim objUserInf As clsCommon
    
    On Error GoTo ErrorTrap

    Dim sHtmlLine
    Dim sUrl, sPost, sParam As String
    Dim sRcvData, sData As String
        

    If ActiveControl.Name = "cmdCancel" Then Exit Sub

        If txtUserID.text = "" Then
            MsgFg = True
            MsgBox "로그온 ID를 입력하세요. ", vbOKOnly + vbExclamation
            MsgFg = False
            txtUserID.SetFocus
            Exit Sub
        End If

        labMsg.Caption = "데이타 베이스에 연결중 ...."
        Screen.MousePointer = vbArrowHourglass

'        sUrl = "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?"
        sUrl = txtURL.text
        
                 sParam = "submit_id=TRLII00104&"
        sParam = sParam & "business_id=lis&"
        sParam = sParam & "ex_interface=" & Trim(txtUserID.text) & "|012&" '사용자ID|기관코드
        sParam = sParam & "instcd=012&" '기관코드
        sParam = sParam & "userid=" & Trim(txtUserID.text) '사용자ID
        
        sRcvData = OpenURLWithIE2(sUrl & sParam, Inet1)
        
            
        If InStr(1, sRcvData, "<?xml version") > 0 Then
            gwTmp1 = ""
        End If
        
        gwTmp1 = gwTmp1 & sRcvData
                
'        XML_Parsing gwTmp1
        
        'sData = mGetP(gwTmp1, 1, "usernm")
        sData = mGetP(mGetP(mGetP(gwTmp1, 2, "usernm"), 2, ">"), 1, "<")
        If sData = "" Then
            MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation
            With txtUserID
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        End If

        Screen.MousePointer = vbDefault
        labMsg.Caption = "데이타 베이스에 연결 되었습니다."

        If sData = "" Then
            MsgBox "등록되지 않은 ID입니다. 로그인 ID를 확인하세요. ", vbOKOnly + vbExclamation
            Set AdoRs_Jet = Nothing
            Set objUserInf = Nothing
            With txtUserID
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        Else
            Timer1.Enabled = False
            With CurrUser
'                varSvcData = Split(strSvrcData, "|")
                .CuUserID = Trim(txtUserID.text)
                .CuUserNM = sData
                .CuUserPW = ""
                txtUserName = .CuUserNM
            End With
            imgNet4.ZOrder 0
            'txtPassword.SetFocus
            cmdOk.SetFocus
        End If
    Exit Sub
    
ErrorTrap:
'    Set AdoRs_Jet = Nothing
'    Set objUserInf = Nothing
    labMsg.Caption = "사용자 ID를 확인하세요"

End Sub

 
'''
'''
'''
'''Public Function GetHtml(ByVal sUrl As String, Optional ByVal Stype As String = "GET", Optional ByVal sHost As String = "", Optional ByVal sBody As Variant = "", Optional ByVal sCookie As String, Optional ByVal sRefer As String = "") As Variant
'''    'On Error GoTo Err
'''
'''    Dim oWinHttp As Object
'''    Dim Result(4) As String
'''    Dim TotBuf() As Byte, ChunkedBuf() As Byte, Converted() As Byte, ni As Long
'''    Dim lSize As Long
''''MsgBox (sUrl)
'''    Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'''    With oWinHttp
'''        If sRefer = "" Then sRefer = sUrl
'''
'''        .Open Stype, sUrl, 0
'''        .setRequestHeader "Referer", sRefer
'''        .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; InfoPath.2; .NET CLR 2.0.50727)"
'''        '.SetRequestHeader "Host", sHost
'''        If sCookie <> "" Then
'''            .setRequestHeader "Cookie", sCookie
'''        End If
'''
'''        '.SetRequestHeader "Content-Length", 10 'Len(sBody)
'''
'''        If Stype = "GET" Then
'''            .send
'''        End If
'''        If Stype = "POST" Then
'''            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'''            .send sBody
'''        End If
'''
'''        .waitForResponse
'''        Result(0) = .Status
'''        Result(1) = .getAllResponseHeaders
'''        sCookie = Cookie_Exe(.getAllResponseHeaders, sCookie)
'''        Result(3) = sCookie
'''
'''        If InStr(Result(1), "Content-Type: text/html; charset=UTF-8") > 0 Then
'''            TotBuf = .responseBody
'''
'''            lSize = MultiByteToWideChar(CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, ByVal 0&, 0&)
'''
'''            ReDim Converted(lSize * 2 - 1)
'''            MultiByteToWideChar CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, Converted(0), lSize
'''        Else
'''            Converted = StrConv(.responseBody, vbUnicode)
'''        End If
'''        If .Status = 200 Then
'''            Result(2) = Converted
'''        Else
'''            Result(2) = .statusText
'''
'''        End If
'''    End With
'''    GetHtml = Result
'''    Exit Function
'''Err:
'''    Result(0) = "페이지못참음"
'''    GetHtml = Result
'''End Function
'''
''''쿠키 정리
'''
'''Public Function Cookie_Exe(ByVal sHeader As String, Optional ByVal sCookie As String)
'''
'''    Dim Tmp1() As String, Tmp2() As String, Tmp3() As String
'''    Dim i As Integer, j As Integer
'''    Dim nCookie1() As String, nCookie2() As String
'''    Dim rCookie As String
'''    Dim chk As Boolean
'''
'''    If InStr(sHeader, "Set-Cookie: ") <= 0 Then
'''        Cookie_Exe = sCookie
'''        Exit Function
'''    End If
'''
'''    Tmp1 = Split(sHeader, "Set-Cookie: ")
'''    For i = 1 To UBound(Tmp1)
'''        Tmp2 = Split(Tmp1(i), "; ")
'''        Tmp3 = Split(Tmp2(0), "=")
'''        ReDim Preserve nCookie1(i)
'''        ReDim Preserve nCookie2(i)
'''        nCookie1(i - 1) = Tmp3(0)
'''        nCookie2(i - 1) = Tmp3(1)
'''    Next i
'''
'''    If sCookie <> "" Then
'''        Tmp1 = Split(sCookie, ";")
'''        For i = 0 To UBound(Tmp1)
'''            Tmp2 = Split(Tmp1(i), "=")
'''            chk = False
'''            For j = 0 To UBound(nCookie1)
'''                If Tmp2(0) = nCookie1(j) Then chk = True
'''            Next j
'''            If chk = False Then
'''                ReDim Preserve nCookie1(UBound(nCookie1) + 1)
'''                ReDim Preserve nCookie2(UBound(nCookie2) + 1)
'''                nCookie1(UBound(nCookie1) - 1) = Tmp2(0)
'''                nCookie2(UBound(nCookie1) - 1) = Tmp2(1)
'''            End If
'''        Next i
'''    End If
'''
'''    If UBound(nCookie1) > 0 Then
'''        For i = 0 To UBound(nCookie1) - 1
'''            rCookie = rCookie & nCookie1(i) & "=" & nCookie2(i) & "; "
'''        Next i
'''    End If
'''    rCookie = Mid(rCookie, 1, Len(rCookie) - 2)
'''    Cookie_Exe = rCookie
'''    Exit Function
'''End Function



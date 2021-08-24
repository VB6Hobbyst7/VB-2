VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPcxView 
   Caption         =   "PCX 결과조회"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   Icon            =   "frmPcxView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmPcxView.frx":0E42
   ScaleHeight     =   8250
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   10125
      Begin VB.Timer tmrSearch 
         Left            =   9660
         Top             =   0
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   0
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8880
         TabIndex        =   1
         Top             =   150
         Width           =   1185
      End
      Begin VB.ComboBox cboSpcPos 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   4680
         TabIndex        =   4
         Top             =   210
         Width           =   1305
      End
      Begin VB.CheckBox chkPopup 
         Caption         =   "팝업보기"
         Height          =   255
         Left            =   6090
         TabIndex        =   5
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   25559041
         CurrentDate     =   39246
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   25559041
         CurrentDate     =   39246
      End
      Begin VB.Label Label1 
         Caption         =   "조회기간"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         Height          =   225
         Left            =   2310
         TabIndex        =   14
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label3 
         Caption         =   "장비명"
         Height          =   195
         Left            =   4050
         TabIndex        =   13
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdInTaskBar 
      Caption         =   "숨김"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10890
      TabIndex        =   11
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdImgSet 
      Caption         =   "배경"
      Height          =   375
      Left            =   10290
      TabIndex        =   10
      Top             =   30
      Width           =   570
   End
   Begin VB.OptionButton optImg 
      Height          =   285
      Index           =   2
      Left            =   10320
      TabIndex        =   9
      Top             =   3180
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton optImg 
      Height          =   285
      Index           =   1
      Left            =   10290
      TabIndex        =   8
      Top             =   2010
      Width           =   195
   End
   Begin VB.OptionButton optImg 
      Height          =   285
      Index           =   0
      Left            =   10320
      TabIndex        =   7
      Top             =   930
      Width           =   195
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   7560
      Left            =   0
      TabIndex        =   6
      Top             =   690
      Width           =   10185
      _Version        =   393216
      _ExtentX        =   17965
      _ExtentY        =   13335
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   15
      MaxRows         =   26
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmPcxView.frx":1B24
   End
   Begin VB.Image imgBack 
      BorderStyle     =   1  '단일 고정
      Height          =   1050
      Index           =   2
      Left            =   10650
      Picture         =   "frmPcxView.frx":23B1
      Stretch         =   -1  'True
      Top             =   2850
      Width           =   1410
   End
   Begin VB.Image imgBack 
      BorderStyle     =   1  '단일 고정
      Height          =   1080
      Index           =   1
      Left            =   10650
      Picture         =   "frmPcxView.frx":76A3
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1410
   End
   Begin VB.Image imgBack 
      BorderStyle     =   1  '단일 고정
      Height          =   1080
      Index           =   0
      Left            =   10650
      Picture         =   "frmPcxView.frx":7C04
      Stretch         =   -1  'True
      Top             =   630
      Width           =   1410
   End
End
Attribute VB_Name = "frmPcxView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmPcxView.frm
'   작성자  : 오세원
'   내  용  : 전주예수병원 PCX 결과조회 (for MDB)
'   작성일  : 2007-06-13
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Dim AdoCn           As ADODB.Connection
Dim AdoRS           As ADODB.Recordset

Dim blnRS           As Boolean
Dim lngMaxCnt       As Long

Private WithEvents mobjPopups   As PopUpMessages
Attribute mobjPopups.VB_VarHelpID = -1
Private mobjDefault             As PopUpMessage

'   설정파일(PCX.INI) 내용 읽기
Private Function Get_PcxConfig(ByVal strConfigNm As String) As String
    Dim strFileName As String
    Dim strReturnedString As String

    strFileName = App.Path & "\pcx.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "PCX", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    Get_PcxConfig = strReturnedString
    
End Function

'   Access DB Connect
Public Function Set_DbConnect_Jet() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile  As String
    Dim strDestFile As String

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = Get_PcxConfig("MDBPath")   '   MDB Full Path & File Name
    UserName = Get_PcxConfig("MDBUser")  '   MDB User Name (Default = 'admin')
    Password = Get_PcxConfig("MDBPass")  '   MDB Pass Word

    If (DB_Name = "") Or (UserName = "") Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
        Exit Function
    End If
        
    With AdoCn
        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Mode").Value = adModeReadWrite
        .Properties("Persist Security Info").Value = False
        .Properties("Data Source").Value = DB_Name
        .Properties("User ID").Value = UserName
        .Properties("Jet OLEDB:Database Password").Value = Password
        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
        .Open
    End With

    Set_DbConnect_Jet = True
    
 Exit Function

ConnectError:
    '   오류처리
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn.State <> adStateOpen Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
    End If

End Function

Private Sub cmdEnd_Click()
    
    Set frmPcxView = Nothing
    
    End
    
End Sub

''   Not Use
'Private Sub cmdImgSet_Click()
'
'    If Me.Width = 12375 Then
'        Me.Width = 10305
'    Else
'        Me.Width = 12375
'    End If
'
'End Sub

''   Not Use
'Private Sub cmdInTaskBar_Click()
'
'    Call ShowInTaskBar(Me.hwnd, True)
'    Me.WindowState = vbMinimized
'
'End Sub

Private Sub cmdSearch_Click()
    
    '   Spread Initializing
    Call Set_SpreadInit(tblComplete)
    
    '   PCX Result Search
    Call Get_SearchList

End Sub


Private Sub Form_Load()
    Dim strPcxNames As Variant
    Dim intPcxCnt   As Integer
    
    If Not Set_DbConnect_Jet Then
        MsgBox "PCX 데이터베이스 경로를 확인하세요" & vbNewLine & vbNewLine & _
               "프로그램이 종료됩니다", vbOKOnly + vbCritical, Me.Caption
        End
    End If

    '   전역변수 초기화
    blnRS = False
    lngMaxCnt = 0
    
    '   Form Size
    Me.Width = 10305
    Me.Height = 8760
    
    '   Date Set
    dtpFrDt.Value = Now
    dtpToDt.Value = Now
    
    '   IIS Use PCX Machine ListUp
    strPcxNames = Get_PcxConfig("PCXNames")
    strPcxNames = Split(strPcxNames, "|")
    
    cboSpcPos.AddItem "ALL"
    
    For intPcxCnt = 0 To UBound(strPcxNames)
        cboSpcPos.AddItem strPcxNames(intPcxCnt)
    Next
        
    cboSpcPos.ListIndex = 0
    
    '   Result Search Interval Time Read & Set (1000 = 1 sec)
    tmrSearch.Interval = Get_PcxConfig("ViewInterval")
    tmrSearch.Enabled = True
        
    '   PopUp Window ShowDelay Time Read & Set (1000 = 1 sec)
    Set mobjPopups = New PopUpMessages
    mobjPopups.ShowDelay = Get_PcxConfig("PopupDelay")
    
    '   PopUp Window Default Set
    Call Set_DefaultPopup
    
    '   Spread Initializing
    Call Set_SpreadInit(tblComplete)
    
    '   PCX Result Search
    Call Get_SearchList
    
End Sub

'   Spread Initializing
Private Sub Set_SpreadInit(ByVal ClrSpread As Object)
    
    With ClrSpread
        .Col = 1
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .BlockMode = True
        .Action = 12    '## ActionClearText
        .BlockMode = False
    End With

End Sub

'   PCX Result Search & Spread ListUp
Private Sub Get_SearchList()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    '   Newest Result
    Set AdoRS = Get_NewResult
    
    If blnRS = False Then
        MsgBox "PCX 결과 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not AdoRS.BOF Then
        '   First Record Move
        AdoRS.MoveFirst
        
        If chkPopup.Value = "1" And lngMaxCnt < AdoRS.Fields("ITEMSEQ").Value Then
            lngMaxCnt = AdoRS.Fields("ITEMSEQ").Value
            
            '   PopUp Header Message
            strHMsg = ""
            strHMsg = strHMsg & AdoRS.Fields("NAME").Value & " " & _
                                AdoRS.Fields("SPCPOS").Value & " 결과등록"
                      
            '   PopUp Detail Message
            '   - Patient Age
            strAge = CLng(AdoRS.Fields("AGEDAY").Value) / 365
            If InStr(strAge, ".") > 0 Then strAge = Mid(strAge, 1, InStr(strAge, ".") - 1)
            
            strDMsg = vbNewLine
            strDMsg = strDMsg & "☞ 결  과 : " & AdoRS.Fields("RESULT").Value & " " & _
                                              AdoRS.Fields("HLDIV").Value & " " & _
                                              AdoRS.Fields("DPDIV").Value & vbNewLine & vbNewLine
            strDMsg = strDMsg & "◎ 바코드 : " & AdoRS.Fields("SPCNO").Value & vbNewLine
            strDMsg = strDMsg & "◎ 환자ID : " & AdoRS.Fields("PTID").Value & vbNewLine
            strDMsg = strDMsg & "◎ 이   름 : " & AdoRS.Fields("NAME").Value & "(" & _
                                              AdoRS.Fields("SEX").Value & "/" & strAge & ")" & vbNewLine
            strDMsg = strDMsg & "◎ 결과일시   " & vbNewLine & "     " & _
                                              Format(AdoRS.Fields("TRANSDT").Value, "####-##-##") & " " & _
                                              Mid(AdoRS.Fields("TRANSTM").Value, 1, 2) & ":" & Mid(AdoRS.Fields("TRANSTM").Value, 3, 2) & vbNewLine
            
            '   PopUp Windows Set & Show
            Call Set_AddPopup(strHMsg, strDMsg)
        End If
    End If
    
    Set AdoRS = Nothing
    blnRS = False
    
    '   Result List
    Set AdoRS = Get_ResultList

    If blnRS = False Then
        MsgBox "PCX 결과 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not AdoRS.BOF Then
        tblComplete.MaxRows = AdoRS.RecordCount
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
            With tblComplete
                If strTransDt <> AdoRS.Fields("TRANSDT").Value Then
                    .SetText 1, intRow, Format(AdoRS.Fields("TRANSDT").Value, "####-##-##")
                End If
                
                strTransDt = AdoRS.Fields("TRANSDT").Value
                strAge = CLng(AdoRS.Fields("AGEDAY").Value) / 365
                If InStr(strAge, ".") > 0 Then strAge = Mid(strAge, 1, InStr(strAge, ".") - 1)
                
                .SetText 2, intRow, Mid(AdoRS.Fields("TRANSTM").Value, 1, 2) & ":" & Mid(AdoRS.Fields("TRANSTM").Value, 3, 2)
                .SetText 3, intRow, AdoRS.Fields("SPCPOS").Value
                .SetText 4, intRow, AdoRS.Fields("SPCNO").Value
                .SetText 5, intRow, AdoRS.Fields("PTID").Value
                .SetText 6, intRow, AdoRS.Fields("NAME").Value
                .SetText 7, intRow, AdoRS.Fields("SEX").Value & "/" & strAge
                .SetText 8, intRow, AdoRS.Fields("INTNM").Value '& "/" & AdoRS.Fields("TESTCD").Value
                .SetText 9, intRow, AdoRS.Fields("RESULT").Value
                .SetText 10, intRow, AdoRS.Fields("HLDIV").Value
                .SetText 11, intRow, AdoRS.Fields("DPDIV").Value

                intRow = intRow + 1
                AdoRS.MoveNext
            End With
        Loop
    End If
    
    Set AdoRS = Nothing
    
End Sub

'   Result List Recordset
Public Function Get_ResultList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
             strSql = "SELECT a.*, b.intnm, b.testcd, b.result, b.hldiv, b.dpdiv "
    strSql = strSql & "  FROM ACC203 a, ACC204 b "
    strSql = strSql & " WHERE a.ITEMSEQ = b.ITEMSEQ "
    strSql = strSql & "   AND a.TRANSDT BETWEEN '" & Format(dtpFrDt.Value, "yyyymmdd") & "' AND '" & Format(dtpToDt.Value, "yyyymmdd") & "'"
    If Trim(cboSpcPos.Text) <> "ALL" Then
        strSql = strSql & "   AND a.SPCPOS = '" & cboSpcPos.Text & "'"
    End If
    strSql = strSql & " ORDER BY a.ITEMSEQ desc "

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList = AdoRS
        blnRS = True
    Else
        Set Get_ResultList = Nothing
        blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
    blnRS = False

End Function

'   Newest Result Recordset
Public Function Get_NewResult() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap
             strSql = "SELECT a.*, b.intnm, b.testcd, b.result, b.hldiv, b.dpdiv "
    strSql = strSql & "  FROM ACC203 a, ACC204 b "
    strSql = strSql & " WHERE a.ITEMSEQ = b.ITEMSEQ "
    strSql = strSql & "   AND a.TRANSDT = '" & Format(Now, "yyyymmdd") & "'"
    If Trim(cboSpcPos.Text) <> "ALL" Then
        strSql = strSql & "   AND a.SPCPOS = '" & cboSpcPos.Text & "'"
    End If
    strSql = strSql & " ORDER BY a.ITEMSEQ desc "

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_NewResult = AdoRS
        blnRS = True
    Else
        Set Get_NewResult = Nothing
        blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
    blnRS = False

End Function

'   Record Set Open
Public Function Get_Recordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                             ByVal AdoRS As ADODB.Recordset, _
                             Optional Call_Name As String, _
                             Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                             Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean

On Error GoTo DBOpenRsError
    
    With AdoRS
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    
    Get_Recordset = True

Exit Function

DBOpenRsError:
    Set AdoRS = Nothing
    Get_Recordset = False

End Function

'   timming Search
Private Sub tmrSearch_Timer()
    
    '   Spread Initializing
    Call Set_SpreadInit(tblComplete)
    
    '   PCX Result Search
    Call Get_SearchList
    
End Sub

'   PopUp Windows Set & Show
Private Sub Set_AddPopup(Optional ByVal strHMsg As String, Optional ByVal strDMsg As String)
    Dim objPopUp    As PopUpMessage
    
    Set objPopUp = New PopUpMessage
    
    With objPopUp
        .Caption = strHMsg
        .Message = strDMsg
        .BackColor = vbGreen
        .Clickable = False
        
        If optImg(0).Value Then
            Set .Background = imgBack.Item(0)
        ElseIf optImg(1).Value Then
            Set .Background = imgBack.Item(1)
        Else
            Set .Background = imgBack.Item(2)
        End If
    End With
    
    mobjPopups.Show objPopUp

End Sub

'   PopUp Windows Default Set
Private Sub Set_DefaultPopup()
    
    Set mobjDefault = New PopUpMessage
    
    With mobjDefault
        .ForeColor = vbWhite
        .Caption = "PCX Glucose"
        .Message = "PCX Glucose"
        .Clickable = True
        
        '   Background Image Set
        If optImg(0).Value Then
            Set .Background = imgBack.Item(0)
        ElseIf optImg(1).Value Then
            Set .Background = imgBack.Item(1)
        Else
            Set .Background = imgBack.Item(2)
        End If
        
        .ProgressBar = True
    End With

End Sub

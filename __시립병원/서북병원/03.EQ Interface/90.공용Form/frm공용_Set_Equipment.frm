VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frm공용_Set_Equipment 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Medical Equipment Info Setting"
   ClientHeight    =   5955
   ClientLeft      =   4785
   ClientTop       =   4470
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm공용_Set_Equipment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10695
   Begin FPSpread.vaSpread spr의료장비리스트 
      Height          =   4755
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   10575
      _Version        =   393216
      _ExtentX        =   18653
      _ExtentY        =   8387
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   23
      MaxRows         =   10
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm공용_Set_Equipment.frx":9F8A
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "선택(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9420
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   10620
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "EMR Interface Medical Equipment List"
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
      Index           =   9
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   3915
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
      Width           =   10575
   End
End
Attribute VB_Name = "frm공용_Set_Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function MM_CANCEL() As Boolean
    Call MM_KEY_CLEAR
End Function

Public Function MM_DELETE() As Boolean

End Function

Private Sub MM_INITIAL()
    Me.Height = 6435
    Me.Width = 10815
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call MM_CANCEL
End Sub

Public Function MM_INPUT() As Boolean

End Function

Private Sub MM_KEY_CLEAR()
    If spr의료장비리스트.MaxRows > 0 Then spr의료장비리스트.MaxRows = 0
End Sub

Public Function MM_PRINT() As Boolean

End Function

Public Function MM_SAVE() As Boolean
    MM_SAVE = False
    
    '/대상 의료장비 정보(레지스터) 광역변수 Setting
    gtypEQ_INFO.EQUIPCODE = GET_CELL(spr의료장비리스트, 1, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.EQUIPNM = GET_CELL(spr의료장비리스트, 2, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.EQUIPSEQ = GET_CELL(spr의료장비리스트, 3, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.DEPTCODE = GET_CELL(spr의료장비리스트, 4, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.EQUIPTYPE = Left(GET_CELL(spr의료장비리스트, 5, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.RECEIVETYPE = Left(GET_CELL(spr의료장비리스트, 6, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.EQUIPPORT = GET_CELL(spr의료장비리스트, 7, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.ORDYN = Left(GET_CELL(spr의료장비리스트, 8, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.QUERYTYPE = Left(GET_CELL(spr의료장비리스트, 9, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.ZIPYN = Left(GET_CELL(spr의료장비리스트, 10, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.ZIPNM = GET_CELL(spr의료장비리스트, 11, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALYN = Left(GET_CELL(spr의료장비리스트, 12, spr의료장비리스트.ActiveRow), 1)
    gtypEQ_INFO.SERIALPORT = GET_CELL(spr의료장비리스트, 13, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALBAUD = GET_CELL(spr의료장비리스트, 14, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALDATABIT = GET_CELL(spr의료장비리스트, 15, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALSTARTBIT = GET_CELL(spr의료장비리스트, 16, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALSTOPBIT = GET_CELL(spr의료장비리스트, 17, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALPARITY = GET_CELL(spr의료장비리스트, 18, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALRTS = GET_CELL(spr의료장비리스트, 19, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.SERIALDTR = GET_CELL(spr의료장비리스트, 20, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.EQIMGFILEPATH = GET_CELL(spr의료장비리스트, 21, spr의료장비리스트.ActiveRow)
    gtypEQ_INFO.FTPIMGFILEPATH = GET_CELL(spr의료장비리스트, 22, spr의료장비리스트.ActiveRow)
        
    MM_SAVE = True
End Function

Public Function MM_UPDATE() As Boolean

End Function

Public Function MM_VIEW() As Boolean
    Dim strEQCD             As String
    Dim strEQNM             As String
    Dim strEQSEQ            As String
    Dim strEQPOS            As String
    Dim strEQTYPE           As String
    Dim strRECEIVETYPE      As String
    Dim strEQUIPPORT        As String
    Dim strORDYN            As String
    Dim strQUERYTYPE        As String
    Dim strZIPYN            As String
    Dim strZIPNM            As String
    Dim strSERIALYN         As String
    Dim strSERIALPORT       As String
    Dim strSERIALBAUD       As String
    Dim strSERIALDATABIT    As String
    Dim strSERIALSTARTBIT   As String
    Dim strSERIALSTOPBIT    As String
    Dim strSERIALPARITY     As String
    Dim strSERIALRTS        As String
    Dim strSERIALDTR        As String
    Dim strEQIMGFILEPATH    As String
    Dim strFTPIMGFILEPATH   As String
    
    Dim strEQCD_Array
    Dim strEQNM_Array
    Dim strEQSEQ_Array
    Dim strEQPOS_Array
    Dim strEQTYPE_Array
    Dim strRECEIVETYPE_Array
    Dim strEQUIPPORT_Array
    Dim strORDYN_Array
    Dim strQUERYTYPE_Array
    Dim strZIPYN_Array
    Dim strZIPNM_Array
    Dim strSERIALYN_Array
    Dim strSERIALPORT_Array
    Dim strSERIALBAUD_Array
    Dim strSERIALDATABIT_Array
    Dim strSERIALSTARTBIT_Array
    Dim strSERIALSTOPBIT_Array
    Dim strSERIALPARITY_Array
    Dim strSERIALRTS_Array
    Dim strSERIALDTR_Array
    Dim strEQIMGFILEPATH_Array
    Dim strFTPIMGFILEPATH_Array
    
    Call MM_KEY_CLEAR
    
    '/대상 의료장비 정보(레지스터) 가져오기
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQNM)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)
    strEQPOS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQPOS)
    strEQTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQTYPE)
    strRECEIVETYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_RECEIVETYPE)
    strEQUIPPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQUIPPORT)
    strORDYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ORDYN)
    strQUERYTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_QUERYTYPE)
    strZIPYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPYN)
    strZIPNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPNM)
    strSERIALYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALYN)
    strSERIALPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPORT)
    strSERIALBAUD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALBAUD)
    strSERIALDATABIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDATABIT)
    strSERIALSTARTBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTARTBIT)
    strSERIALSTOPBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTOPBIT)
    strSERIALPARITY = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPARITY)
    strSERIALRTS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALRTS)
    strSERIALDTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDTR)
    strEQIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQIMGFILEPATH)
    strFTPIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_FTPIMGFILEPATH)
    
    strEQCD_Array = Split(strEQCD, ",")
    strEQNM_Array = Split(strEQNM, ",")
    strEQSEQ_Array = Split(strEQSEQ, ",")
    strEQPOS_Array = Split(strEQPOS, ",")
    strEQTYPE_Array = Split(strEQTYPE, ",")
    strRECEIVETYPE_Array = Split(strRECEIVETYPE, ",")
    strEQUIPPORT_Array = Split(strEQUIPPORT, ",")
    strORDYN_Array = Split(strORDYN, ",")
    strQUERYTYPE_Array = Split(strQUERYTYPE, ",")
    strZIPYN_Array = Split(strZIPYN, ",")
    strZIPNM_Array = Split(strZIPNM, ",")
    strSERIALYN_Array = Split(strSERIALYN, ",")
    strSERIALPORT_Array = Split(strSERIALPORT, ",")
    strSERIALBAUD_Array = Split(strSERIALBAUD, ",")
    strSERIALDATABIT_Array = Split(strSERIALDATABIT, ",")
    strSERIALSTARTBIT_Array = Split(strSERIALSTARTBIT, ",")
    strSERIALSTOPBIT_Array = Split(strSERIALSTOPBIT, ",")
    strSERIALPARITY_Array = Split(strSERIALPARITY, ",")
    strSERIALRTS_Array = Split(strSERIALRTS, ",")
    strSERIALDTR_Array = Split(strSERIALDTR, ",")
    strEQIMGFILEPATH_Array = Split(strEQIMGFILEPATH, ",")
    strFTPIMGFILEPATH_Array = Split(strFTPIMGFILEPATH, ",")
    
    If gstrJobMode = "1" Then '/표준모드면...
        If OpenDB(gstrREG_DB_CONSTR) = True Then
            With spr의료장비리스트
                For intX = 0 To UBound(strEQCD_Array)
                    gstrQuy = "SELECT A.*, B.EQUIPNM, B.EQUIPTYPE "
                    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF A INNER JOIN MM_EMR_EQUIP B ON A.EQUIPCODE = B.EQUIPCODE "
                    gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQUIPCODE = '" & strEQCD_Array(intX) & "' "
                    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPSEQ  =  " & Val(strEQSEQ_Array(intX)) & " "
                    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                        
                    If Not ADR Is Nothing Then
                        .MaxRows = .MaxRows + 1: .Row = .MaxRows
                    
                        .Col = 1: .Text = Trim(ADR!EQUIPCODE & "")  '/장비코드
                        .Col = 2: .Text = Trim(ADR!EQUIPNM & "")    '/장비명
                        .Col = 3: .Text = Trim(ADR!EQUIPSEQ & "")   '/장비SEQ
                        .Col = 4: .Text = Trim(ADR!DEPTCODE & "")   '/DEPTCODE(설치장소)
                        .Col = 5:                                   '/장비결과방식(1.SM, 2.VPM, 3.ICM)
                        Select Case Trim(ADR!EQUIPTYPE & "")
                            Case "1":   .Text = "1.SM"
                            Case "2":   .Text = "2.VPM"
                            Case "3":   .Text = "3.ICM"
                        End Select
                        .Col = 6:                                   '/Image획득방식(1.직접, 2.간접)
                        Select Case Trim(ADR!RECEIVETYPE & "")
                            Case "1":   .Text = "1.직접"
                            Case "2":   .Text = "2.간접"
                        End Select
                        .Col = 7: .Text = Trim(ADR!EQUIPPORT & "")  '/장비PC접속포트(RECEIVETYPE 간접일경우)
                        .Col = 8:                                   '/처방유무(Y.처방, N.미처방)
                        Select Case Trim(ADR!ORDYN & "")
                            Case "Y":   .Text = "Y.처방"
                            Case "N":   .Text = "N.미처방"
                        End Select
                        .Col = 9:                                  '/처방쿼리종류(처방유무가 Y일 때:1.과접수QUARY, 2.처방QUARY, 3.종합검진)
                        Select Case Trim(ADR!QUERYTYPE & "")
                            Case "1":   .Text = "1.과접수QUARY"
                            Case "2":   .Text = "2.처방QUARY"
                            Case "3":   .Text = "3.종합검진"
                        End Select
                        .Col = 10:                                  '/ZanImagePrinter사용여부(Y.사용,N.미사용)
                        Select Case Trim(ADR!ZIPYN & "")
                            Case "Y":   .Text = "Y.사용"
                            Case "N":   .Text = "N.미사용"
                        End Select
                        .Col = 11: .Text = Trim(ADR!ZIPNM & "")     '/ZanImagePrinter Device Name
                        .Col = 12:                                  '/RS232 SERIAL 사용여부(Y.사용, N.미사용)
                        Select Case Trim(ADR!SERIALYN & "")
                            Case "Y":   .Text = "Y.사용"
                            Case "N":   .Text = "N.미사용"
                        End Select
                        .Col = 13: .Text = Trim(ADR!SERIALPORT & "")        '/RS232 SERIAL PORT
                        .Col = 14: .Text = Trim(ADR!SERIALBAUD & "")        '/RS232 통신속도
                        .Col = 15: .Text = Trim(ADR!SERIALDATABIT & "")     '/RS232 DATABIT(7,8)
                        .Col = 16: .Text = Trim(ADR!SERIALSTARTBIT & "")    '/RS232 STARTBIT(1,2)
                        .Col = 17: .Text = Trim(ADR!SERIALSTOPBIT & "")     '/RS232 STOPBIT(1,2)
                        .Col = 18: .Text = Trim(ADR!SERIALPARITY & "")      '/RS232 PARITY(E,N,O)
                        .Col = 19: .Text = Trim(ADR!SERIALRTS & "")         '/RS232 RTS(0,1)
                        .Col = 20: .Text = Trim(ADR!SERIALDTR & "")         '/RS232 DTR(0,1)
                        .Col = 21: .Text = Trim(ADR!EQIMGFILEPATH & "")     '/장비결과파일저장경로(Client)
                        .Col = 22: .Text = Trim(ADR!FTPIMGFILEPATH & "")    '/FTP결과파일저장경로(Client)
                        .Col = 23: .Text = Trim(ADR!REMARK & "")            '/비고
                        
                        If .MaxTextRowHeight(.Row) > 13.3 Then .RowHeight(.Row) = .MaxTextRowHeight(.Row)
                        
                        ADR.Close: Set ADR = Nothing
                    End If
                Next intX
            End With
            
            Call CloseDB
        End If
    Else '/임시모드면...
        With spr의료장비리스트
            For intX = 0 To UBound(strEQCD_Array)
                .MaxRows = .MaxRows + 1: .Row = .MaxRows
            
                On Error Resume Next
                
                .Col = 1:   .Text = strEQCD_Array(intX)
                .Col = 2:   .Text = strEQNM_Array(intX)
                .Col = 3:   .Text = strEQSEQ_Array(intX)
                .Col = 4:   .Text = strEQPOS_Array(intX)
                .Col = 5:   .Text = strEQTYPE_Array(intX)
                .Col = 6:   .Text = strRECEIVETYPE_Array(intX)
                .Col = 7:   .Text = strEQUIPPORT_Array(intX)
                .Col = 8:   .Text = strORDYN_Array(intX)
                .Col = 9:   .Text = strQUERYTYPE_Array(intX)
                .Col = 10:  .Text = strZIPYN_Array(intX)
                .Col = 11:  .Text = strZIPNM_Array(intX)
                .Col = 12:  .Text = strSERIALYN_Array(intX)
                .Col = 13:  .Text = strSERIALPORT_Array(intX)
                .Col = 14:  .Text = strSERIALBAUD_Array(intX)
                .Col = 15:  .Text = strSERIALDATABIT_Array(intX)
                .Col = 16:  .Text = strSERIALSTARTBIT_Array(intX)
                .Col = 17:  .Text = strSERIALSTOPBIT_Array(intX)
                .Col = 18:  .Text = strSERIALPARITY_Array(intX)
                .Col = 19:  .Text = strSERIALRTS_Array(intX)
                .Col = 20:  .Text = strSERIALDTR_Array(intX)
                .Col = 21:  .Text = strEQIMGFILEPATH_Array(intX)
                .Col = 22:  .Text = strFTPIMGFILEPATH_Array(intX)
                
                On Error GoTo 0
                
                If .MaxTextRowHeight(.Row) > 13.3 Then .RowHeight(.Row) = .MaxTextRowHeight(.Row)
            Next intX
        End With
    End If
End Function

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    If spr의료장비리스트.ActiveRow < 1 Then
        MsgBox "선택한 의료장비가 없습니다!", vbCritical, "선택불가": Exit Sub
    End If
    If MM_SAVE = True Then Unload Me
End Sub

Private Sub Form_Load()
    Call MM_INITIAL
    
    Call MM_VIEW
End Sub

Private Sub spr의료장비리스트_DblClick(ByVal Col As Long, ByVal Row As Long)
    If spr의료장비리스트.Row < 1 Then
        MsgBox "선택한 의료장비가 없습니다!", vbCritical, "선택불가": Exit Sub
    End If
    If MM_SAVE = True Then Unload Me
End Sub

Private Sub spr의료장비리스트_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If spr의료장비리스트.ActiveRow < 1 Then
            MsgBox "선택한 의료장비가 없습니다!", vbCritical, "선택불가": Exit Sub
        End If
        If MM_SAVE = True Then Unload Me
    End If
End Sub

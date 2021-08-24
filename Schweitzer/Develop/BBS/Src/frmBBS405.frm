VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmBBS405 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈자 접수내역 조회"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS405.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14700
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin MSCommLib.MSComm MyComm 
      Left            =   -300
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   915
      Left            =   75
      TabIndex        =   6
      Top             =   -45
      Width           =   14400
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   180
         Left            =   12660
         TabIndex        =   13
         Top             =   555
         Width           =   780
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00D1D6DC&
         Caption         =   "Pheresis"
         Height          =   510
         Index           =   3
         Left            =   10005
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   240
         Width           =   1320
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00D1D6DC&
         Caption         =   "지정 헌혈"
         Height          =   510
         Index           =   1
         Left            =   7365
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   240
         Width           =   1320
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00D1D6DC&
         Caption         =   "임의 헌혈"
         Height          =   510
         Index           =   0
         Left            =   6030
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00D1D6DC&
         Caption         =   "Autologous"
         Height          =   510
         Index           =   2
         Left            =   8685
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   240
         Width           =   1320
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00D1D6DC&
         Caption         =   "Phlebotomy"
         Height          =   510
         Index           =   4
         Left            =   11325
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   240
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   330
         Left            =   1215
         TabIndex        =   0
         Top             =   330
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36943
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   345
         Left            =   2730
         TabIndex        =   1
         Top             =   330
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36943
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "접수일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   495
         Index           =   15
         Left            =   4965
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   255
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
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
         Caption         =   "헌혈 종류"
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2565
         TabIndex        =   7
         Top             =   405
         Width           =   135
      End
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7545
      Left            =   75
      TabIndex        =   5
      Tag             =   "10114"
      Top             =   885
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   13309
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   25
      MaxRows         =   29
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS405.frx":076A
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   29
   End
End
Attribute VB_Name = "frmBBS405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    TcName = 1
    tcDOB
    TcSEXAGE
    tcABO
    tcACCDT
    
    TcTMPID
    TcDONORTYPE
    TcSELPTID
    TcBLOODNO
    TcCOMP
    TcVOLUMN
    
    TcACCVAL
    TcRMKVAL
    TcTESTVAL
    TcCANCEL
    tcDONORID
End Enum

Private Sub chkAll_Click()
    optDonorCd(0).Enabled = IIf(chkAll.value = 0, True, False)
    optDonorCd(1).Enabled = IIf(chkAll.value = 0, True, False)
    optDonorCd(2).Enabled = IIf(chkAll.value = 0, True, False)
    optDonorCd(3).Enabled = IIf(chkAll.value = 0, True, False)
    optDonorCd(4).Enabled = IIf(chkAll.value = 0, True, False)
    
End Sub

Private Sub cmdClear_Click()
    medClearTable tblList
    tblList.MaxRows = 30
    dtpFrDt.value = DateAdd("d", -7, GetSystemDate)
    dtpToDt.value = GetSystemDate
    chkAll.value = 1

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS405 = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim objGetSql As clsGetSqlStatement
    Dim objProBar As clsProgress
    Dim Rs        As Recordset
    Dim BRs       As Recordset
    Dim strTmp    As String
    Dim FrDt      As String
    Dim ToDt      As String
    Dim ii        As Integer
    Dim strBldNo  As String
    Dim PtId      As String
    
    
    FrDt = Format(dtpFrDt.value, PRESENTDATE_FORMAT)
    ToDt = Format(dtpToDt.value, PRESENTDATE_FORMAT)
    Set objGetSql = New clsGetSqlStatement
    Set Rs = objGetSql.Get_DonorQuery(FrDt, ToDt)
    
    If Rs.RecordCount > 0 Then
        Set objProBar = New clsProgress
'        Set objProBar.StatusBar = medMain.stsBar
        objProBar.Container = MainFrm.stsBar
        objProBar.Max = Rs.RecordCount
        With tblList
            ii = 1
            .MaxRows = 0
            .ReDraw = False
            Rs.MoveFirst
            Do Until Rs.EOF
                If chkAll.value = 0 Then
                    If Not optDonorCd(Val("" & Rs.Fields("donorcd").value & "")).value Then
                        GoTo Skip
                    End If
                End If
                .MaxRows = .DataRowCnt + 1
                .Row = .MaxRows
                .Col = 25: .CellType = CellTypeStaticText
                           .Text = ""
                If strTmp <> Rs.Fields("donorid").value & "" Then
                    .Col = TblColumn.TcName:   .value = Rs.Fields("donornm").value & ""
                    .Col = TblColumn.tcDOB:    .value = Format(Rs.Fields("dob").value & "", "####-##-##")
                    .Col = TblColumn.TcSEXAGE: .value = Rs.Fields("sex").value & "" & "/"
                                               If Trim(Rs.Fields("dob").value & "") <> "" Then
                                                   .value = .value & medFindAge(Rs.Fields("dob").value & "", "Y")
                                               End If
                    .Col = TblColumn.tcABO: .value = Rs.Fields("abo").value & "" & Rs.Fields("rh").value & ""
                End If
                
                .Col = TblColumn.tcACCDT: .value = Format(Rs.Fields("donoraccdt").value & "", "####/##/##")
                .Col = TblColumn.TcTMPID: .value = Rs.Fields("tmpid").value & ""
                
                .Col = TblColumn.TcACCVAL: .value = IIf(Rs.Fields("okdiv1").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv1").value & "" = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                .Col = TblColumn.TcRMKVAL: .value = IIf(Rs.Fields("okdiv2").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv2").value & "" = "0", "Not", "")):  .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                .Col = TblColumn.TcTESTVAL: .value = IIf(Rs.Fields("okdiv3").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv3").value & "" = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                
                Set BRs = objGetSql.Get_DonorBlood(Rs.Fields("donorid").value & "", Rs.Fields("donoraccdt").value & "")
                If Not BRs.EOF Then
                    Do Until BRs.EOF
                        .Col = TblColumn.TcDONORTYPE
                            Select Case Rs.Fields("donorcd").value & ""
                                Case "0": .value = "임의헌혈"
                                Case "1": .value = "지정헌혈"
                                        If BRs.Fields("reserved").value & "" <> "1" Then
                                            .value = "지정취소"
                                        Else
                                            If Rs.Fields("reservedid").value & "" <> "" And Rs.Fields("reservedid").value & "" <> "0" Then
                                                PtId = GetPtNm(Rs.Fields("reservedid").value & "")
                                                If PtId <> "" Then
                                                    .value = PtId & "(" & Rs.Fields("reservedid").value & "" & ")"
                                                End If
                                            Else
                                                .value = ""
                                            End If
                                        End If
                                Case "2": .value = "Autologos"
                                Case "3": .value = "Pheresis"
                                        If BRs.Fields("pherefg").value & "" <> "1" Then
                                            .value = "Pheresis취소"
                                        Else
                                            If Rs.Fields("reservedid").value & "" <> "" And Rs.Fields("reservedid").value & "" <> "0" Then
                                                PtId = GetPtNm(Rs.Fields("reservedid").value & "")
                                                If PtId <> "" Then
                                                    .value = PtId & "(" & Rs.Fields("reservedid").value & "" & ")"
                                                End If
                                            Else
                                                .value = ""
                                            End If
                                        End If
                                Case "4": .value = "Phlebotomy"
                            End Select
                            strBldNo = BRs.Fields("bldsrc").value & "" & "-" & BRs.Fields("bldyy").value & "" & "-" & Format(BRs.Fields("bldno").value & "", "000000")
                            If strBldNo = "--000000" Then strBldNo = ""
                            .Col = TblColumn.TcBLOODNO: .value = strBldNo
                            .Col = TblColumn.TcCOMP: .value = BRs.Fields("abbrnm").value & ""
                            .Col = TblColumn.TcVOLUMN: .value = Rs.Fields("volumn").value & ""
                            .Col = TblColumn.TcCANCEL: .value = IIf(Rs.Fields("cancelfg").value & "" = "1", "Y", "")
                            
                            .Col = TblColumn.TcACCVAL: .value = IIf(Rs.Fields("okdiv1").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv1").value & "" = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                            .Col = TblColumn.TcRMKVAL: .value = IIf(Rs.Fields("okdiv2").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv2").value & "" = "0", "Not", "")):  .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                            .Col = TblColumn.TcTESTVAL: .value = IIf(Rs.Fields("okdiv3").value & "" = "1", "Ok", IIf(Rs.Fields("okdiv3").value & "" = "0", "Not", "")): .ForeColor = IIf(.value = "Ok", vbBlack, vbRed)
                            
                            If strBldNo <> "" Then
                                .Col = 25: .CellType = CellTypeButton
                                           .TypeButtonText = "Bar"
                            Else
                                .Col = 25: .CellType = CellTypeStaticText
                                           .Text = ""
                            End If
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            BRs.MoveNext
                        
                        Loop
                        .MaxRows = .MaxRows - 1
                
                End If
                Set BRs = Nothing
                
                .Col = TblColumn.tcDONORID: .value = Rs.Fields("donorid").value & ""
                .Col = TblColumn.TcSELPTID
                strTmp = .value
                objProBar.value = ii
                ii = ii + 1
Skip:
                Rs.MoveNext
            Loop
            .ReDraw = True
        End With
    
    Else
        MsgBox "해당 기간엔 헌혈자내역이 없습니다.", vbInformation, "헌혈자조회"
    End If
    
    Set Rs = Nothing
    Set objGetSql = Nothing
End Sub

Private Sub Form_Load()
    Call cmdClear_Click
End Sub


Private Sub tblList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim Label_String As String
    Dim strBldNoP    As String
    Dim strBldNo     As String
    
    tblList.Row = Row
    tblList.Col = TblColumn.TcBLOODNO
    
    
    strBldNoP = tblList.value
    
    strBldNo = AddCheckDigit(Replace(tblList.value, "-", ""))
    
    Label_String = "\1B@z" & vbCrLf & _
                   "\1B@f09" & vbCrLf & _
                   "\1Ba0901840208" & vbCrLf & _
                   "\1Bf09" & vbCrLf & _
                   "\1Bbs090200018000451200800211000" & vbCrLf & _
                   "\1Bds09040002000009122200010" & vbCrLf & _
                   "\1Bbw0902" & strBldNo & vbCrLf & _
                   "\1Bdw0904" & strBldNoP & vbCrLf & _
                   "\1Bq0003"

    Call Label_PrintOut(Label_String)
   
End Sub

Public Sub Label_PrintOut(ByVal barString As String)
   
   Dim FileNo As Long
   'Dim MyComm As Object
   Dim PkSize As Integer

   On Error GoTo Skip
   
   'Set MyComm = frmComm.MSComm1
      
    PkSize = 250
   
    If MyComm.PortOpen Then Exit Sub
    MyComm.CommPort = 2
    MyComm.Settings = "9600,N,8,1"
    MyComm.InputLen = 8192
    
    If Not MyComm.PortOpen Then MyComm.PortOpen = True
    
    If Len(barString) > PkSize Then
        MyComm.Output = Mid(barString, 1, PkSize)
        While (Len(barString)) > PkSize
               barString = Mid(barString, PkSize + 1)
               MyComm.Output = Mid(barString, 1, PkSize)
        Wend
    Else
        MyComm.Output = barString
    End If
    If MyComm.PortOpen Then MyComm.PortOpen = False
   
Skip:
   'Call Clear
   'Set medMain.MyComm = Nothing

End Sub
Private Function AddCheckDigit(sBarcode As String) As String
    Dim iLen%
    Dim i%
    Dim iCheckSum%
    Dim iA%, iB%, iC%, id%
    iLen = Len(sBarcode)
    iCheckSum = 0
    iA = 0
    iB = 0
    For i = 1 To iLen
        If i Mod 2 = 1 Then
            iB = iB + Val(Mid(sBarcode, i, 1))
        Else
            iA = iA + Val(Mid(sBarcode, i, 1))
        End If
    Next
    If iLen Mod 2 = 1 Then
        iC = iB * 3 + iA
    Else
        iC = iB + iA * 3
    End If
    id = iC Mod 10
    If id = 0 Then
        iCheckSum = 0
    Else
        iCheckSum = 10 - id
    End If
    
    AddCheckDigit = sBarcode & Trim(Str(iCheckSum))
End Function


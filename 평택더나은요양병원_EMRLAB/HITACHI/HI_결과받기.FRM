VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmHi_결과받기 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hitachi 747  결과받기"
   ClientHeight    =   7170
   ClientLeft      =   1965
   ClientTop       =   1185
   ClientWidth     =   11580
   Icon            =   "Hi_결과받기.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame sFrame2 
      Height          =   6165
      Left            =   5190
      TabIndex        =   11
      Top             =   30
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   10874
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdWork 
         Height          =   4695
         Left            =   270
         OleObjectBlob   =   "Hi_결과받기.frx":08CA
         TabIndex        =   10
         Top             =   1350
         Width           =   5715
      End
      Begin Threed.SSPanel pnlStatus 
         Height          =   945
         Left            =   300
         TabIndex        =   21
         Top             =   270
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   1667
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Font3D          =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1035
         Left            =   1860
         TabIndex        =   15
         Top             =   210
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   1826
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelInner      =   1
         Begin Threed.SSPanel pnlSample 
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   120
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1244
            _ExtentY        =   661
            _StockProps     =   15
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.74
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   180
            TabIndex        =   17
            Top             =   120
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Sample No"
            ForeColor       =   16711680
            BackColor       =   16744703
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.74
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlPos 
            Height          =   375
            Left            =   3390
            TabIndex        =   18
            Top             =   120
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.74
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   375
            Left            =   2490
            TabIndex        =   19
            Top             =   120
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Pos No"
            ForeColor       =   12582912
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.74
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   405
            Left            =   180
            TabIndex        =   20
            Top             =   540
            Width           =   3735
            _Version        =   65536
            _ExtentX        =   6588
            _ExtentY        =   714
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.74
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin Threed.SSFrame sFrame1 
      Height          =   6975
      Left            =   180
      TabIndex        =   7
      Top             =   30
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   12303
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdList 
         Height          =   4945
         Left            =   420
         OleObjectBlob   =   "Hi_결과받기.frx":0C31
         TabIndex        =   8
         Top             =   1350
         Width           =   3945
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1035
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   1826
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelInner      =   1
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboSelect 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   555
            Width           =   1605
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   345
            Left            =   1530
            TabIndex        =   0
            Top             =   135
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   327680
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-mm-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   390
            Left            =   240
            TabIndex        =   13
            Top             =   120
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "접수일자"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   390
            Left            =   240
            TabIndex        =   14
            Top             =   510
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   688
            _StockProps     =   15
            Caption         =   "조회조건"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            RoundedCorners  =   0   'False
            MouseIcon       =   "Hi_결과받기.frx":1CBE
         End
         Begin Threed.SSCommand cmdQuery 
            Height          =   780
            Left            =   3420
            TabIndex        =   2
            Top             =   120
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   1376
            _StockProps     =   78
            Caption         =   "조 회"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            RoundedCorners  =   0   'False
            Picture         =   "Hi_결과받기.frx":2598
         End
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   495
         Left            =   420
         TabIndex        =   9
         Top             =   6360
         Width           =   3960
         _Version        =   65536
         _ExtentX        =   6985
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Work List 등록"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         Picture         =   "Hi_결과받기.frx":2E72
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4860
      Top             =   3030
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   405
      Left            =   4830
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "SSCommand1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin Threed.SSCommand cmdResult 
      Height          =   825
      Left            =   5370
      TabIndex        =   4
      Top             =   6270
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "결과 수신"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   1
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과받기.frx":2E8E
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   825
      Left            =   9810
      TabIndex        =   6
      Top             =   6270
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "종   료"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   1
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과받기.frx":32E0
   End
   Begin Threed.SSCommand cmdServer 
      Height          =   825
      Left            =   6855
      TabIndex        =   5
      Top             =   6270
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "결과 등록"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   1
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과받기.frx":3BBA
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   825
      Left            =   8340
      TabIndex        =   22
      Top             =   6270
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "화면초기화"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   1
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과받기.frx":4494
   End
End
Attribute VB_Name = "frmHi_결과받기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
    Dim Rcvbuffer  As String
    Dim TablePtr   As Integer
    
    Dim CheckSumCount As Integer
    Dim EtxFlag As Integer
    Dim tData() As String
        
    '----- BarCode Lab_ID 편집 저장
    Dim labdate As String
    Dim SlipCd  As String
    Dim LabNo   As String
        
    Dim TimeCount   As Integer
    
    Dim OldRow  As Integer
    Dim spdRow  As Integer      'spdWork Row
Private Sub Add_Db_Sample()
        
    Dim wk1     As String
    Dim wk2     As Long
        
    If RecordExist(TbSample, "PrimaryKey", TestHeaderTable.SampNo) Then
        TbSample.Edit
    Else
        TbSample.AddNew
        TbSample!SampNo = TestHeaderTable.SampNo ' Sample Number
    End If
    
    TbSample!RackNo = TestHeaderTable.RackNo    ' Rack Number
    TbSample!PosNo = TestHeaderTable.PosiNo    ' Position Number
    TbSample!Lab_ID = TestHeaderTable.Lab_ID
    
    TbSample.Update
    TbSample.MoveLast
    
End Sub


Private Sub Add_Db_Result(SampNo As String, EqCode As String, Result As String, ix1 As Integer)

    If RecordExists(TbResult, "PrimaryKey", SampNo, EqCode) Then
        TbResult.Edit
    Else
        TbResult.AddNew
        TbResult!SampNo = SampNo
        TbResult!EqCode = EqCode
    End If
    TbResult!SeqNo = Format(ix1)
    TbResult!Result = Result

    TbResult.Update
    TbResult.MoveLast

End Sub

' *=====================================================*
' *   Data Display & Data Base Handling                 *
' *=====================================================*
Private Sub DataDisplay()

    Dim SpreadPtr As Integer
    Dim CodeVal   As Integer
    Dim ix1       As Integer
    Dim Tmp     As Variant
        
    If Val(TestHeaderTable.SampNo) = 0 Then
        Exit Sub
    End If

'    TestHeaderTable.SampNo = Format(Val(TestHeaderTable.SampNo), "00000")
    
    pnlSample = TestHeaderTable.SampNo      ' Sample Number
    pnlPos = TestHeaderTable.PosiNo         ' Position Number

    With spdWork
        Call .GetText(1, spdRow, Tmp)
        TestHeaderTable.Lab_ID = Left(Tmp, 8) & Mid(Tmp, 10, 1) & Right(Tmp, 5)     'Lab_ID
        
        Call .SetText(3, spdRow, Trim(TestHeaderTable.SampNo))
        Call .SetText(4, spdRow, Trim(TestHeaderTable.PosiNo))
        spdRow = spdRow + 1
    End With
    '--- Sample 정보 Hi_Sample Table에 입력
    Call Add_Db_Sample
    '--------------------------------------
    
    For ix1 = 1 To MAX_NUM
'        CodeVal = Val(Trim(TestResultTable(ix1).SeqNo))
'
'        If Trim(TestNameTable(CodeVal).Code) <> "" Then
        If Trim(TestResultTable(ix1).TestNo) <> "" Then
            '--- 검사 결과 입력
            Call Add_Db_Result(TestHeaderTable.SampNo, TestResultTable(ix1).TestNo, _
                            TestResultTable(ix1).Result, ix1)   'Result Insert
            '------------------
        End If
    Next ix1
    
End Sub

' *=====================================================*
' *         Data편집 & 응답처리                         *
' *=====================================================*
Private Sub DataEditResponse()
    
    Dim Frame    As String              ' FRAME
    Dim Func     As String              ' FUNCTION
    Dim RackNo   As String              ' Rack No.
    Dim SampNo   As String              ' Sample No.
    Dim PosiNo   As String              ' Posotion No.
    Dim Id_Num   As String              ' ID No.
    Dim SendBuf  As String              ' Send Buffer
    Dim ChekSum  As Integer             ' Check Sum
    Dim ChkSumS  As String              ' Check Sum
    Dim ResData  As String              ' Result Data Buffer
    Dim TstCount As Integer
    Dim StartPtr As Integer
    Dim i        As Long
    Dim wk1      As String
    Dim wk2      As Long
    Dim BarCode  As String
    Dim SampIf  As String
    
    Dim yk_temp  As String
    Dim sSql    As String
    Dim SnapData    As Snapshot
        
    Frame = Left$(Rcvbuffer, 1)        ' get Frame of RcvBuffer.
    
    Select Case Frame
        Case ">"                        ' ANY
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            Comm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until Comm1.OutBufferCount = 0
           
        Case "?"                        ' REP
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            SendBuf = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Comm1.Output = SendBuf
            Do
            '   DoEvents
            Loop Until Comm1.OutBufferCount = 0
       
        Case ";"                        ' SPE
            Func = Mid$(Rcvbuffer, 2, 1)     ' Function
            SampNo = Mid$(Rcvbuffer, 4, 3)   ' Sample No.
            RackNo = Mid$(Rcvbuffer, 7, 1)   ' Rack No.
            PosiNo = Mid$(Rcvbuffer, 8, 2)  ' Position No.
            Id_Num = Mid$(Rcvbuffer, 10, 13)  ' Id No.
            SampIf = Mid$(Rcvbuffer, 4, 34)
                        
        Case "1" To "9"               ' FR1 to FR9
            For i = 1 To 10000        ' loop
            Next i
            '--- 11/25 yk (검사항목 25개 이상일 경우 처리)
            Func = Mid$(Rcvbuffer, 2, 1)     ' Function
            If Func = "K" Or Func = "L" Then
               GoTo Nxt2
            End If
            If Func <> "@" And Func <> "N" And Func <> "M" Then        ' TNNA
                With TestHeaderTable
                    .RackNo = Trim$(Mid$(Rcvbuffer, 7, 1))  ' Rack No.
                    .SampNo = Right$("0000" & Trim$(Mid$(Rcvbuffer, 4, 3)), 5) ' Sample No.
                    .PosiNo = Trim$(Mid$(Rcvbuffer, 8, 2))  ' Position No.
                End With
            End If
            '=============================================
                        
            Func = Mid$(Rcvbuffer, 2, 1)     ' Function
            If Func <> "@" Then              ' TNNA ?
                TstCount = Val(Trim$(Mid$(Rcvbuffer, 38, 2))) ' Test Count
                StartPtr = 40
                For i = 1 To TstCount
                    ResData = Mid$(Rcvbuffer, StartPtr, 9)
                    TestResultTable(TablePtr).TestNo = Right("0" & Trim(Mid$(ResData, 1, 2)), 2) ' Test Number
                    TestResultTable(TablePtr).Result = Mid$(ResData, 3, 6) ' Result
                    TablePtr = TablePtr + 1   ' Table Pointer Increament
                    StartPtr = StartPtr + 9   ' Start Pointer Increament
                Next i
            End If
            ' MOR Send
            Comm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until Comm1.OutBufferCount = 0
    
       Case ":"                        ' END
           Func = Mid$(Rcvbuffer, 2, 1)     ' Function
           If Func = "K" Or Func = "L" Then
              GoTo Nxt2
           End If
    Rem modify  96.02.15
           Dim tmpid    As String
           If Func <> "@" And Func <> "N" And Func <> "M" Then              ' TNNA
                With TestHeaderTable
                    .RackNo = Trim$(Mid$(Rcvbuffer, 7, 1))  ' Rack No.
                    .SampNo = Right$("0000" & Trim$(Mid$(Rcvbuffer, 4, 3)), 5) ' Sample No.
                    .PosiNo = Trim$(Mid$(Rcvbuffer, 8, 2))  ' Position No.
                End With
                tmpid = Trim$(Mid$(Rcvbuffer, 4, 3))
                If Len(tmpid) = 0 Then
                   GoTo Nxt2
                End If
                
                TstCount = Val(Trim$(Mid$(Rcvbuffer, 38, 2))) ' Test Count
                StartPtr = 40
                For i = 1 To TstCount
                    ResData = Mid$(Rcvbuffer, StartPtr, 9)
                    
                    TestResultTable(TablePtr).TestNo = Right$("0" & Trim(Mid$(ResData, 1, 2)), 2) ' Test Number
                    TestResultTable(TablePtr).Result = Mid$(ResData, 3, 6) ' Result
                    TablePtr = TablePtr + 1   ' Table Pointer Increament
                    StartPtr = StartPtr + 9   ' Start Pointer Increament
                Next i
                                
                '--- Spread에 표시/Hi_Sample 입력
                Call DataDisplay
                                
                TablePtr = 0
           End If
           
Nxt2:
           ' MOR Send
           Comm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
           Do
           '   DoEvents
           Loop Until Comm1.OutBufferCount = 0
    End Select

End Sub
'
'   경북대 7060
'
Public Sub DataEditResponse2()
'
'   Dim Frame    As String              ' FRAME
'   Dim Func     As String              ' FUNCTION
'   Dim RackNo   As String              ' Rack No.
'   Dim SampNo   As String              ' Sample No.
'   Dim PosiNo   As String              ' Posotion No.
'   Dim Id_Num   As String              ' ID No.
'   Dim SampIf   As String              ' Sample Information (Add 96.11.28)
'   Dim SendBuf  As String              ' Send Buffer
'   Dim ChekSum  As Integer             ' Check Sum
'   Dim ChkSumS  As String              ' Check Sum
'   Dim ResData  As String              ' Result Data Buffer
'   Dim TstCount As Integer
'   Dim StartPtr As Integer
'   Dim i        As Long
'   Dim wk1      As String
'   Dim wk2      As Long
'   Dim BarCode  As String
'   Dim Ptr1     As Integer
'   Dim Ptr2     As Integer
'   Dim Ptr3     As Integer
'   Dim Ptr4     As Integer
'   Dim Ptr5     As Integer
'   Dim WkResult As Single
'
'   Frame = Left$(Rcvbuffer, 1)        ' get Frame of RcvBuffer.
'   'Print #1, Rcvbuffer
'   Select Case Frame
'      Case ">"                        ' ANY
'         For i = 1 To 10000          ' loop
'         Next i
'         ' MOR Send
'         MSComm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
'         Do
'         '   DoEvents
'         Loop Until MSComm1.OutBufferCount = 0
'      Case "?"                        ' REP
'         For i = 1 To 10000          ' loop
'         Next i
'
'         ' MOR Send
'         SendBuf = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
'         MSComm1.Output = SendBuf
'         Do
'         '   DoEvents
'         Loop Until MSComm1.OutBufferCount = 0
'      Case ";"                        ' SPE
'         'For i = 1 To 10000          ' loop
'         'Next i
'
'         Func = Mid$(Rcvbuffer, 2, 1)     ' Function
'         SampNo = Mid$(Rcvbuffer, 4, 3)   ' Sample No.
'         RackNo = Mid$(Rcvbuffer, 7, 1)   ' Rack No.
'         PosiNo = Mid$(Rcvbuffer, 8, 2)  ' Position No.
'         Id_Num = Mid$(Rcvbuffer, 10, 13)  ' Id No.
'         SampIf = Mid$(Rcvbuffer, 4, 34)   ' Sample Information (Add 96.11.28)
'         TestDat = String$(60, "0")
'         Rem BarCodeTable 읽기
'         Dim tmpseqno$, tmppid$, tmpbc$, sStr$
'         Dim tmpsymd$, tmplno$
'         Dim s작업일자$, sPart코드$, sSlip부분코드$, s작업번호$
'         Dim s검사항목코드$
'         tmpbc = Trim(Id_Num)
'      Rem test remove
'      '   tmpbc = "970716C20001"
'      Rem test remove
'
'         's작업일자 = "19" & Mid(tmpbc, 1, 6)
'         s작업일자 = H0SUB_DATE_6TO8(Mid(tmpbc, 1, 6), "1")
'         sPart코드 = Mid(tmpbc, 7, 1)
'         sSlip부분코드 = Mid(tmpbc, 8, 1)
'         s작업번호 = "0" & Mid(tmpbc, 9, 4)
'
'         Dim tmprno$     ' 상병등록번호
'         Dim tmptime$    ' 처방시간
'         Dim tmpinout$   ' 입원/외래
'
'      Rem WorkList Table(KLB030M)에서 검사항목 얻기
'
'         Dim Result%
'         SendBuf = ";" & Func & " "
'         ' Comment 96.11.28           SendBuf = SendBuf & SampNo & RackNo & PosiNo & Right$("          " + Trim(Id_Num), 13)
'         SendBuf = SendBuf & SampIf     '  Add 96.11.28
'         SendBuf = SendBuf & "48"
'         SendBuf = SendBuf & Mid$(TestDat, 1, 48)
'         SendBuf = SendBuf & "00000"
'         ChekSum = 0
'         For i = 1 To Len(SendBuf)
'             ChekSum = ChekSum + Asc(Mid$(SendBuf, i, 1))
'         Next i
'         ChkSumS = Hex$(ChekSum)
'         ' SPE Send
'         Dim Tmp$
'         Tmp = Chr$(2) & SendBuf & Chr$(3) & Right$(ChkSumS, 2) & Chr$(13)
'
'         'For i = 1 To 30000          ' loop
'         'Next i
'
'         MSComm1.Output = Tmp  'Chr$(2) & SendBuf & Chr$(3) & Right$(ChkSumS, 2) & Chr$(13)
'         'Print #1, tmp
'         Do
'         '   DoEvents
'         Loop Until MSComm1.OutBufferCount = 0
'
'
'         TestPtr = TestPtr + 1
'         If TestPtr > 2 Then TestPtr = 0
'
'      Case "1" To "9"               ' FR1 to FR9
'         For i = 1 To 10000        ' loop
'         Next i
'
'         Func = Mid$(Rcvbuffer, 2, 1)     ' Function
'         If Func <> "@" Then              ' TNNA ?
'            TstCount = Val(Trim$(Mid$(Rcvbuffer, 38, 2))) ' Test Count
'            StartPtr = 40
'            For i = 1 To TstCount
'                ResData = Mid$(Rcvbuffer, StartPtr, 9)
'                TestResultTable(TablePtr).TestNo = Right("0" & Trim(Mid$(ResData, 1, 2)), 2) ' Test Number
'                TestResultTable(TablePtr).Result = Mid$(ResData, 3, 6) ' Result
'                TablePtr = TablePtr + 1   ' Table Pointer Increament
'                StartPtr = StartPtr + 9   ' Start Pointer Increament
'            Next i
'         End If
'         ' MOR Send
'         MSComm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
'         Do
'         '   DoEvents
'         Loop Until MSComm1.OutBufferCount = 0
'
'      Case ":"                        ' END
'
'         Func = Mid$(Rcvbuffer, 2, 1)     ' Function
'         If Func = "K" Or Func = "L" Then
'            GoTo Nxt2
'         End If
'Rem modify  96.02.15
'         Dim tmpid$, tmpidcnt%
'         If Func <> "@" And Func <> "N" And Func <> "M" Then              ' TNNA
'            TestHeaderTable.RackNo = Trim$(Mid$(Rcvbuffer, 7, 1))  ' Rack No.
'            TestHeaderTable.SampNo = Right$("0000" & Trim$(Mid$(Rcvbuffer, 4, 3)), 5) ' Sample No.
'            TestHeaderTable.PosiNo = Trim$(Mid$(Rcvbuffer, 8, 2)) ' Position No.
'            tmpid = Trim$(Mid$(Rcvbuffer, 10, 13))
'            tmpid = Left(tmpid, 8) & "0" & Right(tmpid, 4)
'            tmpidcnt = Len(tmpid)
'            If tmpidcnt = 0 Then
'               GoTo Nxt2
'            End If
'            TestHeaderTable.Id_Num = tmpid ' Id No.
'            Dim s샘플번호$
'            s샘플번호 = tmpid
'            s작업일자 = H0SUB_DATE_6TO8(Mid(s샘플번호, 1, 6), "1")
'            sPart코드 = Mid(s샘플번호, 7, 1)
'            sSlip부분코드 = Mid(s샘플번호, 8, 1)
'            s작업번호 = Mid(s샘플번호, 9, 5)
'
'            SqlDoc = "select labdate from klb030m"
'            SqlDoc = SqlDoc & " where labdate = '" & s작업일자 & "'"
'            SqlDoc = SqlDoc & " and partcd = '" & sPart코드 & "'"
'            SqlDoc = SqlDoc & " and slipcd = '" & sSlip부분코드 & "'"
'            SqlDoc = SqlDoc & " and labsqno = '" & s작업번호 & "'"
'
'            Return_cd = QSqlDBExec(SqlDoc, QsqlConn)
'            If Return_cd <> QSQL_SUCCESS Then
'               Return_cd = QSqlSelectFree(QsqlConn)
'               Exit Sub
'            End If
'
'            Return_cd = QSqlGetRow(sStr, QsqlConn)
'            Dim b검사항목등록 As Boolean
'            b검사항목등록 = False
'            If Return_cd = QSQL_SUCCESS Then
'               b검사항목등록 = True
'               'QSqlGetField 1, sStr, tdata()
'               's등록번호 = tdata(1)
'            End If
'            Return_cd = QSqlSelectFree(QsqlConn)
'
'            TstCount = Val(Trim$(Mid$(Rcvbuffer, 38, 2))) ' Test Count
'            StartPtr = 40
'            For i = 1 To TstCount
'                ResData = Mid$(Rcvbuffer, StartPtr, 9)
'                TestResultTable(TablePtr).TestNo = Right("0" & Trim(Mid$(ResData, 1, 2)), 2) ' Test Number
'                TestResultTable(TablePtr).Result = Mid$(ResData, 3, 6) ' Result
'                TablePtr = TablePtr + 1   ' Table Pointer Increament
'                StartPtr = StartPtr + 9   ' Start Pointer Increament
'            Next i
'
'
'            ' Complete Item Search
'            Ptr1 = -1: Ptr2 = -1: Ptr3 = -1: Ptr4 = -1: Ptr5 = -1
'            For i = 0 To TablePtr - 1
'                If TestResultTable(i).TestNo = "05" Then Ptr1 = i    ' t-cho Search
'                If TestResultTable(i).TestNo = "18" Then Ptr2 = i    ' TG    Search
'                If TestResultTable(i).TestNo = "19" Then Ptr3 = i    ' FE    Search
'                If TestResultTable(i).TestNo = "20" Then Ptr4 = i    ' UIBC  Search
'                If TestResultTable(i).TestNo = "28" Then Ptr5 = i    ' HDL-C Search
'            Next i
'            'Dim s검사항목코드$
'            Dim s검체코드$, s검사항목부분코드$
'            ' LDL-C Complete
'            If Ptr1 <> -1 And Ptr2 <> -1 And Ptr5 <> -1 Then
'               TestResultTable(TablePtr).TestNo = "53"
'               WkResult = CInt(Val(Trim(TestResultTable(Ptr1).Result)) - (Val(Trim(TestResultTable(Ptr2).Result)) / 5) - Val(Trim(TestResultTable(Ptr5).Result)))
''                 WkResult = WkResult / 5
'               TestResultTable(TablePtr).Result = Format(Str$(WkResult), "@@@@@@")
'              ' TestResultTable(TablePtr).Result = Format(Trim(TestResultTable(TablePtr).Result), "######")
'               TablePtr = TablePtr + 1   ' Table Pointer Increament
'               If b검사항목등록 = True Then
'                  s검사항목코드 = TestNameTable(g검사명Index(53)).scode
'                  sPart코드 = Mid(s검사항목코드, 1, 1)
'                  sSlip부분코드 = Mid(s검사항목코드, 2, 1)
'                  s검체코드 = Mid(s검사항목코드, 3, 3)
'                  s검사항목부분코드 = Mid(s검사항목코드, 6, 3)
'                  SqlDoc = " { call add_test('"
'                  SqlDoc = SqlDoc & s작업일자 & "','"
'                  SqlDoc = SqlDoc & sPart코드 & "','"
'                  SqlDoc = SqlDoc & sSlip부분코드 & "','"
'                  SqlDoc = SqlDoc & s작업번호 & "','"
'                  SqlDoc = SqlDoc & s검체코드 & "','"
'                  SqlDoc = SqlDoc & s검사항목부분코드 & "') } "
'                  Return_cd = QSqlDBExec(SqlDoc, QsqlConn)
'                  If Return_cd <> QSQL_SUCCESS Then
'                     'return_cd = QSqlSelectFree(QsqlConn)
'                     'return_cd = QSqlRollBack()
'                     'Exit Sub
'                  End If
'                  Return_cd = QSqlSelectFree(QsqlConn)
'               End If
'            End If
'            ' T-CHO Complete
'            If Ptr1 <> -1 Then
'               TestResultTable(TablePtr).TestNo = "57"
'               WkResult = Val(Trim(TestResultTable(Ptr1).Result)) * 2.59 / 100
'               'TestResultTable(TablePtr).Result = Trim(Str(WkResult))
'               TestResultTable(TablePtr).Result = Format(WkResult, "##0.00")
'               TablePtr = TablePtr + 1   ' Table Pointer Increament
'               If b검사항목등록 = True Then
'                  s검사항목코드 = TestNameTable(g검사명Index(57)).scode
'                  sPart코드 = Mid(s검사항목코드, 1, 1)
'                  sSlip부분코드 = Mid(s검사항목코드, 2, 1)
'                  s검체코드 = Mid(s검사항목코드, 3, 3)
'                  s검사항목부분코드 = Mid(s검사항목코드, 6, 3)
'                  SqlDoc = " { call add_test('"
'                  SqlDoc = SqlDoc & s작업일자 & "','"
'                  SqlDoc = SqlDoc & sPart코드 & "','"
'                  SqlDoc = SqlDoc & sSlip부분코드 & "','"
'                  SqlDoc = SqlDoc & s작업번호 & "','"
'                  SqlDoc = SqlDoc & s검체코드 & "','"
'                  SqlDoc = SqlDoc & s검사항목부분코드 & "') } "
'                  Return_cd = QSqlDBExec(SqlDoc, QsqlConn)
'                  If Return_cd <> QSQL_SUCCESS Then
'
'                  End If
'                  Return_cd = QSqlSelectFree(QsqlConn)
'               End If
'            End If
'            ' TIBC Complete
'            If Ptr3 <> -1 And Ptr4 <> -1 Then
'               TestResultTable(TablePtr).TestNo = "60"
'               WkResult = Val(Trim(TestResultTable(Ptr3).Result)) + Val(Trim(TestResultTable(Ptr4).Result))
'               TestResultTable(TablePtr).Result = Format(WkResult, "###0.0")
'               TablePtr = TablePtr + 1   ' Table Pointer Increament
'               If b검사항목등록 = True Then
'                  s검사항목코드 = TestNameTable(g검사명Index(60)).scode
'                  sPart코드 = Mid(s검사항목코드, 1, 1)
'                  sSlip부분코드 = Mid(s검사항목코드, 2, 1)
'                  s검체코드 = Mid(s검사항목코드, 3, 3)
'                  s검사항목부분코드 = Mid(s검사항목코드, 6, 3)
'                  SqlDoc = " { call add_test('"
'                  SqlDoc = SqlDoc & s작업일자 & "','"
'                  SqlDoc = SqlDoc & sPart코드 & "','"
'                  SqlDoc = SqlDoc & sSlip부분코드 & "','"
'                  SqlDoc = SqlDoc & s작업번호 & "','"
'                  SqlDoc = SqlDoc & s검체코드 & "','"
'                  SqlDoc = SqlDoc & s검사항목부분코드 & "') } "
'                  Return_cd = QSqlDBExec(SqlDoc, QsqlConn)
'                  If Return_cd <> QSQL_SUCCESS Then
'
'                  End If
'                  Return_cd = QSqlSelectFree(QsqlConn)
'               End If
'            End If
'
'            Call addRecord
'            Call DataDisplay
'
'            If chk출력.Value = 1 Then
'               Call prtInterface
'            End If
'            TablePtr = 0
'         End If
'Nxt2:
'         ' MOR Send
'         MSComm1.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
'         Do
'         '   DoEvents
'         Loop Until MSComm1.OutBufferCount = 0
'
'   End Select
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        cmdQuery.SetFocus
    End If
    
End Sub


Private Sub cmdAdd_Click()

    Dim i   As Integer
    Dim K   As Integer
    Dim Tmp As Variant
    Dim temp    As Variant
    Dim ChkExist    As Integer
    
    With spdList
        If .MaxRows = 0 Then
            MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical
            Exit Sub
        End If
        
        ChkExist = False
        For i = 1 To .MaxRows
            Call .GetText(1, i, Tmp)
            If Tmp = True Then
                Call .GetText(2, i, Tmp)
                '--- WorkList Spread에 존재하는제 체크
                If spdWork.MaxRows <> 0 Then
                    For K = 1 To spdWork.MaxRows
                        Call spdWork.GetText(1, K, temp)
                        If Trim(Tmp) = Trim(temp) Then
                            MsgBox "접수번호 " & Trim(Tmp) & " 가 WorkList에 존재합니다. " _
                                    & "확인해 주십시오.", vbCritical
                            Exit Sub
                        End If
                    Next K
                End If
                '-------------------------------------
                spdWork.MaxRows = spdWork.MaxRows + 1
                Call spdWork.SetText(1, spdWork.MaxRows, Trim(Tmp))     '접수번호
                Call .GetText(3, i, Tmp)
                Call spdWork.SetText(2, spdWork.MaxRows, Trim(Tmp))     '성명
                
                .Col = -1: .Row = i
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
                i = i - 1
                
                ChkExist = True
            End If
        Next i
    End With
    
    If ChkExist <> True Then
        MsgBox "Work List에 등록할 자료가 없습니다. 조회를 실행해 주십시오.", vbCritical
    End If
    
End Sub

Private Sub cmdClear_Click()

    pnlSample = ""
    pnlPos = ""
    pnlStatus = ""
    
    spdList.MaxRows = 0
    spdWork.MaxRows = 0
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub




Private Sub CmdQuery_Click()

    Dim iRet    As Integer
    Dim sStr    As String
    Dim tData() As String
        
    '--- 조회조건 체크
    If Not IsDate(mskDate) Then
        MsgBox "조회를 원하는 접수일자를 입력해 주십시오.", vbExclamation
        mskDate.SetFocus
        Exit Sub
    End If
    
    '--- Index Open
    If S0SUB_Open(D0COM_SERVER01, Me.hwnd, QsqlConn) <> QSQL_SUCCESS Then Exit Sub
    
    spdList.MaxRows = 0
    
    SqlStr = " Select Distinct D1.LABDATE,  D1.NUMGBN, D1.LABSQNO, C2.PATNM " _
            & "  from LAB04_DB..DJD010M D1, LAB03_DB..DJC020M C2 " _
            & " where D1.LABDATE = '" & Trim(mskDate.ClipText) & "'" _
            & "   and SUBSTRING(D1.ORDCD,1,2) = 'CH' "
    If cboSelect.ListIndex = 0 Then
        SqlStr = SqlStr & "  and C2.ORDSTAT <> '1' "
    Else
        SqlStr = SqlStr & "  and C2.ORDSTAT = '1' "
    End If
    SqlStr = SqlStr & "  and SUBSTRING(D1.ORDERNO,1,8) = C2.ORDDATE " _
            & "  and SUBSTRING(D1.ORDERNO,9,2) = C2.DEPTCD " _
            & "  and SUBSTRING(D1.ORDERNO,11,5) = C2.SEQNO " _
            & "  and D1.IDNO = C2.IDNO " _
            & "order by D1.NUMGBN, D1.LABSQNO "
                
    If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
        Do
            iRet = QSqlGetRow(sStr, QsqlConn)
            If iRet <> QSQL_SUCCESS Then
                Exit Do
            End If
            
            QSqlGetField 4, sStr, tData()
                    
            With spdList
                .MaxRows = .MaxRows + 1
                Call .SetText(2, .MaxRows, tData(1) & "-" & tData(2) & "-" & tData(3))
                Call .SetText(3, .MaxRows, Trim(tData(4)))
            End With
        Loop Until (iRet <> QSQL_SUCCESS)
    End If
    iRet = QSqlSelectFree(QsqlConn)
    
    If spdList.MaxRows = 0 Then
        MsgBox "해당자료가 존재하지 않습니다.", vbExclamation
    End If
    
    '--- Index Close
    iRet = Qsqlclose(QsqlConn, ONECLOSE)
    
End Sub

Private Sub cmdResult_Click()

    On Error GoTo PortOpen_Err
        
    With cmdResult
        If .Caption = "결과 수신" Then
            Comm1.PortOpen = True
            .Caption = "수신 완료"
            '---
            sFrame1.Enabled = False
            cmdQuery.Enabled = False
            cmdAdd.Enabled = False
            spdList.Enabled = False
            '---
            pnlStatus = "결과 수신 Mode"
        Else
'            Timer1.Enabled = False
            Comm1.PortOpen = False
            .Caption = "결과 수신"
            '---
            sFrame1.Enabled = True
            cmdQuery.Enabled = True
            cmdAdd.Enabled = True
            spdList.Enabled = True
            '---
            pnlStatus = "Work List" & Chr(13) & "작성 Mode"
        End If
    End With
    
    spdRow = 1
    
    Exit Sub
    
PortOpen_Err:
    MsgBox "통신포트를 열 수 없습니다. 확인해 주십시오.", vbCritical
    
End Sub

Private Sub cmdServer_Click()
    
    On Error GoTo Err_Phase
    Load frm결과등록
    frm결과등록.Show

Err_Phase:
    Exit Sub
    
End Sub

Private Sub Comm1_OnComm()
'--- Hitachi 7060 (경북대)

   Dim Wkbuf As String
   Dim WkDat As String
   Dim ix1 As Integer

   Screen.MousePointer = 11
   
   Select Case Comm1.CommEvent
      ' Events
      Case MSCOMM_EV_SEND     ' There are SThreshold number of
                              ' character in the transmit buffer.
         ix1 = 1
      Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
         Wkbuf = Comm1.Input
         For ix1 = 1 To Len(Wkbuf)
            WkDat = Mid$(Wkbuf, ix1, 1)
            Print #1, WkDat;
            pnlMsg = pnlMsg & WkDat
            
            Select Case Asc(WkDat)
               Case 2         ' STX
                  Rcvbuffer = ""
                  pnlMsg = ""
               Case 3         ' EtX
   
               Case 10, 13        ' ****debug****
                  Call DataEditResponse
               Case 21         ' nak
                  pnlMsg = pnlMsg & "NAK"
               Case Else      ' Data
                  Rcvbuffer = Rcvbuffer + WkDat
            End Select
         Next ix1
      Case MSCOMM_EV_CTS      ' Change in the CTS line.
      Case MSCOMM_EV_DSR      ' Change in the DSR line.
      Case MSCOMM_EV_CD       ' Change in the CD line.
      Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
      ' Errors
      Case MSCOMM_ER_BREAK    ' A Break was received.
      ' Code to handle a BREAK goes here, and so on.
      Case MSCOMM_ER_CTSTO    ' CTS Timeout.
      Case MSCOMM_ER_DSRTO    ' DSR Timeout.
      Case MSCOMM_ER_FRAME    ' Framing Error.
      Case MSCOMM_ER_OVERRUN  ' Data Lost.
      Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
      Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
      Case MSCOMM_ER_RXPARITY ' Parity Error.
      Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
   End Select
   
   Screen.MousePointer = 0
   
End Sub


Private Sub Form_Activate()

    Timer1.Enabled = False
    Me.Caption = "Hitachi 747  결과받기"
    
End Sub

Private Sub Form_Deactivate()

    Timer1.Enabled = True
        
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit_Click
    End Select
    
End Sub

Private Sub Form_Load()

    Dim FileName As String
    
    FrmTag = 0
    TablePtr = 1
    '--- 화면 초기화
    mskDate = Format(Now, "yyyy-mm-dd")
    With cboSelect
        .AddItem ("미등록 자료")
        .AddItem ("등록된 자료")
        .ListIndex = 0
    End With
    '---------------
    FileName = strLAB_PATH & "data\hi_"
    
    '---------- 결과Table Open
    Set DbResult = OpenDatabase(FileName + tlFileName + ".mdb")
    
    Set TbResult = DbResult.OpenRecordset("Hi_Result")
    Set TbSample = DbResult.OpenRecordset("Hi_Sample")
            
    '---------- ComPort Open
    Comm1.CommPort = CommCfg.ComPort
    Comm1.Settings = CommCfg.BaudRate & CommCfg.Parity & CommCfg.DataBit & CommCfg.StopBit
    
    Open strLAB_PATH & "Hi_Test.log" For Output As #1   'Test 용
    Open strLAB_PATH & "Hitachi.log" For Output As #2   'Dump 용
        
    Screen.MousePointer = 0
 
End Sub


Private Sub Form_Unload(Cancel As Integer)

    '------ Port Close
    If Comm1.PortOpen <> False Then
        Comm1.PortOpen = False
    End If
    
    '------ DataBase Close
    TbResult.Close
    TbSample.Close
        
    DbResult.Close
          
    '------ Log Text Close
    Close #1
    Close #2

End Sub






Private Sub mskDate_GotFocus()

    With mskDate
        .SelStart = 8
        .SelLength = .MaxLength
    End With
    
End Sub


Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        cboSelect.SetFocus
    End If
    
End Sub


Private Sub sFrame1_Click()

End Sub

Private Sub spdList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim sNo As Long
    Dim eNo As Long
    Dim irow    As Integer
    Dim Tmp As Variant
    
    If BlockRow = 0 Or BlockRow2 = 0 Or BlockRow = BlockRow2 Then Exit Sub
    
    If BlockRow < BlockRow2 Then
        sNo = BlockRow: eNo = BlockRow2
    Else
        sNo = BlockRow2: eNo = BlockRow
    End If

    For irow = sNo To eNo
        With spdList
            Call .GetText(1, irow, Tmp)
            If Tmp = True Then
                Call .SetText(1, irow, "0")
            Else
                Call .SetText(1, irow, "1")
            End If
        End With
    Next irow
    
End Sub

Private Sub spdList_GotFocus()

    With spdList
        If OldRow <> 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = &H80000005
        End If
        If .ActiveRow = 1 Then
            .Row = .ActiveRow
            .Col = -1
    
            If .Lock = False Then
                .BackColor = &HEEEFFF
                OldRow = .ActiveRow
            End If
        End If
    End With
    
End Sub


Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Row <> NewRow Then
        If OldRow = 0 Then OldRow = Row
        
        With spdList
            If NewRow <> -1 Then
                .Row = OldRow
                .Col = -1
                .BackColor = &H80000005
                
                .Row = NewRow
                .Col = -1

                .BackColor = &HEEEFFF
                OldRow = NewRow
                'Call Disp_Data(OldRow)     '해당 Sample의 Order표시
            End If
        End With
    End If
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim Tmp     As Variant
    Dim iRet    As Integer
   
    If Row = 0 Then Exit Sub
        
    '--- spdWork 에서 삭제여부확인
    iRet = spdWork.GetText(1, Row, Tmp)
    If MsgBox("접수번호 " & Tmp & " 를 WorkList에서 삭제하시겠습니까?", vbYesNo) = vbNo Then
        Exit Sub
    End If
   
    '--- 삭제
    With spdList
        .MaxRows = .MaxRows + 1
        Call .SetText(2, .MaxRows, Tmp)       'Sample ID
        iRet = spdWork.GetText(2, Row, Tmp)
        Call .SetText(3, .MaxRows, Tmp)       '이름
    End With
    
    With spdWork
        .Row = Row
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1
    End With
    
End Sub


Private Sub spdWork_GotFocus()

    With spdWork
        If OldRow <> 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = &H80000005
        End If
        If .ActiveRow = 1 Then
            .Row = .ActiveRow
            .Col = -1
    
            If .Lock = False Then
                .BackColor = &HEEEFFF
                OldRow = .ActiveRow
            End If
        End If
    End With
    
End Sub


Private Sub spdWork_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Row <> NewRow Then
        If OldRow = 0 Then OldRow = Row
        
        With spdWork
            If NewRow <> -1 Then
                .Row = OldRow
                .Col = -1
                .BackColor = &H80000005
                
                .Row = NewRow
                .Col = -1

                .BackColor = &HEEEFFF
                OldRow = NewRow
            End If
        End With
    End If
    
End Sub


Private Sub SSCommand1_Click()
   
   Rcvbuffer = ";"
   Rcvbuffer = String(50, " ")
   Mid(Rcvbuffer, 1, 2) = "1 "
   Mid(Rcvbuffer, 38, 2) = "01"
   Mid(Rcvbuffer, 40, 9) = "05  23.5"
   'Mid(Rcvbuffer, 29 + 9, 9) = "06  26.6"
   'Rcvbuffer = Mid("1@")
   Call DataEditResponse
   Rcvbuffer = String(50, " ")
   Mid(Rcvbuffer, 1, 2) = "1 "
   Mid(Rcvbuffer, 38, 2) = "01"
   Mid(Rcvbuffer, 40, 9) = "18  28.8"
   'Mid(Rcvbuffer, 29 + 9, 9) = "06  26.6"
   'Rcvbuffer = Mid("1@")
   Call DataEditResponse
   Rcvbuffer = String(50, " ")
   Mid(Rcvbuffer, 1, 2) = "1 "
   Mid(Rcvbuffer, 38, 2) = "01"
   Mid(Rcvbuffer, 40, 9) = "19  29.9"
   'Mid(Rcvbuffer, 29 + 9, 9) = "06  26.6"
   'Rcvbuffer = Mid("1@")
   Call DataEditResponse
   Rcvbuffer = String(50, " ")
   Mid(Rcvbuffer, 1, 2) = "1 "
   Mid(Rcvbuffer, 38, 2) = "01"
   Mid(Rcvbuffer, 40, 9) = "20  20.0"
   'Mid(Rcvbuffer, 29 + 9, 9) = "06  26.6"
   'Rcvbuffer = Mid("1@")
   Call DataEditResponse
   Rcvbuffer = String(50, " ")
   Mid(Rcvbuffer, 1, 2) = ": "
   Mid(Rcvbuffer, 4, 3) = "028"
   'Mid(Rcvbuffer, 8, 5) = "00001"
   Mid(Rcvbuffer, 8, 2) = " 2"
   Mid(Rcvbuffer, 10, 13) = "             "
                            '1234567890123
   Mid(Rcvbuffer, 38, 2) = "01"
   'Mid(Rcvbuffer, 29, 9) = "03  23.5"
   Mid(Rcvbuffer, 40, 9) = "28  26.6"
   'Rcvbuffer = ";"
   Call DataEditResponse
   
End Sub


Private Sub SSPanel9_Click()

End Sub

Private Sub Timer1_Timer()

    If cmdResult.Caption = "수신 완료" Then
        Me.Caption = "결과 수신중 - " & pnlSample
    End If
    
End Sub

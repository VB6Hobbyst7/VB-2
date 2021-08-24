VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS422 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈자 등록"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   2295
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1425
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  헌혈자 등록 일자"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFEBD7&
      Caption         =   "신규 등록"
      Height          =   555
      Left            =   10890
      Style           =   1  '그래픽
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "15101"
      Top             =   1755
      Width           =   1335
   End
   Begin TabDlg.SSTab tabSel 
      Height          =   5325
      Left            =   2280
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2655
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   9393
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   14411494
      TabCaption(0)   =   "접수정보"
      TabPicture(0)   =   "frmBBS422.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picDiv(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "문진내역"
      TabPicture(1)   =   "frmBBS422.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picDiv(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "검사결과"
      TabPicture(2)   =   "frmBBS422.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picDiv(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   5010
         Index           =   2
         Left            =   -74985
         ScaleHeight     =   5010
         ScaleWidth      =   9900
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   300
         Width           =   9900
         Begin VB.TextBox txtRmk3 
            Height          =   3885
            Left            =   540
            MaxLength       =   499
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   63
            Text            =   "frmBBS422.frx":0054
            Top             =   555
            Width           =   8925
         End
         Begin MedControls1.LisLabel LisLabel10 
            Height          =   315
            Left            =   540
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   225
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " 검사결과"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel11 
            Height          =   315
            Left            =   540
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   4485
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   556
            ForeColor       =   16744576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   "※ 적격 판정된 검사 결과를 텍스트로 입력하십시오."
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   5010
         Index           =   1
         Left            =   15
         ScaleHeight     =   5010
         ScaleWidth      =   9900
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   300
         Width           =   9900
         Begin VB.CommandButton cmdJudge 
            BackColor       =   &H00FFC0C0&
            Caption         =   "적격판정"
            Height          =   315
            Left            =   8625
            Style           =   1  '그래픽
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   225
            Width           =   840
         End
         Begin MedControls1.LisLabel LisLabel8 
            Height          =   315
            Left            =   540
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   3645
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " Remark"
            Appearance      =   0
         End
         Begin VB.TextBox txtRmk2 
            Height          =   825
            Left            =   540
            MaxLength       =   499
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   60
            Top             =   3975
            Width           =   8925
         End
         Begin FPSpread.vaSpread tblAsk 
            Height          =   3030
            Left            =   540
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   540
            Width           =   8925
            _Version        =   196608
            _ExtentX        =   15743
            _ExtentY        =   5345
            _StockProps     =   64
            BackColorStyle  =   1
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   6
            MaxRows         =   100
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            SpreadDesigner  =   "frmBBS422.frx":0825
         End
         Begin MedControls1.LisLabel LisLabel9 
            Height          =   315
            Left            =   540
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   225
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " Remark"
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   5010
         Index           =   0
         Left            =   -74985
         ScaleHeight     =   5010
         ScaleWidth      =   9900
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   300
         Width           =   9900
         Begin MedControls1.LisLabel LisLabel6 
            Height          =   315
            Left            =   1770
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1605
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " 헌혈자 기본 진단"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Left            =   1770
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   450
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   "  헌혈 종류"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel5 
            Height          =   315
            Left            =   4935
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   450
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " 지정 환자"
            Appearance      =   0
         End
         Begin VB.TextBox txtRmk1 
            Height          =   825
            Left            =   1770
            MaxLength       =   499
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   54
            Top             =   3660
            Width           =   6285
         End
         Begin VB.Frame fraAcc3 
            BackColor       =   &H00DBE6E6&
            Height          =   1455
            Left            =   1770
            TabIndex        =   69
            Top             =   1830
            Width           =   6285
            Begin VB.TextBox txtWeight 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   330
               Left            =   1155
               TabIndex        =   37
               Top             =   255
               Width           =   930
            End
            Begin VB.TextBox txtHeight 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   330
               Left            =   4575
               TabIndex        =   40
               Top             =   225
               Width           =   945
            End
            Begin VB.TextBox txtPulse 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   315
               Left            =   1155
               TabIndex        =   51
               Top             =   1005
               Width           =   930
            End
            Begin VB.TextBox txtBldPres1 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   315
               Left            =   1155
               TabIndex        =   43
               Top             =   630
               Width           =   585
            End
            Begin VB.TextBox txtBldPres2 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   315
               Left            =   1875
               TabIndex        =   45
               Top             =   630
               Width           =   555
            End
            Begin VB.TextBox txtBodyTemp 
               Alignment       =   1  '오른쪽 맞춤
               Appearance      =   0  '평면
               Height          =   315
               Left            =   4590
               TabIndex        =   48
               Top             =   585
               Width           =   930
            End
            Begin MedControls1.LisLabel lbldt 
               Height          =   330
               Index           =   12
               Left            =   150
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   255
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "체중"
               Appearance      =   0
            End
            Begin MedControls1.LisLabel lbldt 
               Height          =   330
               Index           =   13
               Left            =   150
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   1005
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "맥박"
               Appearance      =   0
            End
            Begin MedControls1.LisLabel lbldt 
               Height          =   330
               Index           =   14
               Left            =   3570
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   585
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "체온"
               Appearance      =   0
            End
            Begin MedControls1.LisLabel lbldt 
               Height          =   330
               Index           =   15
               Left            =   3570
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   225
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "신장"
               Appearance      =   0
            End
            Begin MedControls1.LisLabel lbldt 
               Height          =   330
               Index           =   16
               Left            =   150
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   630
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "혈압"
               Appearance      =   0
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "kg"
               Height          =   180
               Left            =   2145
               TabIndex        =   38
               Top             =   375
               Width           =   195
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "Cm"
               Height          =   180
               Left            =   5565
               TabIndex        =   41
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "/Min"
               Height          =   180
               Left            =   2130
               TabIndex        =   52
               Top             =   1080
               Width           =   405
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "/"
               Height          =   180
               Left            =   1770
               TabIndex        =   44
               Top             =   690
               Width           =   90
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "mmHg"
               Height          =   180
               Left            =   2505
               TabIndex        =   46
               Top             =   690
               Width           =   555
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "℃"
               Height          =   180
               Left            =   5580
               TabIndex        =   49
               Top             =   660
               Width           =   180
            End
         End
         Begin VB.Frame fraAcc2 
            BackColor       =   &H00DBE6E6&
            Height          =   870
            Left            =   4935
            TabIndex        =   32
            Top             =   675
            Width           =   3135
            Begin VB.TextBox txtReservedID 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   165
               MaxLength       =   10
               TabIndex        =   33
               Text            =   "10293023"
               Top             =   315
               Width           =   1305
            End
            Begin MedControls1.LisLabel lblReservedNm 
               Height          =   315
               Left            =   1500
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   315
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               BackColor       =   12632256
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   0
               Alignment       =   1
               Caption         =   "홍길동"
               Appearance      =   0
            End
         End
         Begin VB.Frame fraAcc1 
            BackColor       =   &H00DBE6E6&
            Height          =   870
            Left            =   1770
            TabIndex        =   28
            Top             =   675
            Width           =   3135
            Begin VB.OptionButton optDonorCd 
               BackColor       =   &H00DBE6E6&
               Caption         =   "지정 헌혈"
               Height          =   435
               Index           =   0
               Left            =   315
               Style           =   1  '그래픽
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   240
               Width           =   1245
            End
            Begin VB.OptionButton optDonorCd 
               BackColor       =   &H00DBE6E6&
               Caption         =   "Pheresis"
               Height          =   435
               Index           =   1
               Left            =   1575
               Style           =   1  '그래픽
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   240
               Width           =   1245
            End
         End
         Begin MedControls1.LisLabel LisLabel7 
            Height          =   315
            Left            =   1770
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   3345
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   556
            BackColor       =   8388608
            ForeColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Caption         =   " Remark"
            Appearance      =   0
         End
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2302
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  헌혈자 기본정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2302
      TabIndex        =   1
      Top             =   450
      Width           =   9945
      Begin VB.CommandButton cmdNewReg 
         BackColor       =   &H00F4F0F2&
         Caption         =   "새 헌혈자 입력"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1050
         Style           =   1  '그래픽
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "15101"
         Top             =   525
         Width           =   1500
      End
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '평면
         Height          =   330
         Left            =   1050
         TabIndex        =   3
         Top             =   165
         Width           =   1515
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   5655
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   5655
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "총 헌혈량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   330
         Left            =   4290
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6645
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "M/100"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   330
         Left            =   8955
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCnt 
         Height          =   330
         Left            =   4290
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   525
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTotVol 
         Height          =   330
         Left            =   6645
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   525
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDonorID 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSSN 
         Height          =   315
         Left            =   8415
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   165
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
         Caption         =   "성   명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3300
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "헌혈횟수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7965
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "cc"
         Height          =   180
         Left            =   7605
         TabIndex        =   16
         Top             =   660
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   495
      Left            =   6720
      Style           =   1  '그래픽
      TabIndex        =   65
      Tag             =   "15101"
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   9270
      Style           =   1  '그래픽
      TabIndex        =   67
      Tag             =   "124"
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   10545
      Style           =   1  '그래픽
      TabIndex        =   68
      Tag             =   "128"
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "등록취소"
      Height          =   495
      Left            =   7995
      Style           =   1  '그래픽
      TabIndex        =   66
      Tag             =   "15101"
      Top             =   8205
      Width           =   1215
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2295
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2340
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  헌혈자 등록 내역"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   690
      Left            =   2295
      TabIndex        =   19
      Top             =   1650
      Width           =   8580
      Begin VB.ComboBox cboDonoraccdt 
         Height          =   300
         Left            =   1950
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   510
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
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
         Caption         =   "등록일자"
         Appearance      =   0
      End
      Begin VB.Label lblAccCnt 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         ForeColor       =   &H000040C0&
         Height          =   180
         Left            =   5235
         TabIndex        =   22
         Top             =   300
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmBBS422"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'헌혈의 약식
'접수, 문진등록, 검사결과 판정이 모두 적격으로 나온 경우에만 등록할 수 있도록 한다.
'혈액등록, Pheresis등록이 안된 경우에는 등록취소 할 수 있도록 한다.
'S2BBS601, S2BBS602, S2BBS603, S2BBS604, S2BBS605 생성

Private Const RowHeight& = 12
Private Const ASK_OK$ = "OK"
Private Const ASK_NOT$ = "NOT"
Private Const CANEDIT_STATUS& = 5   'STATUS 가 5인 경우까지 수정가능(즉, 헌혈등록이나 페리시스 등록이 안된 경우)

Private CurrentStatus As Long

Private Enum TblColumn
    tcASK = 1
    tcYES
    tcNo
    tcISOK
    tcNORMAL
    tcASKCODE
End Enum

Private Sub cboDonoraccdt_Click()
    Call InitInfo
    Call LoadAccInfo
    Call LoadAskRst
    On Error Resume Next
    txtReservedID.SetFocus
End Sub

Private Sub cmdCancel_Click()
'등록취소
    Dim strDonorid As String
    Dim strDonoraccdt As String
    Dim strSql1 As String
    Dim strSql2 As String
    Dim strSql3 As String
    
    
    If txtDonorNm.Text = "" Then Exit Sub
    
    If CurrentStatus = 0 Then
        MsgBox "헌혈자가 아직 등록되지 않았습니다. 등록된 헌혈자만 등록취소 할 수 있습니다.", vbExclamation
        Exit Sub
    End If

    If CurrentStatus > CANEDIT_STATUS Then
        MsgBox "헌혈 등록 및 Pheresis 등록이 되어 헌혈자 등록을 취소할 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("헌혈자 등록을 취소하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strDonorid = lblDonorID.Caption
    strDonoraccdt = Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT)
    
    'S2BBS602, S2BBS603, S2BBS604
    
    strSql1 = " delete " & T_BBS602 & _
              " where " & DBW("donorid=", strDonorid) & _
              " and " & DBW("donoraccdt=", strDonoraccdt)
    strSql2 = " delete " & T_BBS603 & _
              " where " & DBW("donorid=", strDonorid) & _
              " and " & DBW("donoraccdt=", strDonoraccdt)
    strSql3 = " delete " & T_BBS604 & _
              " where " & DBW("donorid=", strDonorid) & _
              " and " & DBW("donoraccdt=", strDonoraccdt)
    
    On Error GoTo ErrTrap
    DBConn.BeginTrans
    
    DBConn.Execute strSql1
    DBConn.Execute strSql2
    DBConn.Execute strSql3
    
    DBConn.CommitTrans
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    Call cmdClear_Click
    Exit Sub
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
End Sub

Private Sub cmdClear_Click()
    txtDonorNm.Text = ""
    Call InitDonor
    Call InitInfo
    On Error Resume Next
    txtDonorNm.SetFocus
End Sub

Private Sub InitDonor()
    lblDonorID.Caption = ""
    lblDOB.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    cboDonoraccdt.Clear
    lblAccCnt.Caption = ""
End Sub

Private Sub InitInfo()
    Dim i As Long
    
    CurrentStatus = 0
    
    tabSel.Tab = 0
    
    optDonorCd(0).value = False
    optDonorCd(1).value = False
    txtReservedID.Text = ""
    lblReservedNm.Caption = ""
    txtWeight.Text = ""
    txtHeight.Text = ""
    txtBldPres1.Text = ""
    txtBldPres2.Text = ""
    txtBodyTemp.Text = ""
    txtPulse.Text = ""
    txtRmk1.Text = ""
    
    With tblAsk
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcYES:   .value = 0
            .Col = TblColumn.tcNo:   .value = 0
            .Col = TblColumn.tcISOK: .value = ""
        Next i
    End With
    
    txtRmk2.Text = ""
    
    txtRmk3.Text = ""
    cmdCancel.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS422 = Nothing
End Sub

Private Sub cmdJudge_Click()
    Dim i As Long
    
    If tblAsk.Enabled = False Then Exit Sub
    
    If MsgBox("모두 적격인 결과로 입력하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    
    With tblAsk
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcNORMAL
            If .value = "1" Then
                .Col = TblColumn.tcYES: .value = 1
            ElseIf .value = "0" Then
                .Col = TblColumn.tcNo: .value = 1
            End If
        Next
    End With
End Sub

Private Sub cmdNew_Click()
    If txtDonorNm.Text = "" Then Exit Sub
    
    If medComboFind(cboDonoraccdt, Format(GetSystemDate, "YYYY-MM-DD")) >= 0 Then
        MsgBox "이미 헌혈자 등록이 진행중이거나 진행 완료된 상태입니다.", vbExclamation
    Else
        cboDonoraccdt.AddItem Format(GetSystemDate, "YYYY-MM-DD")
        cboDonoraccdt.ListIndex = cboDonoraccdt.ListCount - 1
        lblAccCnt.Caption = cboDonoraccdt.ListCount
        Call InitInfo
        Call LockControl(False)
    End If
End Sub

Private Sub cmdNewReg_Click()
    txtDonorNm.Text = ""
    Call InitDonor
    Call InitInfo
    
    frmBBS421.Show vbModal
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim strDonoraccdt As String
    Dim strDonorid As String
    
    If CheckValidation = False Then Exit Sub
    
    's2bbs602, s2bbs603, s2bbs604, s2bbs605
    
    strDonorid = lblDonorID.Caption
    strDonoraccdt = Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT)
    
On Error GoTo ErrTrap
    DBConn.BeginTrans
    '헌혈자 접수 시
    If SaveAccHx(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    '문진등록 시
    If SaveAsk(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    
    If SaveTest(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    
    If SaveOkNot(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    
    DBConn.CommitTrans
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    Call cmdClear_Click
    
    Exit Sub
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If txtDonorNm.Text = "" Then
        MsgBox "헌혈자 성명을 입력하십시오.", vbExclamation
        txtDonorNm.SetFocus
        Exit Function
    End If
    
    If CurrentStatus > CANEDIT_STATUS Then
        MsgBox "헌혈자의 혈액이 등록되어 있으므로 헌혈자에 대한 정보를 변경할 수 없습니다.", vbExclamation
        Exit Function
    End If
    
    If cboDonoraccdt.ListIndex < 0 Then
        MsgBox "등록일자를 선택하거나 신규등록하십시오.", vbExclamation
        cboDonoraccdt.SetFocus
        Exit Function
    End If
    
    If optDonorCd(0).value = False And optDonorCd(1).value = False Then
        MsgBox "헌혈 방법을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If txtReservedID.Text = "" Then
        MsgBox "지정 환자를 입력하십시오.", vbExclamation
        txtReservedID.SetFocus
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Function SaveAccHx(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
'헌혈자 접수시
'S2BBS602 생성
'DonorId , donoraccdt, DONORCD, RESERVEDID, tmpid, Weight, Height, PULSE, BODYTEMP,bldpres1,bldpres2
'
'S2BBS603 생성
'DonorId , donoraccdt, stscd(2), OKDIV1(1), OKDT1, RMK1
    Dim strSql1 As String
    Dim strSql2 As String
    Dim strSql3 As String
    Dim strSql4 As String
    Dim strDonorCd As String
    Dim strOkdt1 As String
    
    If optDonorCd(0).value Then
        strDonorCd = "1"
    ElseIf optDonorCd(1).value Then
        strDonorCd = "3"
    End If
    strOkdt1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    strSql1 = " delete " & T_BBS602 & _
              " where " & DBW("donorid=", vDonorid) & _
              " and " & DBW("donoraccdt=", vDonoraccdt)
    
    strSql2 = " delete " & T_BBS603 & _
              " where " & DBW("donorid=", vDonorid) & _
              " and " & DBW("donoraccdt=", vDonoraccdt)
    
    strSql3 = "insert into " & T_BBS602 & _
            " (donorid,donoraccdt,donorcd,reservedid,weight,height,pulse,bodytemp,bldpres1,bldpres2) values (" & _
            DBV("donorid", vDonorid, 1) & DBV("donoraccdt", vDonoraccdt, 1) & _
            DBV("donorcd", strDonorCd, 1) & DBV("reservedid", txtReservedID.Text, 1) & _
            DBV("weight", txtWeight.Text, 1) & DBV("height", txtHeight.Text, 1) & DBV("pulse", txtPulse.Text, 1) & _
            DBV("bodytemp", txtBodyTemp.Text, 1) & DBV("bldpres1", txtBldPres1.Text, 1) & DBV("bldpres2", txtBldPres2.Text) & " )"
    
    strSql4 = "insert into " & T_BBS603 & "(donorid,donoraccdt,stscd,okdiv1,okdt1,rmk1) " & _
              "values(" & _
              DBV("donorid", vDonorid, 1) & DBV("donoraccdt", vDonoraccdt, 1) & DBV("stscd", "2", 1) & DBV("okdiv1", "1", 1) & _
              DBV("okdt1", strOkdt1, 1) & _
              DBV("rmk1", txtRmk1.Text) & ")"
    
    On Error GoTo ErrTrap
    DBConn.Execute strSql1
    DBConn.Execute strSql2
    DBConn.Execute strSql3
    DBConn.Execute strSql4
    
    SaveAccHx = True
    Exit Function
ErrTrap:
    SaveAccHx = False
End Function

Private Function SaveAsk(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
'문진등록 시
'S2BBS603 Update
'stscd (4), OKDIV2, OKDT2, RMK2
'
'S2BBS604 생성(문진 문항 갯수만큼 생성)
'DonorId , donoraccdt, askcd, yesno, okdiv
    Dim strYes As String
    Dim strNo As String
    Dim strAskCd As String
    Dim strYesNo As String
    Dim strOkdiv As String
    Dim arySql() As String
    Dim strOkdt2 As String
    Dim i As Long
    
    strOkdt2 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    ReDim arySql(1)
    
    arySql(0) = " update " & T_BBS603 & " " & _
                " set " & DBW("stscd", "4", 3) & _
                        DBW("okdiv2", "1", 3) & _
                        DBW("okdt2", strOkdt2, 3) & _
                        DBW("rmk2", txtRmk2.Text, 2) & _
               " where " & DBW("donorid", vDonorid, 2) & _
               " AND " & DBW("donoraccdt", vDonoraccdt, 2)
    
    arySql(1) = " delete FROM " & T_BBS604 & " " & _
                " WHERE " & DBW("donorid", vDonorid, 2) & _
                " AND " & DBW("donoraccdt", vDonoraccdt, 2)
    With tblAsk
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcASKCODE: strAskCd = .value
            .Col = TblColumn.tcYES:     strYes = .value
            .Col = TblColumn.tcNo:      strNo = .value
                                        If strYes = "1" Then
                                            strYesNo = "1"
                                        ElseIf strNo = "1" Then
                                            strYesNo = "0"
                                        Else
                                            strYesNo = ""
                                        End If

            .Col = TblColumn.tcISOK:
                                        If .value = ASK_OK Then
                                            strOkdiv = "1"
                                        ElseIf .value = ASK_NOT Then
                                            strOkdiv = "0"
                                        Else
                                            strOkdiv = ""
                                        End If
            
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = "insert into " & T_BBS604 & "(donorid,donoraccdt,askcd,yesno,okdiv) " & _
                                        "values(" & _
                                        DBV("donorid", vDonorid, 1) & DBV("donoraccdt", vDonoraccdt, 1) & DBV("askcd", strAskCd, 1) & _
                                        DBV("yesno", strYesNo, 1) & DBV("okdiv", strOkdiv) & ")"
        Next i
    End With
    
    On Error GoTo ErrTrap
    
    For i = LBound(arySql) To UBound(arySql)
        If arySql(i) <> "" Then DBConn.Execute arySql(i)
    Next i
    SaveAsk = True
    Exit Function
    
ErrTrap:
    SaveAsk = False
End Function

Private Function SaveTest(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
'검사의뢰시
'S2BBS603 Update
'stscd (5)
'
'S2BBS602 Update
'Cancelfg (0)
'
'S2BBS605 생성(아닐 수도 있음)
'DonorId , donoraccdt, orddt, Seq, WorkArea, accdt, accseq
    Dim strSql1 As String
    Dim strSql2 As String
    
    strSql1 = " update " & T_BBS603 & " " & _
              " set    " & DBW("stscd", "5", 2) & _
              " WHERE " & DBW("donorid", vDonorid, 2) & _
              " AND   " & DBW("donoraccdt", vDonoraccdt, 2)
    strSql2 = " update " & T_BBS602 & " set " & DBW("cancelfg", "0", 2) & _
              " WHERE " & DBW("donorid=", vDonorid) & " AND " & DBW("donoraccdt=", vDonoraccdt)
    
    On Error GoTo ErrTrap
    
    DBConn.Execute strSql1
    DBConn.Execute strSql2
    SaveTest = True
    
    Exit Function
ErrTrap:
    SaveTest = False
End Function

Private Function SaveOkNot(ByVal vDonorid As String, ByVal vDonoraccdt As String) As String
'S2BBS603 Update
'stscd (6), OKDIV3, OKDT3, RMK3// stscd 는 혈액 등록시 변경
    Dim strSQL As String
    Dim strOkdt3 As String
    
    strOkdt3 = Format(GetSystemDate, PRESENTDATE_FORMAT)

    
'                          DBW("stscd", DonorStatus.stsFinish, 2)
    strSQL = " update " & T_BBS603 & _
            " set " & DBW("okdiv3", "1", 3) & _
                      DBW("okdt3", strOkdt3, 3) & _
                      DBW("rmk3", txtRmk3.Text, 2) & _
            " WHERE " & DBW("donorid", vDonorid, 2) & _
            " AND " & DBW("donoraccdt", vDonoraccdt, 2)

    On Error GoTo ErrTrap
    
    DBConn.Execute strSQL
    SaveOkNot = True
    
    Exit Function
ErrTrap:
    SaveOkNot = False
End Function

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    txtDonorNm.Text = ""
    Call InitDonor
    Call InitInfo
    If tabSel.Tab = 0 Then
        For i = picDiv.LBound To picDiv.UBound
            picDiv(i).Enabled = False
        Next
        picDiv(tabSel.Tab).Enabled = True
    Else
        tabSel.Tab = 0
    End If
    
    Call LoadAskList
End Sub

Private Sub LoadAskList()
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT * FROM " & T_COM003
    strSQL = strSQL & " WHERE " & DBW("cdindex=", BC2_ASK)
    strSQL = strSQL & " ORDER BY field2"
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    Call medClearTable(tblAsk)
    
    With tblAsk
        .ReDraw = False
        If RS.RecordCount < 9 Then
            .MaxRows = 9
        Else
            .MaxRows = RS.RecordCount
        End If
        .RowHeight(-1) = RowHeight
        
        Do Until RS.EOF
            .Row = .DataRowCnt + 1
            
            .Col = TblColumn.tcASK: .value = RS.Fields("text1").value & ""
            .Col = TblColumn.tcNORMAL: .value = RS.Fields("field1").value & ""
            .Col = TblColumn.tcASKCODE: .value = RS.Fields("cdval1").value & ""
            
            .RowHeight(.Row) = .MaxTextRowHeight(.Row) + 2
            If .RowHeight(.Row) < RowHeight Then
                .RowHeight(.Row) = RowHeight
            End If
            RS.MoveNext
        Loop

        .ReDraw = True
    End With
    
    Set RS = Nothing
End Sub

Private Sub tabSel_Click(PreviousTab As Integer)
    Dim i As Long
    
    For i = picDiv.LBound To picDiv.UBound
        picDiv(i).Enabled = False
    Next
    picDiv(tabSel.Tab).Enabled = True
End Sub

Private Sub tblAsk_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim strNormal As String
    Dim strChk1 As String
    Dim strChk2 As String
    Dim strYN As String
    
    If txtDonorNm.Text = "" Then Exit Sub
    If Row < 1 Or Row > tblAsk.DataRowCnt Then Exit Sub
    If Col < TblColumn.tcYES Or Col > TblColumn.tcNo Then Exit Sub

    With tblAsk
        '정상치
        .Row = Row: .Col = TblColumn.tcNORMAL: strNormal = .value
        
        .Col = TblColumn.tcYES: strChk1 = .value
        .Col = TblColumn.tcNo: strChk2 = .value

        If strChk1 = "0" And strChk2 = "0" Then
            .Col = TblColumn.tcISOK: .value = ""
        ElseIf strChk1 = "1" And strChk2 = "1" Then
            .Col = IIf(Col = TblColumn.tcYES, TblColumn.tcNo, TblColumn.tcYES): .value = 0
        Else
            If strChk1 = "1" And strChk2 = "0" Then
                strYN = "1"
            ElseIf strChk1 = "0" And strChk2 = "1" Then
                strYN = "0"
            End If
            
            .Col = TblColumn.tcISOK
            .value = IIf(strYN = strNormal, ASK_OK, ASK_NOT)
        End If
    End With
End Sub

Private Sub txtBldPres1_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtBldPres2_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtBodyTemp_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDonorNm_Change()
    If lblDonorID.Caption <> "" Then
        Call InitDonor
        Call InitInfo
    End If
End Sub

Private Sub txtDonorNm_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDonorNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDonorNm.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDonorNm_Validate(Cancel As Boolean)
    If txtDonorNm.Text = "" Then Exit Sub
    If lblDonorID.Caption <> "" Then Exit Sub
    
    If DonorFind = False Then
        Cancel = True
        MsgBox "등록된 헌혈자가 아닙니다. 먼저 새 헌혈자를 입력하십시오.", vbExclamation
    Else
        Call ShowAccList
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function DonorFind() As Boolean
    Dim objDonor As clsBBSBldDonationBusi
        
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor
        DonorFind = .DonorFind(txtDonorNm.Text)
            
        lblDonorID.Caption = .mDonorID
'        txtDonorNm = .mDonorNm
        lblDOB.Caption = .mDOB
        lblSex.Caption = .mSEX
        lblABO.Caption = .mABO
        lblCnt.Caption = .Mcnt
        lblTotVol.Caption = .mTotVol
        lblSSN.Caption = .mSSN
    End With
    Set objDonor = Nothing
End Function

Private Sub ShowAccList()
    Dim strAccDt    As String
    Dim RS          As Recordset
    Dim objMySQL    As clsBBSSQLStatement
    '헌혈자에 대해서 접수된 정보가 있을 경우에 접수 내역을 보여준다.

    Set objMySQL = New clsBBSSQLStatement
    Set RS = New Recordset

    Set RS = objMySQL.GetDonorAccHistory(lblDonorID.Caption)
    
    cboDonoraccdt.Clear
    Do Until RS.EOF
        strAccDt = Format(RS.Fields("donoraccdt").value & "", "####-##-##")
        cboDonoraccdt.AddItem strAccDt
        RS.MoveNext
    Loop
    If cboDonoraccdt.ListCount > 0 Then
        lblAccCnt.Caption = cboDonoraccdt.ListCount
        cboDonoraccdt.ListIndex = 0
    Else
        Call cmdNew_Click
    End If
    
    Set RS = Nothing
    Set objMySQL = Nothing
End Sub

Private Sub LoadAccInfo()
'등록되어 있는 기본정보 조회
    Dim RS As Recordset
    Dim strSQL As String
    
    'donorcd 0:지정헌혈,3:페리시스
    
    strSQL = " SELECT a.donorcd,a.reservedid, a.weight, a.height, a.pulse, a.bodytemp,a.bldpres1, a.bldpres2,  b.stscd, b.rmk1, b.rmk2, b.rmk3 "
    strSQL = strSQL & " FROM " & T_BBS602 & " a, " & T_BBS603 & " b "
    strSQL = strSQL & " WHERE a.donorid=b.donorid"
    strSQL = strSQL & " AND " & DBW("a.donorid=", lblDonorID.Caption)
    strSQL = strSQL & " AND " & DBW("a.donoraccdt=", Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT))
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    With RS
        If .EOF = False Then
            If .Fields("donorcd").value & "" = "1" Then
                optDonorCd(0).value = True
            ElseIf .Fields("donorcd").value & "" = "3" Then
                optDonorCd(1).value = True
            End If
            
            txtReservedID.Text = .Fields("reservedid").value & ""
            lblReservedNm.Caption = GetPtNm(txtReservedID.Text)
            txtWeight.Text = .Fields("weight").value & ""
            txtHeight.Text = .Fields("height").value & ""
            txtPulse.Text = .Fields("pulse").value & ""
            txtBodyTemp.Text = .Fields("bodytemp").value & ""
            txtBldPres1.Text = .Fields("bldpres1").value & ""
            txtBldPres2.Text = .Fields("bldpres2").value & ""
            
            txtRmk1.Text = .Fields("rmk1").value & ""
            txtRmk2.Text = .Fields("rmk2").value & ""
            txtRmk3.Text = .Fields("rmk3").value & ""
            
            CurrentStatus = .Fields("stscd").value & ""
            
            If Val(.Fields("stscd").value & "") > CANEDIT_STATUS Then
                Call LockControl(True)
            Else
                Call LockControl(False)
            End If
        End If
    End With
    
    Set RS = Nothing
End Sub

Private Sub LoadAskRst()
'등록되어 있는 문진 내역 조회
    Dim RS As Recordset
    Dim strSQL As String
    Dim i As Long
    Dim r As Long
    Dim strAskCd As String
    Dim strYesNo As String
    Dim strOkdiv As String
    
    strSQL = " SELECT * FROM s2bbs604"
    strSQL = strSQL & " WHERE " & DBW("donorid=", lblDonorID.Caption)
    strSQL = strSQL & " AND " & DBW("donoraccdt=", Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT))
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        Set RS = Nothing
        Exit Sub
    End If

    With tblAsk
        For i = 1 To RS.RecordCount
            '스프레트에 있는 문진크도와
            '데이터베이스에서 읽은 문진코드가
            '동일한 Row에 대하여 작업한다.
            For r = 1 To .MaxRows
                .Row = r
                .Col = TblColumn.tcASKCODE
                strAskCd = Trim(.value)
                
                If strAskCd = Trim(RS.Fields("askcd").value & "") Then
                    strYesNo = RS.Fields("yesno").value & ""
                    strOkdiv = RS.Fields("okdiv").value & ""
                    
                    If strYesNo = "1" Then
                        .Col = TblColumn.tcYES: .value = 1
                        .Col = TblColumn.tcNo: .value = 0
                    ElseIf strYesNo = "0" Then
                        .Col = TblColumn.tcYES: .value = 0
                        .Col = TblColumn.tcNo: .value = 1
                    Else
                        .Col = TblColumn.tcYES: .value = 0
                        .Col = TblColumn.tcNo: .value = 0
                    End If
                    
                    If strOkdiv = "1" Then
                        .Col = TblColumn.tcISOK: .value = ASK_OK
                    ElseIf strOkdiv = "0" Then
                        .Col = TblColumn.tcISOK: .value = ASK_NOT
                    Else
                        .Col = TblColumn.tcISOK: .value = ""
                    End If
                    
                    Exit For
                End If
            Next r
            RS.MoveNext
        Next i
    End With
    
    Set RS = Nothing
End Sub

Private Sub LockControl(ByVal vLock As Boolean)
    If vLock Then
        fraAcc1.Enabled = False
        fraAcc2.Enabled = False
        fraAcc3.Enabled = False
        txtRmk1.Locked = True
        
        tblAsk.Enabled = False
        txtRmk2.Locked = True
        
        txtRmk3.Locked = True
        
        cmdCancel.Enabled = False
    Else
        fraAcc1.Enabled = True
        fraAcc2.Enabled = True
        fraAcc3.Enabled = True
        txtRmk1.Locked = False
        
        tblAsk.Enabled = True
        txtRmk2.Locked = False
        
        txtRmk3.Locked = False
        
        cmdCancel.Enabled = True
    End If
End Sub

Private Sub txtHeight_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPulse_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtReservedID_Change()
    If lblReservedNm.Caption <> "" Then
        lblReservedNm.Caption = ""
    End If
End Sub

Private Sub txtReservedID_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtReservedID_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtReservedID.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtReservedID_Validate(Cancel As Boolean)
    Dim strReservedNm As String
    
    If txtReservedID.Text = "" Then Exit Sub
    
    strReservedNm = GetPtNm(txtReservedID.Text)
    
    If strReservedNm = "" Then
        Cancel = True
        MsgBox "등록되지 않은 환자입니다.", vbExclamation
    Else
        lblReservedNm.Caption = strReservedNm
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtWeight_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00BF8B59&
   Caption         =   "검사설정"
   ClientHeight    =   12300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21990
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12300
   ScaleWidth      =   21990
   WindowState     =   2  '최대화
   Begin VB.Frame frameTestSet 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   10965
      Left            =   10950
      TabIndex        =   1
      Top             =   60
      Width           =   7545
      Begin BHButton.BHImageButton cmdView 
         Height          =   360
         Left            =   930
         TabIndex        =   85
         Top             =   4860
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   635
         Caption         =   "검사결과 변환 ▼"
         CaptionChecked  =   "V"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmTestSet.frx":1272
         PictureAlignment=   0
         ButtonAttrib    =   2
         ForeColor       =   0
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin VB.TextBox txtEqpCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   270
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H00FFFFFF&
         Height          =   4125
         Left            =   180
         TabIndex        =   49
         Top             =   720
         Width           =   7095
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   465
            Left            =   1110
            TabIndex        =   86
            Top             =   3630
            Width           =   5955
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치형"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   89
               Top             =   180
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "문자형"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   1200
               TabIndex        =   88
               Top             =   180
               Width           =   1185
            End
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치형 / 문자형 혼합"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   2400
               TabIndex        =   87
               Top             =   180
               Width           =   2415
            End
         End
         Begin VB.TextBox txtRefFHigh 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4710
            TabIndex        =   62
            Top             =   2775
            Width           =   1500
         End
         Begin VB.TextBox txtRefFLow 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2940
            TabIndex        =   61
            Top             =   2775
            Width           =   1500
         End
         Begin VB.TextBox txtRefMHigh 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4710
            TabIndex        =   60
            Top             =   2385
            Width           =   1500
         End
         Begin VB.TextBox txtRefMLow 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2940
            TabIndex        =   59
            Top             =   2385
            Width           =   1500
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   58
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txtResSpec 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4350
            TabIndex        =   57
            Top             =   3240
            Width           =   945
         End
         Begin VB.TextBox txtAbbrNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            TabIndex        =   56
            Top             =   855
            Width           =   2025
         End
         Begin VB.TextBox txtOChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   55
            Top             =   495
            Width           =   2025
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   54
            Top             =   855
            Width           =   2025
         End
         Begin VB.TextBox txtTestCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3210
            TabIndex        =   53
            Top             =   1260
            Width           =   1725
         End
         Begin VB.TextBox txtRChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            TabIndex        =   52
            Top             =   495
            Width           =   2025
         End
         Begin VB.CheckBox chkResSpec 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            Caption         =   "소수점 변환사용"
            Height          =   180
            Left            =   1080
            TabIndex        =   51
            Top             =   3330
            Width           =   1665
         End
         Begin VB.ListBox lstTestCode 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   1110
            TabIndex        =   50
            Top             =   1245
            Width           =   2025
         End
         Begin BHButton.BHImageButton cmdSeqUp 
            Height          =   345
            Left            =   2340
            TabIndex        =   63
            Top             =   120
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            Caption         =   "▲"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":37E4
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   33023
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdSeqDown 
            Height          =   345
            Left            =   2760
            TabIndex        =   64
            Top             =   120
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            Caption         =   "▼"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":53D2
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   16744576
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdAdd 
            Height          =   420
            Left            =   4980
            TabIndex        =   65
            Top             =   1260
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   741
            Caption         =   "코드Add"
            CaptionChecked  =   "V"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":6FC0
            PictureAlignment=   0
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   33023
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdRemove 
            Height          =   420
            Left            =   4980
            TabIndex        =   66
            Top             =   1725
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   741
            Caption         =   "코드Remove"
            CaptionChecked  =   "V"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":9532
            PictureAlignment=   0
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   16744576
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdSpecUP 
            Height          =   345
            Left            =   5340
            TabIndex        =   67
            Top             =   3240
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            Caption         =   "▲"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":BAA4
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   33023
            ImgOutLineSize  =   3
         End
         Begin BHButton.BHImageButton cmdSpecDown 
            Height          =   345
            Left            =   5760
            TabIndex        =   68
            Top             =   3240
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            Caption         =   "▼"
            CaptionChecked  =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "frmTestSet.frx":D692
            ButtonAttrib    =   2
            ForeColor       =   16777215
            BackColor       =   16744576
            ImgOutLineSize  =   3
         End
         Begin VB.Label Label14 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 결과형태"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   108
            Top             =   3720
            Width           =   1035
         End
         Begin VB.Label Label13 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 변환자릿수"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3000
            TabIndex        =   107
            Top             =   3270
            Width           =   1035
         End
         Begin VB.Label Label12 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 소수점"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   106
            Top             =   3240
            Width           =   1035
         End
         Begin VB.Label Label10 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 여성 참고치"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1110
            TabIndex        =   104
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label9 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 남성 참고치"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1170
            TabIndex        =   103
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label8 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 참고치"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   102
            Top             =   2460
            Width           =   1035
         End
         Begin VB.Label Label7 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 검사코드's"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   60
            TabIndex        =   101
            Top             =   1230
            Width           =   1035
         End
         Begin VB.Label Label6 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 결사약어"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3180
            TabIndex        =   100
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label5 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 검사명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   90
            TabIndex        =   99
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label4 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 결과채널"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3150
            TabIndex        =   98
            Top             =   510
            Width           =   1035
         End
         Begin VB.Label Label3 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 오더채널"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   60
            TabIndex        =   97
            Top             =   510
            Width           =   1035
         End
         Begin VB.Label Label2 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   " 순번"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   60
            TabIndex        =   96
            Top             =   135
            Width           =   1035
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   4500
            TabIndex        =   70
            Top             =   2880
            Width           =   135
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   4500
            TabIndex        =   69
            Top             =   2490
            Width           =   135
         End
         Begin VB.Image Image1 
            Height          =   1260
            Left            =   6300
            Picture         =   "frmTestSet.frx":F280
            Top             =   2115
            Width           =   705
         End
      End
      Begin VB.Frame fraResultTrans 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5685
         Left            =   180
         TabIndex        =   2
         Top             =   5070
         Visible         =   0   'False
         Width           =   7095
         Begin VB.Frame fraNC 
            Height          =   495
            Left            =   90
            TabIndex        =   90
            Top             =   5160
            Width           =   6915
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치 판정"
               Height          =   255
               Index           =   4
               Left            =   6420
               TabIndex        =   95
               Top             =   240
               Width           =   1125
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "판정 수치"
               Height          =   255
               Index           =   3
               Left            =   4950
               TabIndex        =   94
               Top             =   240
               Width           =   1125
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치(판정)"
               Height          =   255
               Index           =   2
               Left            =   3750
               TabIndex        =   93
               Top             =   240
               Width           =   1185
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "판정(수치)"
               Height          =   255
               Index           =   1
               Left            =   2550
               TabIndex        =   92
               Top             =   240
               Width           =   1185
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "변환없음"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   91
               Top             =   240
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.Label Label19 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   " 결과표기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   0
               TabIndex        =   113
               Top             =   120
               Width           =   1035
            End
         End
         Begin VB.Frame fraC 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2805
            Left            =   60
            TabIndex        =   16
            Top             =   2340
            Width           =   7005
            Begin VB.Frame fraTypeChange 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               Height          =   2835
               Left            =   4050
               TabIndex        =   34
               Top             =   -15
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   7
                  Left            =   60
                  TabIndex        =   48
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   7
                  Left            =   1650
                  TabIndex        =   47
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   8
                  Left            =   60
                  TabIndex        =   46
                  Top             =   450
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   8
                  Left            =   1650
                  TabIndex        =   45
                  Top             =   450
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   9
                  Left            =   60
                  TabIndex        =   44
                  Top             =   780
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   9
                  Left            =   1650
                  TabIndex        =   43
                  Top             =   780
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   10
                  Left            =   60
                  TabIndex        =   42
                  Top             =   1110
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   10
                  Left            =   1650
                  TabIndex        =   41
                  Top             =   1110
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   11
                  Left            =   60
                  TabIndex        =   40
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   11
                  Left            =   1650
                  TabIndex        =   39
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   12
                  Left            =   60
                  TabIndex        =   38
                  Top             =   1770
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   12
                  Left            =   1650
                  TabIndex        =   37
                  Top             =   1770
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   13
                  Left            =   60
                  TabIndex        =   36
                  Top             =   2100
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   13
                  Left            =   1650
                  TabIndex        =   35
                  Top             =   2100
                  Width           =   1215
               End
               Begin BHButton.BHImageButton cmdUnView 
                  Height          =   360
                  Left            =   1920
                  TabIndex        =   75
                  Top             =   2460
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  Caption         =   "숨김 ▶"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":10AF2
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin BHButton.BHImageButton cmdConfirm 
                  Height          =   360
                  Index           =   4
                  Left            =   900
                  TabIndex        =   76
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  Caption         =   "적용"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":13064
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin BHButton.BHImageButton cmdConfirm 
                  Height          =   360
                  Index           =   5
                  Left            =   60
                  TabIndex        =   78
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   635
                  Caption         =   "전체적용"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":155D6
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin VB.Label Label21 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  '단일 고정
                  Caption         =   "▶"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Left            =   1380
                  TabIndex        =   115
                  Top             =   630
                  Width           =   255
               End
            End
            Begin VB.TextBox txtAMRLimit 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   1440
               TabIndex        =   22
               Top             =   780
               Width           =   1005
            End
            Begin VB.TextBox txtAMRResult 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   2820
               TabIndex        =   21
               Top             =   780
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   1440
               TabIndex        =   20
               Top             =   450
               Width           =   1005
            End
            Begin VB.TextBox txtAMRResult 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   2820
               TabIndex        =   19
               Top             =   450
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   1440
               TabIndex        =   18
               Top             =   120
               Width           =   1005
            End
            Begin VB.TextBox txtAMRResult 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   2820
               TabIndex        =   17
               Top             =   120
               Width           =   1125
            End
            Begin BHButton.BHImageButton cmdTypeChange 
               Height          =   420
               Left            =   1020
               TabIndex        =   72
               Top             =   1200
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   741
               Caption         =   "◀ 더 많은 문자결과변환"
               CaptionChecked  =   "V"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TransparentPicture=   "frmTestSet.frx":17B48
               PictureAlignment=   0
               ButtonAttrib    =   2
               ForeColor       =   16777215
               BackColor       =   12553049
               ImgOutLineSize  =   3
            End
            Begin VB.Label Label18 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "▶"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   2550
               TabIndex        =   112
               Top             =   180
               Width           =   255
            End
            Begin VB.Label Label17 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   " 검사결과"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   270
               TabIndex        =   111
               Top             =   120
               Width           =   1035
            End
         End
         Begin VB.Frame fraN 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            ForeColor       =   &H80000008&
            Height          =   2205
            Left            =   60
            TabIndex        =   3
            Top             =   120
            Width           =   7005
            Begin VB.Frame fraNTypeChange 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               Height          =   2295
               Left            =   4050
               TabIndex        =   23
               Top             =   -90
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   18
                  Left            =   1650
                  TabIndex        =   33
                  Top             =   1530
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   18
                  Left            =   60
                  TabIndex        =   32
                  Top             =   1530
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   17
                  Left            =   1650
                  TabIndex        =   31
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   17
                  Left            =   60
                  TabIndex        =   30
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   16
                  Left            =   1650
                  TabIndex        =   29
                  Top             =   870
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   16
                  Left            =   60
                  TabIndex        =   28
                  Top             =   870
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   15
                  Left            =   1650
                  TabIndex        =   27
                  Top             =   540
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   15
                  Left            =   60
                  TabIndex        =   26
                  Top             =   540
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRResult 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   14
                  Left            =   1650
                  TabIndex        =   25
                  Top             =   210
                  Width           =   1215
               End
               Begin VB.TextBox txtAMRLimit 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   14
                  Left            =   60
                  TabIndex        =   24
                  Top             =   210
                  Width           =   1215
               End
               Begin BHButton.BHImageButton cmdConfirm 
                  Height          =   360
                  Index           =   7
                  Left            =   810
                  TabIndex        =   73
                  Top             =   1890
                  Visible         =   0   'False
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  Caption         =   "적용"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":1A0BA
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin BHButton.BHImageButton cmdNUnView 
                  Height          =   360
                  Left            =   1860
                  TabIndex        =   74
                  Top             =   1890
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  Caption         =   "숨김 ▶"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":1C62C
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin BHButton.BHImageButton cmdConfirm 
                  Height          =   360
                  Index           =   6
                  Left            =   60
                  TabIndex        =   77
                  Top             =   1890
                  Visible         =   0   'False
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   635
                  Caption         =   "전체적용"
                  CaptionChecked  =   "V"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TransparentPicture=   "frmTestSet.frx":1EB9E
                  PictureAlignment=   0
                  ButtonAttrib    =   2
                  ForeColor       =   16777215
                  BackColor       =   12553049
                  ImgOutLineSize  =   3
               End
               Begin VB.Label Label20 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  '단일 고정
                  Caption         =   "▶"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Left            =   1410
                  TabIndex        =   114
                  Top             =   870
                  Width           =   255
               End
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   15
               Text            =   ">="
               Top             =   1110
               Width           =   315
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   14
               Text            =   ">"
               Top             =   780
               Width           =   315
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "<="
               Top             =   450
               Width           =   315
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "<"
               Top             =   120
               Width           =   315
            End
            Begin VB.TextBox txtAMRResult 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   2820
               TabIndex        =   11
               Top             =   1110
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1410
               TabIndex        =   10
               Top             =   1110
               Width           =   1035
            End
            Begin VB.TextBox txtAMRResult 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   2820
               TabIndex        =   9
               Top             =   780
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1410
               TabIndex        =   8
               Top             =   780
               Width           =   1035
            End
            Begin VB.TextBox txtAMRResult 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   2820
               TabIndex        =   7
               Top             =   450
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1410
               TabIndex        =   6
               Top             =   450
               Width           =   1035
            End
            Begin VB.TextBox txtAMRResult 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   2820
               TabIndex        =   5
               Top             =   120
               Width           =   1125
            End
            Begin VB.TextBox txtAMRLimit 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1410
               TabIndex        =   4
               Top             =   120
               Width           =   1035
            End
            Begin BHButton.BHImageButton cmdNTypeChange 
               Height          =   420
               Left            =   1020
               TabIndex        =   71
               Top             =   1530
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   741
               Caption         =   "◀ 더 많은 수치결과변환"
               CaptionChecked  =   "V"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TransparentPicture=   "frmTestSet.frx":21110
               PictureAlignment=   0
               ButtonAttrib    =   2
               ForeColor       =   16777215
               BackColor       =   12553049
               ImgOutLineSize  =   3
            End
            Begin VB.Label Label16 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "▶"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   2460
               TabIndex        =   110
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label15 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   " 결사결과"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   30
               TabIndex        =   109
               Top             =   120
               Width           =   1035
            End
         End
      End
      Begin BHButton.BHImageButton cmdConfirm 
         Height          =   450
         Index           =   0
         Left            =   240
         TabIndex        =   79
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "저장"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmTestSet.frx":23682
         PictureAlignment=   0
         ButtonAttrib    =   2
         ForeColor       =   16777215
         BackColor       =   12553049
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdExit 
         Height          =   465
         Left            =   2400
         TabIndex        =   80
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   820
         Caption         =   "닫기"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmTestSet.frx":25450
         ButtonAttrib    =   2
         ForeColor       =   16777215
         BackColor       =   12553049
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdConfirm 
         Height          =   450
         Index           =   1
         Left            =   1320
         TabIndex        =   81
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   794
         Caption         =   "삭제"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmTestSet.frx":27346
         PictureAlignment=   0
         ButtonAttrib    =   2
         ForeColor       =   16777215
         BackColor       =   12553049
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdConfirm 
         Height          =   405
         Index           =   3
         Left            =   6450
         TabIndex        =   82
         Top             =   225
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   714
         Caption         =   "Refresh"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTestSet.frx":29114
         TransparentPicture=   "frmTestSet.frx":2926E
         ButtonAttrib    =   2
         ForeColor       =   16777215
         BackColor       =   12553049
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdConfirm 
         Height          =   420
         Index           =   2
         Left            =   5910
         TabIndex        =   83
         Top             =   225
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   741
         Caption         =   "전체저장"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTestSet.frx":2B03C
         TransparentPicture=   "frmTestSet.frx":2B196
         PictureAlignment=   0
         ButtonAttrib    =   2
         ForeColor       =   16777215
         BackColor       =   12553049
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label11 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   " 장비코드"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3960
         TabIndex        =   105
         Top             =   300
         Width           =   1035
      End
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   9555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10545
      _Version        =   393216
      _ExtentX        =   18600
      _ExtentY        =   16854
      _StockProps     =   64
      BackColorStyle  =   3
      ColsFrozen      =   6
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕코딩"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestSet.frx":2CF64
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
    Dim i As Integer
    
    With lstTestCode
        For i = 0 To .ListCount
            If txtTestCd.Text = .List(i) Then
                Exit Sub
            End If
        Next
        .AddItem txtTestCd.Text
        txtTestCd.Text = ""
    End With
    
End Sub

Private Sub cmdConfirm_Click(Index As Integer)
    Dim Test_Property       As Scripting.Dictionary
    Dim objTest_Property    As clsCommon
    Dim i                   As Integer
    Dim strTmp              As String
    Dim intINQuant          As Integer
    Dim intResUse           As Integer
    Dim strItemCodes        As String
    
    '결과표기
    If optINQuant(0).Value = True Then
        intINQuant = 0
    ElseIf optINQuant(1).Value = True Then
        intINQuant = 1      '정성(정량)
    ElseIf optINQuant(2).Value = True Then
        intINQuant = 2      '정량(정성)
    ElseIf optINQuant(3).Value = True Then
        intINQuant = 3      '정성_정량
    ElseIf optINQuant(4).Value = True Then
        intINQuant = 4      '정량_정성
    End If
    
    '결과형태
    If optResType(0).Value = True Then
        intResUse = 0       '수치
    ElseIf optResType(1).Value = True Then
        intResUse = 1       '판정결과 (문자형)
    ElseIf optResType(2).Value = True Then
        intResUse = 2       '수치/판정결과 (문자형)
    End If
    
    If Index = 1 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If


        If MsgBox(txtTestNm.Text & "를 삭제하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
             Exit Sub
        End If
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .DelTestInfo(Test_Property) Then
                '-- 삭제 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)
    '저장
    ElseIf Index = 0 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
'            txtOChannel.SetFocus
'            Exit Sub
'        End If
'
        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If

        If lstTestCode.ListCount <= 0 Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If

        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If

        'EQPMASTER 저장
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            '소수점 사용여부
            .Add "RESUSE", IIf(chkResSpec.Value = "0", "0", "1")
            '변환소수점
            .Add "RES", txtResSpec.Text
            .Add "REFML", txtRefMLow.Text
            .Add "REFMH", txtRefMHigh.Text
            .Add "REFFL", txtRefFLow.Text
            .Add "REFFH", txtRefFHigh.Text
            '결과형태 : 0:정량,1:정성,2:정량/정성
            .Add "USERESULT", intResUse
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetEqpInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        'TESTMASTER 저장
        Set Test_Property = New Scripting.Dictionary
        
        strItemCodes = ""
        For i = 0 To lstTestCode.ListCount - 1
            strItemCodes = strItemCodes & lstTestCode.List(i) & "|"
        Next
        With Test_Property
            .Add "RCH", txtRChannel.Text
            .Add "SEQ", txtSeq.Text
            .Add "TESTCD", strItemCodes
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With
        
        'AMRMASTER 저장
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "RCH", txtRChannel.Text
            .Add "AMRINRESULT", intINQuant
            '-- 결과변환 : 수치형
            .Add "AMRLIMIT1", txtAMRLimit(0).Text
            .Add "AMRLIMIT2", txtAMRLimit(1).Text
            .Add "AMRLIMIT3", txtAMRLimit(2).Text
            .Add "AMRLIMIT4", txtAMRLimit(3).Text
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT5", txtAMRLimit(4).Text
            .Add "AMRLIMIT6", txtAMRLimit(5).Text
            .Add "AMRLIMIT7", txtAMRLimit(6).Text
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT8", txtAMRLimit(7).Text
            .Add "AMRLIMIT9", txtAMRLimit(8).Text
            .Add "AMRLIMIT10", txtAMRLimit(9).Text
            .Add "AMRLIMIT11", txtAMRLimit(10).Text
            .Add "AMRLIMIT12", txtAMRLimit(11).Text
            .Add "AMRLIMIT13", txtAMRLimit(12).Text
            .Add "AMRLIMIT14", txtAMRLimit(13).Text
            '-- 결과변환 : 수치형
            .Add "AMRLIMIT15", txtAMRLimit(14).Text
            .Add "AMRLIMIT16", txtAMRLimit(15).Text
            .Add "AMRLIMIT17", txtAMRLimit(16).Text
            .Add "AMRLIMIT18", txtAMRLimit(17).Text
            .Add "AMRLIMIT19", txtAMRLimit(18).Text
            .Add "AMRRESULT1", txtAMRResult(0).Text
            .Add "AMRRESULT2", txtAMRResult(1).Text
            .Add "AMRRESULT3", txtAMRResult(2).Text
            .Add "AMRRESULT4", txtAMRResult(3).Text
            .Add "AMRRESULT5", txtAMRResult(4).Text
            .Add "AMRRESULT6", txtAMRResult(5).Text
            .Add "AMRRESULT7", txtAMRResult(6).Text
            .Add "AMRRESULT8", txtAMRResult(7).Text
            .Add "AMRRESULT9", txtAMRResult(8).Text
            .Add "AMRRESULT10", txtAMRResult(9).Text
            .Add "AMRRESULT11", txtAMRResult(10).Text
            .Add "AMRRESULT12", txtAMRResult(11).Text
            .Add "AMRRESULT13", txtAMRResult(12).Text
            .Add "AMRRESULT14", txtAMRResult(13).Text
            .Add "AMRRESULT15", txtAMRResult(14).Text
            .Add "AMRRESULT16", txtAMRResult(15).Text
            .Add "AMRRESULT17", txtAMRResult(16).Text
            .Add "AMRRESULT18", txtAMRResult(17).Text
            .Add "AMRRESULT19", txtAMRResult(18).Text
        
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetAMRInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With
        
        
        Call GetTestList
        Call GetTestMaster(spdTest)

    ElseIf Index = 2 Then
        SQL = ""
        SQL = SQL & "DELETE FROM EQPMASTER"
                
        Call DBExec(AdoCn_Local, SQL)
        
        With spdTest
            For i = 1 To .MaxRows
                SQL = ""
                SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
                SQL = SQL & "(EQUIPCD, SEQNO, SENDCHANNEL, RSLTCHANNEL"
                SQL = SQL & " , TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC "
                SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH,RESTYPE" & vbCrLf
                '-- AMR
                SQL = SQL & " , AMRLimit1, AMRResult1, AMRLimit2, AMRResult2, AMRLimit3, AMRResult3 " & vbCrLf
                SQL = SQL & " , AMRLimit4, AMRResult4, AMRLimit5, AMRResult5, AMRLimit6, AMRResult6 " & vbCrLf
                SQL = SQL & " , AMRLimit7, AMRResult7, AMRINResult)                                 " & vbCrLf
                SQL = SQL & " VALUES (" & vbCrLf
                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLOCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTNM))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLABBRNM))
                SQL = SQL & "," & GetText(spdTest, i, colLRESSPECUSE)
                SQL = SQL & "," & GetText(spdTest, i, colLRESSPEC)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLMLOW))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLMHIGH))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLFLOW))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLFHIGH))
                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE))
                '-- AMR
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 1))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 2))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 3))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 4))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 5))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 6))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 7))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE + 8))
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 3 Then
        
        Call GetTestList
        Call GetTestMaster(spdTest)
    
    ElseIf Index = 4 Then
        '문자현재코드적용
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT8", txtAMRLimit(7).Text
            .Add "AMRLIMIT9", txtAMRLimit(8).Text
            .Add "AMRLIMIT10", txtAMRLimit(9).Text
            .Add "AMRLIMIT11", txtAMRLimit(10).Text
            .Add "AMRLIMIT12", txtAMRLimit(11).Text
            .Add "AMRLIMIT13", txtAMRLimit(12).Text
            .Add "AMRLIMIT14", txtAMRLimit(13).Text
        
            .Add "AMRRESULT8", txtAMRResult(7).Text
            .Add "AMRRESULT9", txtAMRResult(8).Text
            .Add "AMRRESULT10", txtAMRResult(9).Text
            .Add "AMRRESULT11", txtAMRResult(10).Text
            .Add "AMRRESULT12", txtAMRResult(11).Text
            .Add "AMRRESULT13", txtAMRResult(12).Text
            .Add "AMRRESULT14", txtAMRResult(13).Text
            
            '-- 결과변환 : 수치형
            .Add "AMRLIMIT15", txtAMRLimit(14).Text
            .Add "AMRLIMIT16", txtAMRLimit(15).Text
            .Add "AMRLIMIT17", txtAMRLimit(16).Text
            .Add "AMRLIMIT18", txtAMRLimit(17).Text
            .Add "AMRLIMIT19", txtAMRLimit(18).Text
        
            .Add "AMRRESULT15", txtAMRResult(14).Text
            .Add "AMRRESULT16", txtAMRResult(15).Text
            .Add "AMRRESULT17", txtAMRResult(16).Text
            .Add "AMRRESULT18", txtAMRResult(17).Text
            .Add "AMRRESULT19", txtAMRResult(18).Text

        
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetAMRInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 5 Then
        '문자전체코드적용
        SQL = ""
        SQL = SQL & "DELETE FROM AMRMASTER"
                            
        Call DBExec(AdoCn_Local, SQL)
        
        With spdTest
            For i = 1 To .MaxRows
                SQL = ""
                SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
                SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE"
                SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14 " & vbCrLf
                SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14 ) " & vbCrLf
                SQL = SQL & " VALUES (" & vbCrLf
                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
                SQL = SQL & "," & STS(txtAMRLimit(7).Text)
                SQL = SQL & "," & STS(txtAMRLimit(8).Text)
                SQL = SQL & "," & STS(txtAMRLimit(9).Text)
                SQL = SQL & "," & STS(txtAMRLimit(10).Text)
                SQL = SQL & "," & STS(txtAMRLimit(11).Text)
                SQL = SQL & "," & STS(txtAMRLimit(12).Text)
                SQL = SQL & "," & STS(txtAMRLimit(13).Text)
                SQL = SQL & "," & STS(txtAMRResult(7).Text)
                SQL = SQL & "," & STS(txtAMRResult(8).Text)
                SQL = SQL & "," & STS(txtAMRResult(9).Text)
                SQL = SQL & "," & STS(txtAMRResult(10).Text)
                SQL = SQL & "," & STS(txtAMRResult(11).Text)
                SQL = SQL & "," & STS(txtAMRResult(12).Text)
                SQL = SQL & "," & STS(txtAMRResult(13).Text)
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 6 Then
        '수치전체코드적용
        SQL = ""
        SQL = SQL & "DELETE FROM AMRMASTER"
                
        Call DBExec(AdoCn_Local, SQL)
        
        With spdTest
            For i = 1 To .MaxRows
                SQL = ""
                SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
                SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE"
                SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14 " & vbCrLf
                SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14 ) " & vbCrLf
                SQL = SQL & " VALUES (" & vbCrLf
                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
                SQL = SQL & "," & STS(txtAMRLimit(7).Text)
                SQL = SQL & "," & STS(txtAMRLimit(8).Text)
                SQL = SQL & "," & STS(txtAMRLimit(9).Text)
                SQL = SQL & "," & STS(txtAMRLimit(10).Text)
                SQL = SQL & "," & STS(txtAMRLimit(11).Text)
                SQL = SQL & "," & STS(txtAMRLimit(12).Text)
                SQL = SQL & "," & STS(txtAMRLimit(13).Text)
                SQL = SQL & "," & STS(txtAMRResult(7).Text)
                SQL = SQL & "," & STS(txtAMRResult(8).Text)
                SQL = SQL & "," & STS(txtAMRResult(9).Text)
                SQL = SQL & "," & STS(txtAMRResult(10).Text)
                SQL = SQL & "," & STS(txtAMRResult(11).Text)
                SQL = SQL & "," & STS(txtAMRResult(12).Text)
                SQL = SQL & "," & STS(txtAMRResult(13).Text)
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
    
    ElseIf Index = 7 Then
        '수치현재코드적용
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT8", txtAMRLimit(7).Text
            .Add "AMRLIMIT9", txtAMRLimit(8).Text
            .Add "AMRLIMIT10", txtAMRLimit(9).Text
            .Add "AMRLIMIT11", txtAMRLimit(10).Text
            .Add "AMRLIMIT12", txtAMRLimit(11).Text
            .Add "AMRLIMIT13", txtAMRLimit(12).Text
            .Add "AMRLIMIT14", txtAMRLimit(13).Text
        
            .Add "AMRRESULT8", txtAMRResult(7).Text
            .Add "AMRRESULT9", txtAMRResult(8).Text
            .Add "AMRRESULT10", txtAMRResult(9).Text
            .Add "AMRRESULT11", txtAMRResult(10).Text
            .Add "AMRRESULT12", txtAMRResult(11).Text
            .Add "AMRRESULT13", txtAMRResult(12).Text
            .Add "AMRRESULT14", txtAMRResult(13).Text
            
            '-- 결과변환 : 수치형
            .Add "AMRLIMIT15", txtAMRLimit(14).Text
            .Add "AMRLIMIT16", txtAMRLimit(15).Text
            .Add "AMRLIMIT17", txtAMRLimit(16).Text
            .Add "AMRLIMIT18", txtAMRLimit(17).Text
            .Add "AMRLIMIT19", txtAMRLimit(18).Text
        
            .Add "AMRRESULT15", txtAMRResult(14).Text
            .Add "AMRRESULT16", txtAMRResult(15).Text
            .Add "AMRRESULT17", txtAMRResult(16).Text
            .Add "AMRRESULT18", txtAMRResult(17).Text
            .Add "AMRRESULT19", txtAMRResult(18).Text
        
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetAMRInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)
    End If
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdNTypeChange_Click()
    If fraNTypeChange.Visible = True Then
        fraNTypeChange.Visible = False
        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
    Else
        fraNTypeChange.Visible = True
        cmdNTypeChange.Caption = "▶ 수치결과변환 숨김"
        txtAMRLimit(14).SetFocus
    End If

End Sub

Private Sub cmdNUnView_Click()
    
    If fraNTypeChange.Visible = True Then
        fraNTypeChange.Visible = False
        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
    Else
        fraNTypeChange.Visible = True
        cmdNTypeChange.Caption = "▶ 수치결과변환 숨김"
    End If
    
End Sub

Private Sub cmdSave_Click()

    
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    
    With lstTestCode
        For i = 0 To .ListCount
            If txtTestCd.Text = .List(i) Then
                .RemoveItem i
                txtTestCd.Text = ""
                Exit Sub
            End If
        Next
    End With

End Sub

Private Sub cmdSeqDown_Click()

    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    
    txtSeq.Text = txtSeq.Text + 1

End Sub

Private Sub cmdSpecDown_Click()

    txtResSpec.Text = txtResSpec.Text - 1

End Sub

Private Sub cmdSpecUP_Click()
    
    If txtResSpec.Text <> "" Then
        txtResSpec.Text = txtResSpec.Text + 1
    End If
    
End Sub


Private Sub cmdTypeChange_Click()
    
    If fraTypeChange.Visible = True Then
        fraTypeChange.Visible = False
        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
    Else
        fraTypeChange.Visible = True
        cmdTypeChange.Caption = "▶ 문자결과변환 숨김"
        txtAMRLimit(7).SetFocus
    End If

End Sub


Private Sub cmdUnView_Click()
    
    If fraTypeChange.Visible = True Then
        fraTypeChange.Visible = False
        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
    Else
        fraTypeChange.Visible = True
        cmdTypeChange.Caption = "▶ 문자결과변환 숨김"
    End If
    
End Sub

Private Sub cmdView_Click()

    If fraResultTrans.Visible = False Then
        cmdView.Caption = "검사결과 변환 ▲"
        fraResultTrans.Visible = True
    Else
        cmdView.Caption = "검사결과 변환 ▼"
        fraResultTrans.Visible = False
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox("검사코드 설정화면을 닫으시겠습니까?", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    
    With spdTest
        Call SetText(spdTest, "검사약어", 0, 7):
        Call SetText(spdTest, "소수점변환", 0, 8):
        Call SetText(spdTest, "변환자릿수", 0, 9):
        Call SetText(spdTest, "남성(하한치)", 0, 10):
        Call SetText(spdTest, "남성(상한치)", 0, 11):
        Call SetText(spdTest, "여성(하한치)", 0, 12):
        Call SetText(spdTest, "여성(상한치)", 0, 13):
        Call SetText(spdTest, "결과형태", 0, 14):
        Call SetText(spdTest, "결과표기", 0, 22):
        .ColWidth(1) = 0
        .ColWidth(8) = 10
        .ColWidth(9) = 10
        .ColWidth(10) = 10
        .ColWidth(11) = 10
        .ColWidth(12) = 10
        .ColWidth(13) = 10
        .ColWidth(14) = 10
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 10
        .MaxRows = 0
    End With
    
    
    Call frmClear

    Call GetTestMaster(spdTest)
    
End Sub

Private Sub frmClear()
    Dim i As Integer
    
    
    For i = 1 To 18
        txtAMRLimit(i).Text = ""
        txtAMRResult(i).Text = ""
    Next
        
End Sub


Private Sub Form_Resize()
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'spdTest.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 160
    spdTest.HEIGHT = Me.ScaleHeight - 160
    spdTest.WIDTH = Me.ScaleWidth - frameTestSet.WIDTH - 160
    frameTestSet.LEFT = Me.ScaleWidth - frameTestSet.WIDTH
    frameTestSet.HEIGHT = spdTest.HEIGHT

End Sub

Private Sub lstTestCode_Click()
    
    txtTestCd.Text = lstTestCode.Text
    
End Sub

Private Sub optResType_Click(Index As Integer)
    
    If Index = 0 Then
        fraN.Enabled = True
        fraC.Enabled = False
        fraNC.Enabled = False
    ElseIf Index = 1 Then
        fraN.Enabled = False
        fraC.Enabled = True
        fraNC.Enabled = False
    ElseIf Index = 2 Then
        fraN.Enabled = True
        fraC.Enabled = True
        fraNC.Enabled = True
    End If
        
End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strResUse   As String
    Dim varTestCode As Variant
    Dim intCnt      As Integer
    
    varTestCode = ""
    
    If Row = 0 Then
        cmdNTypeChange.Enabled = False
        cmdTypeChange.Enabled = False
        Exit Sub
    End If

    With spdTest
        varTestCode = GetTestCode(GetText(spdTest, Row, colLRCHANNEL))
        varTestCode = Split(varTestCode, "@")
        lstTestCode.Clear
        txtTestCd.Text = ""
        If UBound(varTestCode) > 0 Then
            For intCnt = 0 To UBound(varTestCode) - 1
                lstTestCode.AddItem varTestCode(intCnt)
            Next
            txtTestCd.Text = lstTestCode.List(0)
        End If

        cmdNTypeChange.Enabled = True
        cmdTypeChange.Enabled = True
        
'        fraNTypeChange.Visible = False
'        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
'
'        fraTypeChange.Visible = False
'        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
        
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        'txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
        
        If GetText(spdTest, Row, colLRESSPECUSE) = "0" Then
            chkResSpec.Value = "0"
        Else
            chkResSpec.Value = "1"
        End If
        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
        txtRefMLow.Text = GetText(spdTest, Row, colLMLOW)
        txtRefMHigh.Text = GetText(spdTest, Row, colLMHIGH)
        txtRefFLow.Text = GetText(spdTest, Row, colLFLOW)
        txtRefFHigh.Text = GetText(spdTest, Row, colLFHIGH)
        
        strResUse = GetText(spdTest, Row, colRESTYPE)
        
'        If strResUse = "" Or strResUse = "0" Then
'            optResType(0).Value = True
'        ElseIf strResUse = "1" Then
'            optResType(1).Value = True
'        ElseIf strResUse = "2" Then
'            optResType(2).Value = True
'        End If
        If strResUse = "" Or strResUse = "수치형" Then
            optResType(0).Value = True
        ElseIf strResUse = "문자형" Then
            optResType(1).Value = True
        ElseIf strResUse = "수치/문자형" Then
            optResType(2).Value = True
        End If
        
        'AMR
        txtAMRLimit(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 1, "|")
        txtAMRResult(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 2, "|")
    
        txtAMRLimit(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 1, "|")
        txtAMRResult(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 2, "|")
    
        txtAMRLimit(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 1, "|")
        txtAMRResult(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 2, "|")
    
        txtAMRLimit(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 1, "|")
        txtAMRResult(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 2, "|")
    
        txtAMRLimit(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 1, "|")
        txtAMRResult(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 2, "|")
    
        txtAMRLimit(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 1, "|")
        txtAMRResult(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 2, "|")
    
        txtAMRLimit(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 1, "|")
        txtAMRResult(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 2, "|")
    
        If GetText(spdTest, Row, colRESTYPE + 8) = "" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "변환없음" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "판정(수치)" Then
            optINQuant(1).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "수치(판정)" Then
            optINQuant(2).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "판정 수치" Then
            optINQuant(3).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "수치 판정" Then
            optINQuant(4).Value = True
        End If
        
        
        Call frmClear
        Call GetAMRMaster(txtSeq.Text, txtRChannel.Text, txtTestCd.Text)
        
    End With
    
    'txtTestCd.SetFocus
End Sub

Private Sub txtOChannel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtRChannel.Text = txtOChannel.Text
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text <> "" Then
            Call cmdAdd_Click
        End If
    End If

End Sub

Private Sub txtTestNm_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtAbbrNm.Text = txtTestNm.Text
    End If

End Sub

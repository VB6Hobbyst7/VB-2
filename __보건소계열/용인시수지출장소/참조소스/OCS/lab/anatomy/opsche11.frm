VERSION 4.00
Begin VB.Form FrmViewIlls 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "질병코드조회"
   ClientHeight    =   8328
   ClientLeft      =   5532
   ClientTop       =   1392
   ClientWidth     =   6180
   BeginProperty Font 
      name            =   "굴림체"
      charset         =   1
      weight          =   400
      size            =   12
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Height          =   8712
   Left            =   5484
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8328
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Top             =   1056
   Width           =   6276
   Begin VBX.SpreadSheet SS1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "굴림체"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3195
      InterfaceDesigner=   "OPSCHE11.frx":0000
      Left            =   6540
      MaxCols         =   5
      MaxRows         =   1000
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Threed.SSPanel PanelSearch 
      Height          =   840
      Left            =   60
      TabIndex        =   35
      Top             =   1740
      Width           =   6072
      _version        =   65536
      _extentx        =   10716
      _extenty        =   1482
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelouter      =   1
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "ALL"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   26
         Left            =   4980
         TabIndex        =   36
         Top             =   30
         Width           =   1050
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Z"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   25
         Left            =   4560
         TabIndex        =   33
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Y"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   4185
         TabIndex        =   32
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "X"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   3810
         TabIndex        =   31
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "W"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   3435
         TabIndex        =   30
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "V"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   21
         Left            =   3060
         TabIndex        =   29
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "U"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   2685
         TabIndex        =   28
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "T"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   19
         Left            =   2310
         TabIndex        =   27
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   1935
         TabIndex        =   26
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "R"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1560
         TabIndex        =   25
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Q"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1185
         TabIndex        =   24
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "P"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   810
         TabIndex        =   23
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "O"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   435
         TabIndex        =   22
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "N"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   60
         TabIndex        =   21
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "M"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   4560
         TabIndex        =   20
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "L"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   4185
         TabIndex        =   19
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "K"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   3810
         TabIndex        =   18
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "J"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   3435
         TabIndex        =   17
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "I"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   3060
         TabIndex        =   16
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "H"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2685
         TabIndex        =   15
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "G"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   2310
         TabIndex        =   14
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "F"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1935
         TabIndex        =   13
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "E"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1560
         TabIndex        =   12
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "D"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1185
         TabIndex        =   11
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "C"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   810
         TabIndex        =   10
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "B"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   435
         TabIndex        =   9
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "A"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   360
      End
   End
   Begin Threed.SSPanel Panel 
      Height          =   372
      Left            =   60
      TabIndex        =   38
      Top             =   36
      Width           =   6072
      _version        =   65536
      _extentx        =   10716
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelouter      =   1
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   36
         Width           =   972
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&1.개인"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   1
         Left            =   1056
         TabIndex        =   2
         Top             =   36
         Width           =   972
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&2.과별"
         forecolor       =   0
         backcolor       =   16777152
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   36
         Width           =   972
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&3.전체"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   3
         Left            =   3036
         TabIndex        =   4
         Top             =   36
         Width           =   972
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&4.계통별"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   4
         Left            =   4020
         TabIndex        =   5
         Top             =   36
         Width           =   972
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&5.찾기"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   5
         Left            =   5016
         TabIndex        =   6
         Top             =   36
         Width           =   1008
         _version        =   65536
         _extentx        =   1773
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&9.종료"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   570
      TabIndex        =   34
      Top             =   4050
      Width           =   4992
   End
   Begin Threed.SSPanel PanelFind 
      Height          =   840
      Left            =   60
      TabIndex        =   40
      Top             =   2976
      Width           =   6072
      _version        =   65536
      _extentx        =   10716
      _extenty        =   1482
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelouter      =   1
      Begin VB.TextBox TxtFind 
         BeginProperty Font 
            name            =   "굴림"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   7
         Top             =   240
         Width           =   2625
      End
      Begin VB.CommandButton CmdFindOK 
         BackColor       =   &H0000FFFF&
         Caption         =   "찾기 확인"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4710
         TabIndex        =   8
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label 
         Caption         =   "찾고자하는 상병명의 단어를 입력하세요."
         BeginProperty Font 
            name            =   "굴림"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   41
         Top             =   210
         Width           =   1815
      End
   End
   Begin Threed.SSPanel PanelSet 
      Height          =   372
      Left            =   60
      TabIndex        =   43
      Top             =   456
      Width           =   6072
      _version        =   65536
      _extentx        =   10716
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelinner      =   1
      Begin Threed.SSPanel PanelSort 
         Height          =   288
         Index           =   0
         Left            =   4512
         TabIndex        =   48
         Top             =   48
         Width           =   732
         _version        =   65536
         _extentx        =   1296
         _extenty        =   503
         _stockprops     =   15
         caption         =   "코드순"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   8.4
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         floodcolor      =   0
         alignment       =   8
      End
      Begin Threed.SSPanel PanelSort 
         Height          =   288
         Index           =   1
         Left            =   5268
         TabIndex        =   49
         Top             =   48
         Width           =   732
         _version        =   65536
         _extentx        =   1296
         _extenty        =   503
         _stockprops     =   15
         caption         =   "상병순"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   8.4
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         alignment       =   8
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   288
         Left            =   48
         TabIndex        =   51
         Top             =   48
         Width           =   3480
         _version        =   65536
         _extentx        =   6138
         _extenty        =   503
         _stockprops     =   15
         caption         =   "찾고자하는 상병을 선택하세요."
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         bevelouter      =   0
      End
      Begin Threed.SSCheck ROCheck 
         Height          =   225
         Left            =   3560
         TabIndex        =   50
         Top             =   75
         Visible         =   0   'False
         Width           =   705
         _version        =   65536
         _extentx        =   1244
         _extenty        =   397
         _stockprops     =   78
         caption         =   "R/O"
         forecolor       =   4210752
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   700
            size            =   8.4
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   47
         Top             =   90
         Width           =   735
         _version        =   65536
         _extentx        =   1296
         _extenty        =   344
         _stockprops     =   78
         caption         =   "양측"
         forecolor       =   -2147483630
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   2
         Left            =   1960
         TabIndex        =   46
         Top             =   90
         Width           =   795
         _version        =   65536
         _extentx        =   1402
         _extenty        =   344
         _stockprops     =   78
         caption         =   "좌측"
         forecolor       =   -2147483630
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   45
         Top             =   90
         Width           =   795
         _version        =   65536
         _extentx        =   1402
         _extenty        =   344
         _stockprops     =   78
         caption         =   "우측"
         forecolor       =   -2147483630
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   44
         Top             =   90
         Width           =   795
         _version        =   65536
         _extentx        =   1402
         _extenty        =   344
         _stockprops     =   78
         caption         =   " None"
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   7.8
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
      End
   End
   Begin MSOutl.Outline OutlineIlls 
      Height          =   6780
      Left            =   60
      TabIndex        =   39
      Top             =   885
      Width           =   6075
      _version        =   65536
      _extentx        =   10716
      _extenty        =   11959
      _stockprops     =   77
      forecolor       =   -2147483630
      backcolor       =   12640511
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      mouseicon       =   "OPSCHE11.frx":11A40
      style           =   5
      pictureplus     =   "OPSCHE11.frx":11A5C
      pictureminus    =   "OPSCHE11.frx":11CB6
      pictureleaf     =   "OPSCHE11.frx":11F10
      pictureopen     =   "OPSCHE11.frx":1216A
      pictureclosed   =   "OPSCHE11.frx":123C4
   End
   Begin VB.Label LabelName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "굴림"
         charset         =   1
         weight          =   400
         size            =   11.4
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   60
      TabIndex        =   37
      Top             =   7740
      Width           =   6075
   End
End
Attribute VB_Name = "FrmViewIlls"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL              As String

Dim i, j, K             As Integer
Dim nSELECT             As Integer  '1.상용, 2.과상용, 3.전체상병
Dim nSET                As Integer  '0.None, 1. R/O,   2.좌측,   3.우측,   4.양측
Dim nLoadOutLine        As Integer  'OutLine View 에 상병 Load Flag
Dim nSort               As Integer  '0:코드순, 1:상병순

Dim strKorEng           As String   '한글,영문 토글

Dim nIllsIndex          As Integer
Dim GnIllSort           As Integer
Dim GstrIllSort         As String

Sub OutLineIlls_Init()
        
    Dim i               As Integer
    
    OutlineIlls.Clear
    
    For i = 1 To sS1.DataRowCnt
        sS1.Row = i
        sS1.Col = 1
        If Trim(sS1.Text) = "0" Then
            Select Case strKorEng
                Case "KOR": sS1.Col = 2     '한글명
                Case Else:  sS1.Col = 5     '영문명
            End Select
            OutlineIlls.AddItem " " & Trim(sS1.Text), -1
            OutlineIlls.ItemData(OutlineIlls.ListCount - 1) = sS1.Row
        End If
        sS1.Col = 1
        If Trim(sS1.Text) > "0" Then Exit For
    Next i
    

End Sub


Private Sub Read_SubTitle()

    Dim nIndent             As Integer
    Dim nItemData           As Integer
    Dim nListIndex          As Integer
    Dim i                   As Integer
    Dim nCNT                As Integer
    Dim strFrom             As String * 3
    Dim strTo               As String * 3
    Dim strIllName          As String * 199
    
    On Error Resume Next
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    nListIndex = OutlineIlls.ListIndex
    nIndent = OutlineIlls.Indent(nListIndex)
    nItemData = OutlineIlls.ItemData(nListIndex)
        
    If OutlineIlls.HasSubItems(nListIndex) = False Then
        Select Case nIndent
            Case 1
                GoSub Read_Indent_1
            Case 2
                GoSub Read_Indent_2
            Case 3
                GoSub Read_Indent_3
            Case 4
                Exit Sub
        End Select
    End If
    
    If OutlineIlls.Expand(nListIndex) = True Then
        Do While OutlineIlls.Indent(nListIndex + 1) = (nIndent + 1)
            OutlineIlls.RemoveItem (nListIndex + 1)
        Loop
        OutlineIlls.PictureType(nListIndex) = outClosed
    Else
        OutlineIlls.Expand(nListIndex) = True
        OutlineIlls.PictureType(nListIndex) = outOpen
    End If
    

Exit Sub


'/-------------------------------------------------------------------------------------------/

Read_Indent_1:
    
    nCNT = 0
    
    For i = 1 To sS1.DataRowCnt
        sS1.Row = i
        sS1.Col = 1
        If Val(Trim(sS1.Text)) = nItemData Then
            Select Case strKorEng
                Case "KOR": sS1.Col = 2     '한글명
                Case Else:  sS1.Col = 5     '영문명
            End Select
            OutlineIlls.AddItem " " & Trim(sS1.Text)
            nCNT = nCNT + 1
            OutlineIlls.ItemData(nListIndex + nCNT) = sS1.Row
        End If
        sS1.Col = 1
        If Val(Trim(sS1.Text)) > nItemData Then Return
    Next i
    
    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_2:

    sS1.Row = OutlineIlls.ItemData(nListIndex)
    sS1.Col = 3:    strFrom = Trim(sS1.Text)
    sS1.Col = 4:    strTo = Trim(sS1.Text)
    
    strSQL = "FOR ALL SELECT IllNameK, IllNameE, IllCode "
    strSQL = strSQL & " FROM TWBAS_ILLS "
    strSQL = strSQL & "WHERE IllClass = '1'         "
    strSQL = strSQL & "  AND IllCode >= :cIllCodeF: "
    strSQL = strSQL & "  AND IllCode <= :cIllCodeT: "
    strSQL = strSQL & "  AND SUBSTR(IllCode, 4, 1) = ' ' "
    
    Result = dosql("OPEN SCOPE")
    
    GlueSetString "cIllCodeF", 0, strFrom & "   "
    GlueSetString "cIllCodeT", 0, strTo & "ZZZ"
    
    Result = dosql(strSQL)
    
    For i = 0 To rowindicator - 1
        Select Case strKorEng
            Case "KOR": strIllName = GlueGetString("IllNameK", i)   '한글명
            Case Else:  strIllName = GlueGetString("IllNameE", i)   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & GlueGetString("IllCode", i) 'IllCode 201
    Next i
    
    Result = dosql("CLOSE SCOPE")

    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_3:

    Result = dosql("OPEN SCOPE")
    
    strFrom = MidB$(OutlineIlls.List(nListIndex), 201)
    strTo = MidB$(OutlineIlls.List(nListIndex), 201)
    GlueSetString "cIllCodeF", 0, strFrom & "   "
    GlueSetString "cIllCodeT", 0, strTo & "ZZZ"

    strSQL = "FOR ALL SELECT IllNameK, IllNameE, IllCode "
    strSQL = strSQL & " FROM TWBAS_ILLS "
    strSQL = strSQL & "WHERE IllClass = '1'         "
    strSQL = strSQL & "  AND IllCode >= :cIllCodeF: "
    strSQL = strSQL & "  AND IllCode <= :cIllCodeT: "
    Result = dosql(strSQL)
        
    For i = 0 To rowindicator - 1
        Select Case strKorEng
            Case "KOR": strIllName = GlueGetString("IllNameK", i)   '한글명
            Case Else:  strIllName = GlueGetString("IllNameE", i)   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & GlueGetString("IllCode", i)
        OutlineIlls.PictureType(nListIndex + i + 1) = outLeaf
    Next i
    
    Result = dosql("CLOSE SCOPE")
    
    Return

End Sub

Sub Read_Ills(argIndex, ArgDeptDr)

    Dim i                   As Integer
    Dim strDeptDr           As String * 6
    Dim strIllCode          As String * 8
    
    Result = Execsql("Open Scope")
    
    List1.Clear

    GoSub Option_Sql_Made
    GoSub Read_Ill
    
    Result = Execsql("Close Scope")
    
    Exit Sub
    

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    
    strDeptDr = ArgDeptDr
    If Trim(strDeptDr) = "GY" Then strDeptDr = "OB"
    GlueSetString "cDeptDr", 0, strDeptDr
    GlueSetString "cIllCode1", 0, CmdSearch(argIndex).Caption & "%"
    GlueSetString "cIllCode2", 0, LCase(CmdSearch(argIndex).Caption) & "%"
     
    Select Case nSELECT
        Case 3:
            If nSort = 0 Then
                    strSQL = "FOR ALL  SELECT IllCode, IllNameE      "
                    strSQL = strSQL & " FROM TWBAS_ILLS "
                    strSQL = strSQL & "WHERE IllClass = '1'         "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( IllCode Like :cIllCode1:  "
                    strSQL = strSQL & "   OR   IllCode Like :cIllCode2:) "
                End If
            Else
                    strSQL = "FOR ALL  SELECT IllCode, IllNameE      "
                    strSQL = strSQL & " FROM TWBAS_ILLS "
                    strSQL = strSQL & "WHERE IllClass = '1'         "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( IllNameE Like :cIllCode1:  "
                    strSQL = strSQL & "   OR   IllNameE Like :cIllCode2:) "
                End If
            End If
        Case Else
            If nSort = 0 Then
                    strSQL = "FOR ALL  SELECT A.IllCode, B.IllNameE     "
                    strSQL = strSQL & " FROM TWOCS_OILLDEF A,          "
                    strSQL = strSQL & "      TWBAS_ILLS B  "
                    strSQL = strSQL & "WHERE A.DeptDr   = :cDeptDr:    "
                    strSQL = strSQL & "  AND A.IllCode  > ' '          "
                    strSQL = strSQL & "  AND A.IllCode  = B.IllCode(+) "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( B.IllCode Like :cIllCode1:  "
                    strSQL = strSQL & "   OR   B.IllCode Like :cIllCode2:) "
                End If
            Else
                    strSQL = "FOR ALL  SELECT A.IllCode, B.IllNameE     "
                    strSQL = strSQL & " FROM TWOCS_OILLDEF A,          "
                    strSQL = strSQL & "      TWBAS_ILLS B  "
                    strSQL = strSQL & "WHERE A.DeptDr   = :cDeptDr:    "
                    strSQL = strSQL & "  AND A.IllCode  > ' '          "
                    strSQL = strSQL & "  AND A.IllCode  = B.IllCode(+) "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( B.IllNameE Like :cIllCode1:  "
                    strSQL = strSQL & "   OR   B.IllNameE Like :cIllCode2:) "
                End If
            End If
    End Select
    
    If nSort = 0 Then
        strSQL = strSQL & " ORDER BY IllCode "
    Else
        strSQL = strSQL & " ORDER BY IllNameE "
    End If
            
    Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

    Result = dosql(strSQL)
    
    For i = 0 To rowindicator - 1
        If Trim$(GlueGetString("IllNameE", i)) > "" Then
            strIllCode = GlueGetString("IllCode", i)
            List1.AddItem strIllCode & GlueGetString("IllNameE", i)
        End If
    Next i
    
    Return
    
        
End Sub


Private Sub CmdFav_Click()
    
End Sub

Private Sub CmdFindOK_Click()
    
    Dim strIllCode          As String * 8
    Dim strFind             As String
    
    strFind = Trim$(TxtFind.Text)
    
    Result = Execsql("Open Scope")
    List1.Clear
    
    GoSub Option_Sql_Made
    GoSub Read_Ill
    
    Result = Execsql("Close Scope")
    
Exit Sub
    

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    
    GlueSetString "cIllCode1", 0, "%" & UCase(LeftB$(strFind, 1)) & LCase(MidB$(strFind, 2)) & "%"
    GlueSetString "cIllCode2", 0, "%" & LCase(strFind) & "%"
    GlueSetString "cIllCode3", 0, "%" & strFind & "%"
    
    strSQL = "FOR 200  SELECT Distinct A.IllCode, IllNameE "
    strSQL = strSQL & " FROM TWOCS_OILLDEF A, TWBAS_ILLS B "
    strSQL = strSQL & "WHERE ( IllNameE Like :cIllCode1:  "
    strSQL = strSQL & "   OR   IllNameE Like :cIllCode2:  "
    strSQL = strSQL & "   OR   IllNameE Like :cIllCode3:) "
    strSQL = strSQL & "  AND   B.IllClass = '1'           "
    strSQL = strSQL & "  AND   A.IllCode  > 'A'           "
    strSQL = strSQL & "  AND   A.IllCode  = B.IllCode     "
    
    strSQL = strSQL & " ORDER BY 2 "
            
Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

    Result = dosql(strSQL)
    
    For i = 0 To rowindicator - 1
        strIllCode = GlueGetString("IllCode", i)
        List1.AddItem strIllCode & GlueGetString("IllNameE", i)
    Next i
    
Return
    
End Sub

Private Sub CmdSearch_Click(Index As Integer)

    GnIllSort = Index
    
    List1.Clear
    
    Select Case nSELECT
        Case 1:     Call Read_Ills(Index, GstrIdnumber)
        Case 2:     Call Read_Ills(Index, GstrDeptCode)
        Case 3:     Call Read_Ills(Index, " ")
    End Select
    
End Sub

Private Sub Form_Activate()
    
    nSET = 0
    'OptSets(0).Value = True
    GstrSELECTIllcode = ""
    Me.Refresh
    GstrDeptCode = FrmAppoint.cboDeptCode
    
End Sub

Private Sub Form_Load()
    
    Me.Top = 100
    Me.Left = 5610
    Me.Width = 6300
    Me.Height = 8790
    
    OutlineIlls.Top = 885
    PanelSearch.Top = 885
    PanelFind.Top = 885
    
    OutlineIlls.Left = Panel.Left
    PanelSearch.Left = Panel.Left
    PanelFind.Left = Panel.Left
    
    List1.Top = 1800
    List1.Left = OutlineIlls.Left
    List1.Width = OutlineIlls.Width
    List1.Height = 5910
    
    OutlineIlls.Visible = False
    PanelFind.Visible = False
    nLoadOutLine = False
    sS1.Visible = False
    
    If Trim(GstrDeptCode) = "IM" Or Trim(GstrDeptCode) = "CS" Then
        Me.Caption = "질병코드조회 : 개인 상용상병중 조회"
        strKorEng = "ENG"   '영문 기본
        nSELECT = 1         '과   상용
        nSET = 0            '기본 조회
    
        Call PanelMenus_Click(0)
        PanelMenus(0).BackColor = RGB(128, 255, 255)
    Else
        Me.Caption = "질병코드조회 : 과 상용상병중 조회"
        strKorEng = "ENG"   '영문 기본
        nSELECT = 2         '과   상용
        nSET = 0            '기본 조회
        
    End If
    
    GnIllSort = 0
    
    If GstrIllSort = "명순" Then
        Call PanelSort_Click(1)
    Else
        Call PanelSort_Click(0)
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set FrmViewIlls = Nothing
    
    Unload Me

End Sub

Private Sub LabelName_DblClick()
    
    If OutlineIlls.Visible = True Then
        Select Case strKorEng
            Case "KOR": strKorEng = "ENG"
            Case Else:  strKorEng = "KOR"
        End Select
        
        Call OutLineIlls_Init
    End If
    
End Sub


Private Sub List1_Click()
    
    If List1.ListIndex = -1 Then Exit Sub
    
    LabelName.Caption = List1.List(List1.ListIndex)
    
End Sub

Private Sub List1_DblClick()
    
    Dim i                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    
    If List1.ListIndex = -1 Then Exit Sub
    
    strIllCode = LeftB$(Trim(List1.List(List1.ListIndex)), 6)
    strIllNameE = Trim(MidB$(List1.List(List1.ListIndex), 9, 80))
    
    FrmViewIlls.Tag = strIllCode
    
    Me.Hide
    
    
End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call List1_DblClick

End Sub






Private Sub OptSets_Click(Index As Integer, Value As Integer)
    
    For i = 0 To 3
        OptSets(i).ForeColor = RGB(0, 0, 0)
    Next i
    
    OptSets(Index).ForeColor = RGB(0, 0, 255)
    nSET = Index
    
End Sub

Private Sub OutlineIlls_Click()

    Dim strIlls             As String * 10
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    strIlls = MidB$(OutlineIlls.List(OutlineIlls.ListIndex), 201)
    
    If Trim(strIlls) = "" Then
        strIlls = "SESSION : "
    Else
        strIlls = LeftB(strIlls, 7) & " : "
    End If
    
    LabelName.Caption = strIlls & OutlineIlls.List(OutlineIlls.ListIndex)

End Sub

Private Sub OutlineIlls_DblClick()
    
    Dim i                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    Call Read_SubTitle
    
    If OutlineIlls.PictureType(OutlineIlls.ListIndex) = outLeaf Then GoSub Data_Send: Exit Sub
    If OutlineIlls.Indent(OutlineIlls.ListIndex) > 2 Then
        If OutlineIlls.PictureType(OutlineIlls.ListIndex) = outOpen Then
            If OutlineIlls.Indent(OutlineIlls.ListIndex) = OutlineIlls.Indent(OutlineIlls.ListIndex + 1) Then
                GoSub Data_Send
            End If
        End If
    End If
    
    Exit Sub
    
    
'/--------------------------------------------------------------------------------------/

Data_Send:

    strIllCode = Trim(MidB$(OutlineIlls.List(OutlineIlls.ListIndex), 201, 6))
    
    FrmViewIlls.Tag = strIllCode
    Me.Hide
    
    Return
    
'/-------------------------------------------------------------------------------------------/
End Sub


Private Sub OutlineIlls_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Call OutlineIlls_DblClick
    
End Sub


Private Sub PanelMenus_Click(Index As Integer)
    
    Dim strIllCode              As String * 6
    Dim nYESNO                  As Integer
    
    If Index = 3 And PanelMenus(Index).BackColor = RGB(128, 255, 255) Then
        Call OutLineIlls_Init
        Exit Sub
    End If
    
    If PanelMenus(Index).BackColor = RGB(128, 255, 255) Then Exit Sub
    LabelName.Caption = ""
    
    If Index < 5 Then
        PanelMenus(0).BackColor = RGB(192, 192, 192)
        PanelMenus(1).BackColor = RGB(192, 192, 192)
        PanelMenus(2).BackColor = RGB(192, 192, 192)
        PanelMenus(3).BackColor = RGB(192, 192, 192)
        PanelMenus(4).BackColor = RGB(192, 192, 192)
        PanelMenus(Index).BackColor = RGB(128, 255, 255)
    End If
    
    Select Case Index
        Case 0: GoSub Menu_Search_1     '개인 상용상병 조회
        Case 1: GoSub Menu_Search_2     '과별 상용상병 조회
        Case 2: GoSub Menu_Search_3     '전체 상병코드 조회
        Case 3: GoSub Menu_Search_4     '계통별 상병   조회
        Case 4: GoSub Menu_Search_5     '상병 단어별   찾기
        Case 5: Me.Hide
    End Select
    
    
    Exit Sub
    
'/---------------------------------------------------------------------/
Menu_Search_1:      '개인 상용상병 조회
    nSELECT = 1
    Me.Caption = "질병코드조회 : 개인 상용상병중 조회"
'    CmdFav.Enabled = False
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    
    'GnIllSort = 26
    Call Read_Ills(GnIllSort, GstrIdnumber)
    
Return


Menu_Search_2:      '과별 상용상병 조회
    nSELECT = 2
    Me.Caption = "질병코드조회 : 과별 상용상병중 조회"
'    CmdFav.Enabled = True
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    'GnIllSort = 26
    Call Read_Ills(GnIllSort, GstrDeptCode)
Return


Menu_Search_3:      '전체 상병코드 조회
    nSELECT = 3
    Me.Caption = "질병코드조회 : 전체 상병중 조회"
'    CmdFav.Enabled = True
    CmdSearch(26).Enabled = False
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
'    GnIllSort = 0
    Call Read_Ills(GnIllSort, " ")
Return


Menu_Search_4:      '계통별 상병   조회
    Me.Caption = "질병코드조회 : 계통별 조회"
    PanelSearch.Visible = False
    OutlineIlls.Visible = True
    PanelFind.Visible = False
    List1.Visible = False
    Call OutLineIlls_Init
Return


Menu_Search_5:      '상병 단어별   찾기
    Me.Caption = "질병코드조회 : 상병 단어별 찾기"
    PanelSearch.Visible = False
    OutlineIlls.Visible = False
    PanelFind.Visible = True
    List1.Visible = True
    List1.Clear
    TxtFind.SetFocus
Return


End Sub

Private Sub PanelMenus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If PanelMenus(Index).BackColor <> RGB(128, 255, 255) Then
        PanelMenus(Index).BackColor = RGB(255, 255, 0)
    End If
    
End Sub


Private Sub PanelMenus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If PanelMenus(Index).BackColor <> RGB(128, 255, 255) Then
        PanelMenus(Index).BackColor = RGB(192, 192, 192)
    End If
    
End Sub

Private Sub PanelSort_Click(Index As Integer)

    nSort = Index
    If Index = 0 Then
        PanelSort(0).BackColor = &HFFFFC0
        PanelSort(1).BackColor = &HC0C0C0
        GstrIllSort = "코드순"
    Else
        PanelSort(1).BackColor = &HFFFFC0
        PanelSort(0).BackColor = &HC0C0C0
        GstrIllSort = "명순"
    End If
    
'    If Trim(GstrDeptCode) = "IM" Then
'        Call Read_Ills(GnIllSort, GstrIdnumber)
'    Else
        Call Read_Ills(GnIllSort, GstrDeptCode)   'GstrDrCode)
'    End If
    
End Sub

Private Sub TxtFind_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
    
End Sub



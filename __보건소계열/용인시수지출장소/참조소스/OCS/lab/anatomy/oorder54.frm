VERSION 4.00
Begin VB.Form FrmViewIlls 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "질병코드조회"
   ClientHeight    =   8385
   ClientLeft      =   3585
   ClientTop       =   435
   ClientWidth     =   8340
   ControlBox      =   0   'False
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
   Height          =   8790
   Left            =   3525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Top             =   90
   Width           =   8460
   Begin VBX.SpreadSheet SS2 
      DisplayRowHeaders=   0   'False
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "굴림체"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GridColor       =   &H00EAFFFF&
      Height          =   7425
      InterfaceDesigner=   "OORDER54.frx":0000
      Left            =   60
      MaxCols         =   2
      MaxRows         =   26
      ScrollBars      =   2  'Vertical
      SelectBlockOptions=   0
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   885
      UserResize      =   0
      Width           =   2160
   End
   Begin VBX.SpreadSheet SS1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "굴림체"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3195
      InterfaceDesigner=   "OORDER54.frx":054F
      Left            =   8880
      MaxCols         =   5
      MaxRows         =   1000
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Threed.SSPanel PanelSearch 
      Height          =   840
      Left            =   2250
      TabIndex        =   35
      Top             =   1740
      Width           =   6075
      _version        =   65536
      _extentx        =   10716
      _extenty        =   1482
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelouter      =   1
      Begin VB.CommandButton CmdFav 
         BackColor       =   &H0000FFFF&
         Caption         =   "개인등록"
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
         Left            =   4980
         TabIndex        =   39
         Top             =   420
         Width           =   1050
      End
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
      Height          =   375
      Left            =   2250
      TabIndex        =   38
      Top             =   30
      Width           =   6075
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
         Top             =   30
         Width           =   975
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&1.개인"
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
         Index           =   1
         Left            =   1050
         TabIndex        =   2
         Top             =   30
         Width           =   975
         _version        =   65536
         _extentx        =   1720
         _extenty        =   529
         _stockprops     =   15
         caption         =   "&2.과별"
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
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   30
         Width           =   975
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
         Left            =   3030
         TabIndex        =   4
         Top             =   30
         Width           =   975
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
         Top             =   30
         Width           =   975
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
         Left            =   5010
         TabIndex        =   6
         Top             =   30
         Width           =   1005
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
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   2760
      TabIndex        =   34
      Top             =   4050
      Width           =   4995
   End
   Begin Threed.SSPanel PanelFind 
      Height          =   840
      Left            =   2250
      TabIndex        =   41
      Top             =   2970
      Width           =   6075
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
         TabIndex        =   42
         Top             =   210
         Width           =   1815
      End
   End
   Begin Threed.SSPanel PanelSet 
      Height          =   375
      Left            =   2250
      TabIndex        =   44
      Top             =   450
      Width           =   6075
      _version        =   65536
      _extentx        =   10716
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483630
      borderwidth     =   1
      bevelinner      =   1
      Begin Threed.SSPanel PanelSort 
         Height          =   285
         Index           =   0
         Left            =   4515
         TabIndex        =   49
         Top             =   45
         Width           =   735
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
            size            =   9.01
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         floodcolor      =   0
         alignment       =   8
      End
      Begin Threed.SSPanel PanelSort 
         Height          =   285
         Index           =   1
         Left            =   5265
         TabIndex        =   50
         Top             =   45
         Width           =   735
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
            size            =   9.01
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         alignment       =   8
      End
      Begin Threed.SSCheck ROCheck 
         Height          =   225
         Left            =   3560
         TabIndex        =   51
         Top             =   75
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
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
      End
   End
   Begin Threed.SSPanel PanelSession 
      Height          =   795
      Left            =   60
      TabIndex        =   52
      Top             =   45
      Width           =   2145
      _version        =   65536
      _extentx        =   3784
      _extenty        =   1402
      _stockprops     =   15
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      alignment       =   6
      Begin Threed.SSCommand cmdDel 
         Height          =   330
         Left            =   90
         TabIndex        =   54
         Top             =   390
         Width           =   1965
         _version        =   65536
         _extentx        =   3466
         _extenty        =   582
         _stockprops     =   78
         caption         =   "Session내 코드삭제"
         forecolor       =   -2147483630
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         bevelwidth      =   1
         outline         =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "개인질병 Session"
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   180
         Left            =   240
         TabIndex        =   53
         Top             =   105
         Width           =   1635
      End
   End
   Begin MSOutl.Outline OutlineIlls 
      Height          =   6780
      Left            =   2250
      TabIndex        =   40
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
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      mouseicon       =   "OORDER54.frx":1203F
      style           =   5
      pictureplus     =   "OORDER54.frx":1205B
      pictureminus    =   "OORDER54.frx":122B5
      pictureleaf     =   "OORDER54.frx":1250F
      pictureopen     =   "OORDER54.frx":12769
      pictureclosed   =   "OORDER54.frx":129C3
   End
   Begin VB.Label LabelName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "굴림"
         charset         =   1
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2250
      TabIndex        =   37
      Top             =   7740
      Width           =   6075
   End
   Begin VB.Menu MenuSession 
      Caption         =   "MenuSession"
      Visible         =   0   'False
      Begin VB.Menu MenuSessionI 
         Caption         =   "Insert"
      End
      Begin VB.Menu MenuSessionLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSessionU 
         Caption         =   "Update"
      End
      Begin VB.Menu MenuSessionLine2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSessionD 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "FrmViewIlls"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim strSql              As String

Dim I, j, K             As Integer
Dim nSELECT             As Integer  '1.상용, 2.과상용, 3.전체상병
Dim nSET                As Integer  '0.None, 1. R/O,   2.좌측,   3.우측,   4.양측
Dim nLoadOutLine        As Integer  'OutLine View 에 상병 Load Flag
Dim nSort               As Integer  '0:코드순, 1:상병순

Dim strKorEng           As String   '한글,영문 토글
Dim FstrDeptDr          As String * 6
Dim FstrGbOP            As String
Dim FstrError           As String

Sub OutLineIlls_Init()
        
    Dim I               As Integer
    
    OutlineIlls.Clear
    
    For I = 1 To sS1.DataRowCnt
        sS1.Row = I
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
    Next I
    

End Sub


Private Sub Read_SubTitle()

    Dim nIndent             As Integer
    Dim nItemData           As Integer
    Dim nListIndex          As Integer
    Dim I                   As Integer
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
    
    For I = 1 To sS1.DataRowCnt
        sS1.Row = I
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
    Next I
    
    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_2:

    sS1.Row = OutlineIlls.ItemData(nListIndex)
    sS1.Col = 3:    strFrom = Trim(sS1.Text)
    sS1.Col = 4:    strTo = Trim(sS1.Text)
    
    strSql = "FOR ALL SELECT IllNameK, IllNameE, IllCode "
    strSql = strSql & " FROM TWBAS_ILLS "
    strSql = strSql & "WHERE IllClass = '" & FstrGbOP & "' "
    strSql = strSql & "  AND IllCode >= :cIllCodeF: "
    strSql = strSql & "  AND IllCode <= :cIllCodeT: "
    strSql = strSql & "  AND SUBSTR(IllCode, 4, 1) = ' ' "
    
    Result = dosql("OPEN SCOPE")
    
    GlueSetString "cIllCodeF", 0, strFrom & "   "
    GlueSetString "cIllCodeT", 0, strTo & "ZZZ"
    
    Result = dosql(strSql)
    
    For I = 0 To rowindicator - 1
        Select Case strKorEng
            Case "KOR": strIllName = GlueGetString("IllNameK", I)   '한글명
            Case Else:  strIllName = GlueGetString("IllNameE", I)   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & GlueGetString("IllCode", I) 'IllCode 201
    Next I
    
    Result = dosql("CLOSE SCOPE")

    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_3:

    Result = dosql("OPEN SCOPE")
    
    strFrom = MidB$(OutlineIlls.List(nListIndex), 201)
    strTo = MidB$(OutlineIlls.List(nListIndex), 201)
    GlueSetString "cIllCodeF", 0, strFrom & "   "
    GlueSetString "cIllCodeT", 0, strTo & "ZZZ"

    strSql = "FOR ALL SELECT IllNameK, IllNameE, IllCode "
    strSql = strSql & " FROM TWBAS_ILLS "
    strSql = strSql & "WHERE IllClass = '" & FstrGbOP & "' "
    strSql = strSql & "  AND IllCode >= :cIllCodeF: "
    strSql = strSql & "  AND IllCode <= :cIllCodeT: "
    Result = dosql(strSql)
        
    For I = 0 To rowindicator - 1
        Select Case strKorEng
            Case "KOR": strIllName = GlueGetString("IllNameK", I)   '한글명
            Case Else:  strIllName = GlueGetString("IllNameE", I)   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & GlueGetString("IllCode", I)
        OutlineIlls.PictureType(nListIndex + I + 1) = outLeaf
    Next I
    
    Result = dosql("CLOSE SCOPE")
    
    Return

End Sub

Sub Read_Ills(ArgIndex, argDeptDr)

    Dim I                   As Integer
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
    
    strDeptDr = argDeptDr
    GlueSetString "cDeptDr", 0, strDeptDr
    GlueSetString "cIllCode1", 0, CmdSearch(ArgIndex).Caption & "%"
    GlueSetString "cIllCode2", 0, LCase(CmdSearch(ArgIndex).Caption) & "%"
     
    Select Case nSELECT
        Case 3:
            If nSort = 0 Then
                    GstrSql = "FOR ALL  SELECT IllCode, IllNameE      "
                    GstrSql = GstrSql & " FROM TWBAS_ILLS "
                    GstrSql = GstrSql & "WHERE IllClass = '" & FstrGbOP & "' "
                If CmdSearch(ArgIndex).Caption <> "ALL" Then
                    GstrSql = GstrSql & "  AND ( IllCode Like :cIllCode1:  "
                    GstrSql = GstrSql & "   OR   IllCode Like :cIllCode2:) "
                End If
            Else
                    GstrSql = "FOR ALL  SELECT IllCode, IllNameE      "
                    GstrSql = GstrSql & " FROM TWBAS_ILLS "
                    GstrSql = GstrSql & "WHERE IllClass = '" & FstrGbOP & "' "
                If CmdSearch(ArgIndex).Caption <> "ALL" Then
                    GstrSql = GstrSql & "  AND ( IllNameE Like :cIllCode1:  "
                    GstrSql = GstrSql & "   OR   IllNameE Like :cIllCode2:) "
                End If
            End If
        Case Else
            If nSort = 0 Then
                    GstrSql = "FOR ALL  SELECT /*+ INDEX (TWBAS_ILLS INDEX_ILLS0) */ "
                    GstrSql = GstrSql & "      A.IllCode, B.IllNameE    "
                    GstrSql = GstrSql & " FROM TWOCS_OILLDEF A,         "
                    GstrSql = GstrSql & "      TWBAS_ILLS B  "
                    GstrSql = GstrSql & "WHERE A.DeptDr   = :cDeptDr:   "
                    GstrSql = GstrSql & "  AND A.IllCode  > ' '         "
                    GstrSql = GstrSql & "  AND A.IllCode  = B.IllCode   "
                    GstrSql = GstrSql & "  AND IllClass = '" & FstrGbOP & "' "
                If CmdSearch(ArgIndex).Caption <> "ALL" Then
                    GstrSql = GstrSql & "  AND ( B.IllCode Like :cIllCode1:  "
                    GstrSql = GstrSql & "   OR   B.IllCode Like :cIllCode2:) "
                End If
            Else
                    GstrSql = "FOR ALL  SELECT /*+ INDEX (TWBAS_ILLS INDEX_ILLS0) */ "
                    GstrSql = GstrSql & "      A.IllCode, B.IllNameE    "
                    GstrSql = GstrSql & " FROM TWOCS_OILLDEF A,         "
                    GstrSql = GstrSql & "      TWBAS_ILLS B  "
                    GstrSql = GstrSql & "WHERE A.DeptDr   = :cDeptDr:   "
                    GstrSql = GstrSql & "  AND A.IllCode  > ' '         "
                    GstrSql = GstrSql & "  AND A.IllCode  = B.IllCode   "
                    GstrSql = GstrSql & "  AND IllClass = '" & FstrGbOP & "' "
                If CmdSearch(ArgIndex).Caption <> "ALL" Then
                    GstrSql = GstrSql & "  AND ( B.IllNameE Like :cIllCode1:  "
                    GstrSql = GstrSql & "   OR   B.IllNameE Like :cIllCode2:) "
                End If
            End If
    End Select
    
    If nSort = 0 Then
        GstrSql = GstrSql & " ORDER BY IllCode "
    Else
        GstrSql = GstrSql & " ORDER BY IllNameE "
    End If
            
    Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

    Result = dosql(GstrSql)
    
    For I = 0 To rowindicator - 1
        If Trim$(GlueGetString("IllNameE", I)) > "" Then
            strIllCode = GlueGetString("IllCode", I)
            List1.AddItem strIllCode & GlueGetString("IllNameE", I) & Space(100) & "@@@@@@@@"
        End If
    Next I
    
    Return

End Sub

Private Sub cmdDel_Click()

    If List1.Visible = False Then Exit Sub
    If List1.ListIndex < 0 Then Exit Sub
    
    If Trim(RightB$(Trim(List1.List(List1.ListIndex)), 20)) = "@@@@@@@@" Then Exit Sub
    
    GstrSql = " DELETE TWOCS_ODRSLIPS "
    GstrSql = GstrSql & "    WHERE  ROWID = '" & Trim(RightB$(Trim(List1.List(List1.ListIndex)), 20)) & "' "
    
    Result = dosql(GstrSql)
    If Result = -1 Then
        Result = dosql("Rollback")
        MsgBox "Session 내의 질병코드 삭제 Error !" & Chr(13) & Chr(13) & _
               "전산실로 문의 하세요.", vbCritical, "작업 확인"
        Exit Sub
    End If
    
    Result = dosql("Commit")
    
    List1.RemoveItem List1.ListIndex
        

End Sub

Private Sub CmdFav_Click()
    Dim strIllCode          As String
    Dim nYESNO              As Integer
    
    If List1.ListIndex = -1 Then Exit Sub
    
    strIllCode = LeftB$(Trim(List1.List(List1.ListIndex)), 6)
    If Trim(strIllCode) = "" Then Exit Sub
    
    nYESNO = MsgBox("개인 코드로 등록 하시겠습니까??", _
             vbYesNo + vbDefaultButton2, "알림")
    
    If nYESNO = IDYES Then
        GstrSql = "INSERT INTO TWOCS_OILLDEF VALUES ( "
        GstrSql = GstrSql & "'" & Trim(FstrDeptDr) & "', '" & Trim(strIllCode) & "' ) "
        Result = dosql1(GstrSql)
        Result = dosql1("COMMIT")
    End If
    
End Sub

Private Sub CmdFindOK_Click()
    
    Dim strIllCode          As String * 8
    Dim strFind             As String
    
    strFind = Trim$(TxtFind.Text)
    If strFind = "" Then Exit Sub
    
    Result = Execsql("Open Scope")
    List1.Clear
    
    GoSub Option_Sql_Made
    GoSub Read_Ill
    
    Result = Execsql("Close Scope")
    
Exit Sub
    

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    
    GlueSetString "cIllCode1", 0, UCase(LeftB$(strFind, 1)) & LCase(MidB$(strFind, 2)) & "%"
    GlueSetString "cIllCode2", 0, LCase(strFind) & "%"
    GlueSetString "cIllCode3", 0, strFind & "%"
    
    GstrSql = "FOR 200  SELECT /*+ INDEX (TWBAS_ILLS INDEX_ILLS0) */ "
    GstrSql = GstrSql & "      Distinct A.IllCode, IllNameE     "
    GstrSql = GstrSql & " FROM TWOCS_OILLDEF A, TWBAS_ILLS B "
    GstrSql = GstrSql & "WHERE ( IllNameE Like :cIllCode1:  "
    GstrSql = GstrSql & "   OR   IllNameE Like :cIllCode2:  "
    GstrSql = GstrSql & "   OR   IllNameE Like :cIllCode3:) "
   'GstrSql = GstrSql & "  AND   B.IllClass = '1'           "
    GstrSql = GstrSql & "  AND   B.IllClass = '" & FstrGbOP & "' "
    GstrSql = GstrSql & "  AND   A.IllCode  > 'A'           "
    GstrSql = GstrSql & "  AND   A.IllCode  = B.IllCode     "
    
    GstrSql = GstrSql & " ORDER BY 2    "
            
Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

    Result = dosql(GstrSql)
    
    For I = 0 To rowindicator - 1
        strIllCode = GlueGetString("IllCode", I)
        List1.AddItem strIllCode & GlueGetString("IllNameE", I)
    Next I
    
Return
    
End Sub

Private Sub CmdSearch_Click(Index As Integer)

    GnIllSort = Index
    
    List1.Clear
    
    PanelSession.BackColor = RGB(192, 192, 192)
    PanelMenus(nSELECT - 1).BackColor = RGB(128, 255, 255)
    Select Case nSELECT
        Case 1:         Call Read_Ills(Index, FstrDeptDr)
        Case 2:         Call Read_Ills(Index, GstrDeptCode)
        Case 3:         Call Read_Ills(Index, " ")
    End Select
    
End Sub

Private Sub Form_Activate()

    Dim nReLoad         As Integer
    Dim nindex          As Integer
    
    nReLoad = 0
    nSET = 0
    OptSets(0).Value = True
    GstrSELECTIllcode = ""
    
    sS2.CursorStyle = SS_CURSOR_STYLE_ARROW
    
    Select Case RGB(128, 255, 255)
        Case PanelMenus(0).BackColor:   nindex = 0
        Case PanelMenus(1).BackColor:   nindex = 1
        Case PanelMenus(2).BackColor:   nindex = 2
    End Select
    
    If Me.Tag = "OPSCHE" Then
        If LeftB(Me.Caption, 4) = "수술" Then nReLoad = 1
    ElseIf Me.Tag = "OPSCHE_OP" Then
        FstrGbOP = "3"
        If nindex = 3 Then nindex = 1: nReLoad = 1
        If LeftB(Me.Caption, 4) = "질병" Then nReLoad = 1
    End If

    If nReLoad = 1 Then
        PanelMenus(nindex).BackColor = &HC0C0C0
        Call PanelMenus_Click(nindex)
    End If

    If FstrGbOP = "3" Then
        PanelSession.Enabled = False
        sS2.Enabled = False
    Else
        PanelSession.Enabled = True
        sS2.Enabled = True
    End If

    Me.Refresh

End Sub

Private Sub Form_Deactivate()

    Me.Tag = ""

End Sub

Private Sub Form_Load()
    
    Me.Top = 90
    Me.Left = 3525
    Me.Width = 8430
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
    
    Me.Caption = "질병코드조회 : 개인 상용상병중 조회"
    strKorEng = "ENG"   '영문 기본
    nSELECT = 1         '개인 상용
    nSET = 0            '기본 조회
    CmdFav.Enabled = False
    
    FstrGbOP = "1"
    If Trim(GstrDrCode_Dae) <> "" Then
        FstrDeptDr = GstrDrCode_Dae
    Else
        FstrDeptDr = GstrDrCode
    End If
    GnIllSort = 26
    
    If GstrIllSort = "명순" Then
        Call PanelSort_Click(0)
    Else
        Call PanelSort_Click(1)
    End If
    
    Call Read_DrSlips
    
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
    
    LabelName.Caption = Trim(LeftB$(List1.List(List1.ListIndex), 100))
    
End Sub

Private Sub List1_DblClick()
    
    Dim I                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    On Error GoTo GET_NOT_ORDER_FORM    'GorderFORM를 못찾을 경우(Order 화면이 아닌 경우)
    
    If List1.ListIndex = -1 Then Exit Sub
    If Left(Me.Tag, 6) = "OPSCHE" Then GoSub OPSCHE_PROC    'OP Schedule
    If GOrderFORM.SSIlls.DataRowCnt + 1 > GOrderFORM.SSIlls.MaxRows Then Exit Sub
    
    strIllCode = LeftB$(Trim(List1.List(List1.ListIndex)), 6)
    
    For I = 1 To GOrderFORM.SSIlls.DataRowCnt
        GOrderFORM.SSIlls.Row = I
        GOrderFORM.SSIlls.Col = 1
        If Trim(strIllCode) = Trim(Mid$(GOrderFORM.SSIlls.Text, 1, 4)) Then
            MsgBox "이미 상병 Code 가 있습니다", MB_OK, "추가불가"
            Exit Sub
        End If
    Next I
    
    strIllNameE = Trim(MidB$(List1.List(List1.ListIndex), 9, 80))
    strSetCode = strIllCode
    
    If Trim(MidB$(strIllCode, 5, 1)) = "" Then
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 4) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 4) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 4) & "c"
        End Select
    Else
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 5) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 5) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 5) & "c"
        End Select
    End If
    
    If strIllCode <> strSetCode Then
        GstrSql = "FOR 1 SELECT IllNameE FROM TWBAS_ILLS"
        GstrSql = GstrSql & "   WHERE IllCode = '" & strSetCode & "' "
        Result = dosql(GstrSql)
        If rowindicator > 0 Then
            strIllCode = strSetCode
            strIllNameE = GlueGetString("IllNameE", 0)
        End If
    End If
    
    GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt + 1
    GOrderFORM.SSIlls.Col = 1:  GOrderFORM.SSIlls.Text = Trim(strIllCode)
    GOrderFORM.SSIlls.Col = 2:  GOrderFORM.SSIlls.Text = Trim(strIllNameE)
    GOrderFORM.SSIlls.Col = 3:  GOrderFORM.SSIlls.Text = ""
    
    Select Case nSET
        Case 1:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OD"
        Case 2:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OS"
        Case 3:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OU"
        Case Else:  GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = ""
    End Select
    
    If ROCheck.Value = True Then
        GOrderFORM.SSIlls.Col = 5:  GOrderFORM.SSIlls.Text = "R/O"
    Else
        GOrderFORM.SSIlls.Col = 5:  GOrderFORM.SSIlls.Text = ""
    End If
    
    GOrderFORM.SSIlls.Col = 6:  GOrderFORM.SSIlls.Col2 = GOrderFORM.SSIlls.MaxCols
    GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.Row
    GOrderFORM.SSIlls.Row2 = GOrderFORM.SSIlls.Row
    GOrderFORM.SSIlls.BlockMode = True
    GOrderFORM.SSIlls.Text = ""
    GOrderFORM.SSIlls.BlockMode = False
    
    If GOrderFORM.SSIlls.DataRowCnt < GOrderFORM.SSIlls.MaxRows Then
        'GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt + 1
        GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt
        GOrderFORM.SSIlls.Col = 1
        GOrderFORM.SSIlls.Action = SS_ACTION_ACTIVE_CELL
    End If
    
    For I = 0 To 7
        GOrderFORM.PicBoowi1(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth1" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi2(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth2" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi3(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth3" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi4(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth4" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi1(I).Tag = "0"
        GOrderFORM.PicBoowi2(I).Tag = "0"
        GOrderFORM.PicBoowi3(I).Tag = "0"
        GOrderFORM.PicBoowi4(I).Tag = "0"
    Next I
    
    Exit Sub
    
    
'/-------------------------------------------------------------------------------------------/

GET_NOT_ORDER_FORM:
    
    strIllCode = LeftB$(Trim(List1.List(List1.ListIndex)), 6)
    
    If Trim(MidB$(strIllCode, 5, 1)) = "" Then
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 4) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 4) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 4) & "c"
        End Select
    Else
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 5) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 5) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 5) & "c"
        End Select
    End If
    
    GstrSELECTIllcode = strIllCode
    
    Set FrmViewIlls = Nothing
    
    Unload Me
    
'/-------------------------------------------------------------------------------------------/

OPSCHE_PROC:

    GstrSELECTIllcode = LeftB$(Trim(List1.List(List1.ListIndex)), 6)

    Set FrmViewIlls = Nothing

    Unload Me
    
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call List1_DblClick

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo File_Error

    If Button = vbRightButton Then Me.Hide: Exit Sub
    
    If List1.ListIndex < 0 Then Exit Sub
    If FstrError <> "" Then Exit Sub
    If FstrGbOP = "3" Then Exit Sub
    
    List1.DragIcon = LoadPicture(GstrDragPicPath & "PRMDrag.ico")
    List1.Drag
    
    Exit Sub
    
'----------------------------------------------------------------------------------
File_Error:

    FstrError = "ERROR"
    MsgBox "경로가 올바르지 않습니다." & vbCrLf & vbCrLf & _
           "내려 받기를 다시 실행한후에 Group작업을 하십시요." & vbCrLf & vbCrLf _
           , vbCritical, "Group 작업"

End Sub

Function Get_DrSlipsSeqNo() As Integer

    Dim I               As Integer
    
    Get_DrSlipsSeqNo = 0
    
    GstrSql = " FOR ALL "
    GstrSql = GstrSql & " SELECT SeqNo "
    GstrSql = GstrSql & "   FROM TWOCS_ODRSLIPS "
    GstrSql = GstrSql & "  WHERE DeptDr =  '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "    AND Rank   = 0 "
    GstrSql = GstrSql & "    AND Slipno = 'ILLS' "
    GstrSql = GstrSql & "  ORDER BY SeqNo   "
    
    Result = dosql(GstrSql)
    
    For I = 1 To rowindicator
        If I > GlueGetNumber("SeqNo", I - 1) Then Exit For
    Next I

    Get_DrSlipsSeqNo = I

End Function

Sub Read_DrSlips()

    Dim I               As Integer

    GstrSql = " FOR ALL "
    GstrSql = GstrSql & " SELECT HeaderName R_NAME, SeqNo R_SEQNO "
    GstrSql = GstrSql & "   FROM TWOCS_ODRSLIPS    "
    GstrSql = GstrSql & "  WHERE DeptDr =  '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "    AND Rank   = 0 "
    GstrSql = GstrSql & "    AND Slipno = 'ILLS' "
    GstrSql = GstrSql & "  ORDER BY HeaderName "

    Result = dosql(GstrSql)
    
    sS2.MaxRows = 0
    If rowindicator > 26 Then
        sS2.MaxRows = rowindicator
    Else
        sS2.MaxRows = 26
    End If

    For I = 0 To rowindicator - 1
        sS2.Row = I + 1
        sS2.Col = 1:        sS2.Text = GlueGetString("R_NAME", I)
        sS2.Col = 2:        sS2.Text = GlueGetString("R_SEQNO", I)
    Next I

End Sub

Private Sub MenuSessionD_Click()

    Dim strName         As String
    Dim nSeqno          As Integer
    Dim strDeptDr       As String * 6

    sS2.Row = sS2.ActiveRow
    sS2.Col = 1:            strName = Trim(sS2.Text)
    sS2.Col = 2:            nSeqno = Val(sS2.Text)
    
    If strName = "" Or nSeqno = 0 Then Exit Sub
    
    GstrSql = "  FOR  1  SELECT Count(*) CNT_D "
    GstrSql = GstrSql & "  FROM TWOCS_ODRSLIPS "
    GstrSql = GstrSql & " WHERE DeptDr =  '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "   AND SeqNo  =   " & nSeqno & "   "
    GstrSql = GstrSql & "   AND Slipno =  'ILLS' "
    GstrSql = GstrSql & "   AND Rank  <> 0 "

    Result = dosql(GstrSql)
    Result = MsgBox(GlueGetNumber("CNT_D", 0) & " 건의 질병코드가 있습니다." & Chr(13) & Chr(13) & _
                    strName & " Session 질병코드를 삭제하시겠습니까 ? " & Chr(13) & Chr(13), _
                    vbQuestion + vbYesNo, "Session 명칭 삭제")
    If Result = vbNo Then Exit Sub
    
    GstrSql = " DELETE TWOCS_ODRSLIPS "
    GstrSql = GstrSql & " WHERE DeptDr = '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "   AND SeqNo  =  " & nSeqno & "   "
    GstrSql = GstrSql & "   AND Slipno = 'ILLS' "
            
    Result = dosql(GstrSql)
    If Result = -1 Then
        Result = dosql("Rollback")
        MsgBox "Session 명칭 삭제 Error !" & Chr(13) & Chr(13) & _
               "전산실로 문의 하십시요."
        Exit Sub
    End If
    
    Result = dosql("Commit")

    Call Read_DrSlips


End Sub

Private Sub MenuSessionI_Click()
    
    Dim strName         As String
    Dim nSeqno          As Integer
    Dim strDeptDr       As String * 6

    strName = ""
    strName = InputBox("새로 생성할 Session 명칭을 입력 하십시요.", "Session 명칭 생성")
    
    If Trim(strName) = "" Then Exit Sub
    
    GstrSql = "  FOR  1  SELECT Count(*) CNT_I "
    GstrSql = GstrSql & "  FROM TWOCS_ODRSLIPS "
    GstrSql = GstrSql & " WHERE DeptDr     = '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "   AND HeaderName = '" & strName & "' "
    GstrSql = GstrSql & "   AND Slipno     = 'ILLS' "
    
    Result = dosql(GstrSql)
    If GlueGetNumber("CNT_I", 0) > 0 Then
        MsgBox "Session명이 존재합니다." & Chr(13) & Chr(13) & _
               strName, vbExclamation, "중복 입력"
        Exit Sub
    End If
    
    Result = dosql("Open Scope")
    
    nSeqno = Get_DrSlipsSeqNo

    GlueSetString "cDeptDr", 0, FstrDeptDr
    GlueSetnumber "cSeqNo", 0, nSeqno
    GlueSetString "cHeaderName", 0, strName
    
    GstrSql = " INSERT INTO TWOCS_ODRSLIPS VALUES "
    GstrSql = GstrSql & " (:cDeptDr:, :cSeqNo:, :cHeaderName:, '', 'ILLS', 0 ) "
    
    Result = dosql(GstrSql)
    
    If Result = -1 Then
        Result = dosql("Rollback")
        MsgBox "Session 입력 Error !" & Chr(13) & Chr(13) & _
               "전산실로 문의 하십시요."
    End If
    
    Result = dosql("Commit")
    Result = dosql("Close Scope")

    Call Read_DrSlips
    

End Sub


Private Sub MenuSessionU_Click()
    
    Dim strNameN        As String
    Dim strNameO        As String
    Dim nSeqno          As Integer
    Dim strDeptDr       As String * 6

    sS2.Row = sS2.ActiveRow
    sS2.Col = 1:            strNameO = Trim(sS2.Text)
    sS2.Col = 2:            nSeqno = Val(sS2.Text)
    
    If strNameO = "" Then Exit Sub

    strNameN = ""
    strNameN = InputBox("수정할 Session 명칭을 입력 하십시요.", "Session 명칭 수정", strNameO)
    
    If Trim(strNameN) = "" Or Trim(strNameN) = strNameO Then Exit Sub
    
    GstrSql = "  FOR  1  SELECT Count(*) CNT_U "
    GstrSql = GstrSql & "  FROM TWOCS_ODRSLIPS "
    GstrSql = GstrSql & " WHERE DeptDr     = '" & FstrDeptDr & "' "
    GstrSql = GstrSql & "   AND HeaderName = '" & strNameN & "' "
    GstrSql = GstrSql & "   AND Slipno     = 'ILLS' "
'   GstrSql = GstrSql & "   AND Rank       = 0"
    
    Result = dosql(GstrSql)
    If GlueGetNumber("CNT_U", 0) > 0 Then
        MsgBox "Session명이 존재합니다." & Chr(13) & Chr(13) & _
               strNameN, vbExclamation, "중복 입력"
        Exit Sub
    End If
    
    Result = dosql("Open Scope")
    
    GstrSql = " UPDATE TWOCS_ODRSLIPS "
    GstrSql = GstrSql & "   SET HeaderName =  '" & strNameN & "'  "
    GstrSql = GstrSql & " WHERE DeptDr     =  '" & FstrDeptDr & "'"
    GstrSql = GstrSql & "   AND SeqNo      =   " & nSeqno
    GstrSql = GstrSql & "   AND Slipno     =  'ILLS' "
    GstrSql = GstrSql & "   AND Rank       = 0 "

    Result = dosql(GstrSql)
    If Result = -1 Then
        Result = dosql("Rollback")
        MsgBox "Session 명칭 수정 Error !" & Chr(13) & Chr(13) & _
               "전산실로 문의 하세요.", vbExclamation, "Session 명칭 수정"
    End If
    
    Result = dosql("Commit")
    Result = dosql("Close Scope")

    Call Read_DrSlips

End Sub

Private Sub OptSets_Click(Index As Integer, Value As Integer)
    
    For I = 0 To 3
        OptSets(I).ForeColor = RGB(0, 0, 0)
    Next I
    
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
    
    Dim I                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    On Error GoTo GET_NOT_ORDER_FORM    'GorderFORM를 못찾을 경우(Order 화면이 아닌 경우)
    
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

    If GOrderFORM.SSIlls.DataRowCnt + 1 > GOrderFORM.SSIlls.MaxRows Then Exit Sub
    
    strIllCode = Trim(MidB$(OutlineIlls.List(OutlineIlls.ListIndex), 201))
    
    For I = 1 To GOrderFORM.SSIlls.DataRowCnt
        GOrderFORM.SSIlls.Row = I
        GOrderFORM.SSIlls.Col = 1
        If Trim(strIllCode) = Trim(Mid$(GOrderFORM.SSIlls.Text, 1, 4)) Then
            MsgBox "이미 Order Code 가 있습니다", MB_OK, "추가불가"
            Exit Sub
        End If
    Next I
    
    strIllNameE = OutlineIlls.List(OutlineIlls.ListIndex)
    strSetCode = strIllCode
    
    If Trim(MidB$(strIllCode, 5, 1)) = "" Then
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 4) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 4) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 4) & "c"
        End Select
    Else
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 5) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 5) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 5) & "c"
        End Select
    End If
    
    If strIllCode <> strSetCode Then
        GstrSql = "FOR 1 SELECT IllNameE FROM TWBAS_ILLS"
        GstrSql = GstrSql & "   WHERE IllCode = '" & strSetCode & "' "
        Result = dosql(GstrSql)
        If rowindicator > 0 Then
            strIllCode = strSetCode
            strIllNameE = GlueGetString("IllNameE", 0)
        End If
    End If
    
    GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt + 1
    GOrderFORM.SSIlls.Col = 1:  GOrderFORM.SSIlls.Text = Trim(strIllCode)
    GOrderFORM.SSIlls.Col = 2:  GOrderFORM.SSIlls.Text = Trim(strIllNameE)
    GOrderFORM.SSIlls.Col = 3:  GOrderFORM.SSIlls.Text = ""
    
    Select Case nSET
        Case 1:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OD"
        Case 2:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OS"
        Case 3:     GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = "OU"
        Case Else:  GOrderFORM.SSIlls.Col = 4:  GOrderFORM.SSIlls.Text = ""
    End Select
    
    If ROCheck.Value = True Then
        GOrderFORM.SSIlls.Col = 5:  GOrderFORM.SSIlls.Text = "R/O"
    Else
        GOrderFORM.SSIlls.Col = 5:  GOrderFORM.SSIlls.Text = ""
    End If
    
    GOrderFORM.SSIlls.Col = 6:  GOrderFORM.SSIlls.Col2 = GOrderFORM.SSIlls.MaxCols
    GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.Row
    GOrderFORM.SSIlls.Row2 = GOrderFORM.SSIlls.Row
    GOrderFORM.SSIlls.BlockMode = True
    GOrderFORM.SSIlls.Text = ""
    GOrderFORM.SSIlls.BlockMode = False
    
    If GOrderFORM.SSIlls.DataRowCnt < GOrderFORM.SSIlls.MaxRows Then
        'GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt + 1
        GOrderFORM.SSIlls.Row = GOrderFORM.SSIlls.DataRowCnt
        GOrderFORM.SSIlls.Col = 1
        GOrderFORM.SSIlls.Action = SS_ACTION_ACTIVE_CELL
    End If
    
    For I = 0 To 7
        GOrderFORM.PicBoowi1(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth1" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi2(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth2" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi3(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth3" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi4(I).Picture = LoadPicture("C:\TwHIS\Data\icons\teeth4" & Format(I, "0") & ".off")
        GOrderFORM.PicBoowi1(I).Tag = "0"
        GOrderFORM.PicBoowi2(I).Tag = "0"
        GOrderFORM.PicBoowi3(I).Tag = "0"
        GOrderFORM.PicBoowi4(I).Tag = "0"
    Next I
    
    Return
    
    
'/-------------------------------------------------------------------------------------------/

GET_NOT_ORDER_FORM:

    strIllCode = Trim(MidB$(OutlineIlls.List(OutlineIlls.ListIndex), 201))
    
    If Trim(MidB$(strIllCode, 5, 1)) = "" Then
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 4) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 4) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 4) & "c"
        End Select
    Else
        Select Case nSET
            Case 1: strSetCode = Mid$(strIllCode, 1, 5) & "b"
            Case 2: strSetCode = Mid$(strIllCode, 1, 5) & "a"
            Case 3: strSetCode = Mid$(strIllCode, 1, 5) & "c"
        End Select
    End If
    
    GstrSELECTIllcode = strIllCode
    
    Set FrmViewIlls = Nothing

    Unload Me
    
End Sub

Private Sub OutlineIlls_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Call OutlineIlls_DblClick
    
End Sub

Private Sub OutlineIlls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.Hide
        Exit Sub
    End If

End Sub

Private Sub PanelMenus_Click(Index As Integer)
    
    Dim strIllCode              As String * 6
    Dim strTitle                As String
    Dim nYESNO                  As Integer
    
    If Index = 3 And PanelMenus(Index).BackColor = RGB(128, 255, 255) Then
        Call OutLineIlls_Init
        Exit Sub
    End If
    
    If PanelMenus(Index).BackColor = RGB(128, 255, 255) Then Exit Sub
    LabelName.Caption = ""
    
    PanelSession.BackColor = RGB(192, 192, 192)
    
    If Index < 5 Then
        PanelMenus(0).BackColor = RGB(192, 192, 192)
        PanelMenus(1).BackColor = RGB(192, 192, 192)
        PanelMenus(2).BackColor = RGB(192, 192, 192)
        PanelMenus(3).BackColor = RGB(192, 192, 192)
        PanelMenus(4).BackColor = RGB(192, 192, 192)
        PanelMenus(Index).BackColor = RGB(128, 255, 255)
    End If
    
    strTitle = "질병코드조회 : "
    If Me.Tag = "OPSCHE_OP" Then strTitle = "수술코드조회 : "
    
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
    Me.Caption = strTitle & "개인 상용상병중 조회"
    CmdFav.Enabled = False
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    
    GnIllSort = 26
    Call Read_Ills(26, FstrDeptDr)
    
Return


Menu_Search_2:      '과별 상용상병 조회
    nSELECT = 2
    Me.Caption = strTitle & "과별 상용상병중 조회"
    CmdFav.Enabled = True
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    GnIllSort = 26
    Call Read_Ills(26, GstrDeptCode)
Return


Menu_Search_3:      '전체 상병코드 조회
    nSELECT = 3
    Me.Caption = strTitle & "전체 상병중 조회"
    CmdFav.Enabled = True
    CmdSearch(26).Enabled = False
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    GnIllSort = 26
    Call Read_Ills(26, " ")
Return


Menu_Search_4:      '계통별 상병   조회
    Me.Caption = strTitle & "계통별 조회"
    PanelSearch.Visible = False
    OutlineIlls.Visible = True
    PanelFind.Visible = False
    List1.Visible = False
    Call OutLineIlls_Init
Return


Menu_Search_5:      '상병 단어별   찾기
    Me.Caption = strTitle & "상병 단어별 찾기"
    PanelSearch.Visible = False
    OutlineIlls.Visible = False
    PanelFind.Visible = True
    List1.Visible = True
    List1.Clear
    TxtFind.Text = "1자 이상입력 하세요"
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
    
    If PanelSession.BackColor = RGB(128, 255, 255) Then
        If sS2.ActiveRow > 0 Then Call sS2_DblClick(1, sS2.ActiveRow)
    Else
        If nSELECT = 2 Then
            Call Read_Ills(GnIllSort, GstrDeptCode)
        Else
            Call Read_Ills(GnIllSort, FstrDeptDr)
        End If
    End If
    
End Sub

Private Sub sS2_DblClick(Col As Long, Row As Long)

    Dim I                   As Integer
    Dim strDeptDr           As String * 6
    Dim strIllCode          As String * 8
    Dim nSeqno              As Integer
    
    sS2.Col = 1
    sS2.Row = Row
    If Trim(sS2.Text) = "" Then Exit Sub
    
    sS2.Col = 2
    sS2.Row = Row:  nSeqno = Val(sS2.Text)
    
    Result = Execsql("Open Scope")
    
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    CmdFav.Enabled = True
    List1.Visible = True
    
    PanelMenus(0).BackColor = RGB(192, 192, 192)
    PanelMenus(1).BackColor = RGB(192, 192, 192)
    PanelMenus(2).BackColor = RGB(192, 192, 192)
    PanelMenus(3).BackColor = RGB(192, 192, 192)
    PanelMenus(4).BackColor = RGB(192, 192, 192)
    PanelSession.BackColor = RGB(128, 255, 255)
    
    List1.Visible = True
    List1.Clear

    GlueSetString "cDeptDr", 0, FstrDeptDr
     
    GstrSql = "FOR ALL  SELECT /*+ INDEX (TWBAS_ILLS INDEX_ILLS0) */ "
    GstrSql = GstrSql & "      IllCode, IllNameE, B.ROWID     "
    GstrSql = GstrSql & " FROM TWBAS_ILLS A, TWOCS_ODRSLIPS B "
    GstrSql = GstrSql & "WHERE IllClass = '" & FstrGbOP & "' "
   'GstrSql = GstrSql & "WHERE IllClass = '1'        "
    GstrSql = GstrSql & "  AND DeptDr   = :cDeptDr:  "
    GstrSql = GstrSql & "  AND Slipno   = 'ILLS'     "
    GstrSql = GstrSql & "  AND A.IllCode = OrderCode "
    GstrSql = GstrSql & "  AND B.SeqNo     =   " & nSeqno
    
    If nSort = 0 Then
        GstrSql = GstrSql & " ORDER BY IllCode "
    Else
        GstrSql = GstrSql & " ORDER BY IllNameE "
    End If
            
    Result = dosql(GstrSql)
    
    For I = 0 To rowindicator - 1
        If Trim$(GlueGetString("IllNameE", I)) > "" Then
            strIllCode = GlueGetString("IllCode", I)
            List1.AddItem strIllCode & GlueGetString("IllNameE", I) & Space(100) & GlueGetString("ROWID", I)
        End If
    Next I
    
    Result = Execsql("Close Scope")
    

End Sub

Private Sub sS2_DragDrop(Source As Control, X As Single, Y As Single)

    Dim strName             As String
    Dim nSeqno              As Integer
    Dim strOrderCode        As String * 8
    Dim strDeptDr           As String * 6
    
    If List1.ListIndex < 0 Then Exit Sub
    If List1.Visible = False Then Exit Sub
        
    sS2.Row = sS2.ActiveRow
    sS2.Col = 1:            strName = Trim(sS2.Text)
    sS2.Col = 2:            nSeqno = Val(sS2.Text)
    strOrderCode = Trim(LeftB$(List1.List(List1.ListIndex), 8))
    
    Result = dosql("Open Scope")
    
    GlueSetString "cDeptDr", 0, FstrDeptDr
    GlueSetnumber "cSeqNo", 0, nSeqno
    GlueSetString "cOrderCode", 0, strOrderCode
    
    GstrSql = "  FOR  1  SELECT SeqNo          "
    GstrSql = GstrSql & "  FROM TWOCS_ODRSLIPS "
    GstrSql = GstrSql & " WHERE DeptDr    = :cDeptDr:    "
    GstrSql = GstrSql & "   AND SeqNo     = :cSeqNo:     "
    GstrSql = GstrSql & "   AND OrderCode = :cOrderCode: "
    GstrSql = GstrSql & "   AND SlipNo    = 'ILLS'       "
   
    Result = dosql(GstrSql)
    
    If rowindicator = 0 Then
        GstrSql = "  FOR  1  SELECT MAX(Rank) MAX_RANK "
        GstrSql = GstrSql & "  FROM TWOCS_ODRSLIPS     "
        GstrSql = GstrSql & " WHERE DeptDr    = :cDeptDr: "
        GstrSql = GstrSql & "   AND SeqNo     = :cSeqNo:  "
        GstrSql = GstrSql & "   AND Rank     <> 0         "
        GstrSql = GstrSql & "   AND Slipno    = 'ILLS'    "
        
        Result = dosql(GstrSql)
        
        GlueSetnumber "cRank", 0, GlueGetNumber("MAX_RANK", 0) + 1

        GstrSql = " INSERT INTO TWOCS_ODRSLIPS VALUES "
        GstrSql = GstrSql & " (:cDeptDr:, :cSeqNo:, '', :cOrderCode:, 'ILLS', :cRank: ) "
        
        Result = dosql(GstrSql)
        If Result = -1 Then
            Result = dosql("Rollback")
            MsgBox "Session 질병코드 등록 Error !" & Chr(13) & Chr(13) & _
                   "전산실로 문의 하세요.", vbCritical, "Session 질병등록"
        Else
            Result = dosql("Commit")
        End If
        
    End If

    Result = dosql("Close Scope")
    

End Sub

Private Sub sS2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    Dim nCol        As Long
    Dim nRow        As Long

    Select Case State
        Case 0
            List1.DragIcon = LoadPicture(GstrDragPicPath & "PRMDrop.ico")
        Case 1
            List1.DragIcon = LoadPicture(GstrDragPicPath & "PRMDrag.ico")
    End Select

    Call SpreadGetCellFromScreenCoord(sS2, nCol, nRow, X, Y)
    
    If nRow > 0 Then
        sS2.Row = nRow
        sS2.Col = 1
        sS2.Action = SS_ACTION_ACTIVE_CELL
    End If


End Sub

Private Sub sS2_RightClick(ClickType As Integer, Col As Long, Row As Long, MouseX As Long, MouseY As Long)

    Dim strName         As String

    sS2.Col = 1
    sS2.Row = sS2.ActiveRow
    
    strName = Trim(sS2.Text)
    MenuSessionI.Caption = "Session 명칭 입력  "
    MenuSessionU.Caption = "Session 명칭 수정  " & strName
    MenuSessionD.Caption = "Session 명칭 삭제  " & strName

    PopupMenu MenuSession

End Sub

Private Sub TxtFind_GotFocus()

    TxtFind.SelStart = 0
    TxtFind.SelLength = Len(TxtFind)

End Sub

Private Sub TxtFind_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
    
End Sub


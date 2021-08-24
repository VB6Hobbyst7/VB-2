VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form frmConfig 
   Caption         =   "통신설정"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9330
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtTemp 
      Height          =   255
      Left            =   30
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1245
   End
   Begin FPSpread.vaSpread vasComList 
      Height          =   1425
      Left            =   60
      TabIndex        =   11
      Top             =   3780
      Width           =   9165
      _Version        =   393216
      _ExtentX        =   16166
      _ExtentY        =   2514
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   3
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SpreadDesigner  =   "frmConfig.frx":0442
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3045
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   9165
      _Version        =   65536
      _ExtentX        =   16166
      _ExtentY        =   5371
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.ComboBox txtGubun 
         Height          =   315
         ItemData        =   "frmConfig.frx":18EA
         Left            =   2310
         List            =   "frmConfig.frx":18F1
         TabIndex        =   22
         Top             =   780
         Width           =   2115
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확 인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7770
         TabIndex        =   21
         Top             =   120
         Width           =   1275
      End
      Begin VB.ComboBox Combo_Parity 
         Height          =   315
         ItemData        =   "frmConfig.frx":18FA
         Left            =   2325
         List            =   "frmConfig.frx":18FC
         TabIndex        =   17
         Top             =   1980
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit 
         Height          =   315
         ItemData        =   "frmConfig.frx":18FE
         Left            =   2325
         List            =   "frmConfig.frx":1900
         TabIndex        =   16
         Top             =   2370
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS 
         Height          =   315
         ItemData        =   "frmConfig.frx":1902
         Left            =   2325
         List            =   "frmConfig.frx":1904
         TabIndex        =   15
         Top             =   1590
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Port 
         Height          =   315
         ItemData        =   "frmConfig.frx":1906
         Left            =   2325
         List            =   "frmConfig.frx":1908
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkUse 
         Caption         =   "포트 사용 여부"
         Height          =   285
         Left            =   615
         TabIndex        =   13
         Top             =   270
         Width           =   2595
      End
      Begin VB.CheckBox chkDTR 
         Height          =   225
         Left            =   6330
         TabIndex        =   8
         Top             =   2415
         Width           =   345
      End
      Begin VB.CheckBox chkRTS 
         Height          =   225
         Left            =   6330
         TabIndex        =   7
         Top             =   2025
         Width           =   345
      End
      Begin VB.ComboBox Combo_Stopbit 
         Height          =   315
         Left            =   6330
         TabIndex        =   1
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "구분"
         Height          =   195
         Index           =   3
         Left            =   1005
         TabIndex        =   18
         Top             =   855
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "DTR Enabled"
         Height          =   195
         Index           =   7
         Left            =   5010
         TabIndex        =   10
         Top             =   2430
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "RTS Enabled"
         Height          =   195
         Index           =   6
         Left            =   5010
         TabIndex        =   9
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "전송속도"
         Height          =   195
         Index           =   1
         Left            =   1005
         TabIndex        =   5
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "데이터 비트"
         Height          =   195
         Index           =   2
         Left            =   1005
         TabIndex        =   4
         Top             =   2430
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "정지 비트"
         Height          =   195
         Index           =   4
         Left            =   5010
         TabIndex        =   3
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "패러티"
         Height          =   195
         Index           =   5
         Left            =   1005
         TabIndex        =   2
         Top             =   2040
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel_machine 
      Height          =   675
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   9165
      _Version        =   65536
      _ExtentX        =   16166
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "    통신설정"
      ForeColor       =   12582912
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   1
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫 기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7755
         TabIndex        =   19
         Top             =   120
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lRow, i As Long
        
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        
        Call WritePrivateProfileString("COM", "Use", CStr(chkUse.Value), App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "Port", txtGubun.Text, App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "Speed", Combo_BPS.Text, App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "Parity", Combo_Parity.Text, App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "DataBit", Combo_Databit.Text, App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "StopBit", Combo_Stopbit.Text, App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "RTSEnable", CStr(chkRTS.Value), App.Path & "\interface.ini")
        Call WritePrivateProfileString("COM", "DTREnable", CStr(chkDTR.Value), App.Path & "\interface.ini")
        
        lRow = -1
        For i = 1 To vasComList.DataRowCnt
            If Trim(GetText(vasComList, i, 3)) = "" Then
                lRow = i
                Exit For
            Else
                If Trim(GetText(vasComList, i, 3)) = Trim(Combo_Port.Text) Then
                    lRow = i
                    Exit For
                End If
            End If
        Next i
        If lRow = -1 Then lRow = vasComList.DataRowCnt + 1
        
        vasComList.Row = lRow
        vasComList.Col = 1
        vasComList.Value = chkUse.Value
        
        SetText vasComList, Trim(txtGubun.Text), lRow, 2
        SetText vasComList, Trim(Combo_Port.Text), lRow, 3
        SetText vasComList, Trim(Combo_BPS.Text), lRow, 4
        SetText vasComList, Trim(Combo_Parity.Text), lRow, 5
        SetText vasComList, Trim(Combo_Databit.Text), lRow, 6
        SetText vasComList, Trim(Combo_Stopbit.Text), lRow, 7
        SetText vasComList, CStr(chkRTS.Value), lRow, 8
        SetText vasComList, CStr(chkDTR.Value), lRow, 9
        
    End If
        
    Exit Sub
 
ErrorHandler:
    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
    
    Combo_Port.AddItem ("1")
    Combo_Port.AddItem ("2")
    Combo_Port.AddItem ("3")
    Combo_Port.AddItem ("4")
    
    Combo_BPS.AddItem ("150")
    Combo_BPS.AddItem ("300")
    Combo_BPS.AddItem ("600")
    Combo_BPS.AddItem ("1200")
    Combo_BPS.AddItem ("2400")
    Combo_BPS.AddItem ("4800")
    Combo_BPS.AddItem ("9600")
    Combo_BPS.AddItem ("14400")
    Combo_BPS.AddItem ("19200")
    
    Combo_Databit.AddItem ("7")
    Combo_Databit.AddItem ("8")
    
    Combo_Stopbit.AddItem ("1")
    Combo_Stopbit.AddItem ("1.5")
    Combo_Stopbit.AddItem ("2")
    
    Combo_Parity.AddItem ("N")
    Combo_Parity.AddItem ("E")
    Combo_Parity.AddItem ("O")
    
    lRow = lRow + 1
    
    vasComList.Row = lRow
    vasComList.Col = 1
    If Trim(txtTemp) = "1" Then
        vasComList.Value = 1
    Else
        vasComList.Value = 0
    End If
    
    db_tmp = ""
    Call GetPrivateProfileString("COM", "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 2
    
    db_tmp = ""
    Call GetPrivateProfileString("COM", "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 3
    
    db_tmp = ""
    Call GetPrivateProfileString("COM", "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 4

    db_tmp = ""
    Call GetPrivateProfileString("COM", "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 5

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 6

    db_tmp = ""
    Call GetPrivateProfileString("COM", "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 7

    db_tmp = ""
    Call GetPrivateProfileString("COM", "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 8

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    SetText vasComList, Trim(txtTemp), lRow, 9

End Sub


Private Sub vaSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub vasComList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    Dim Ret As Integer
        
    If Row < 1 Or Row > vasComList.DataRowCnt Then Exit Sub
    
    vasComList.Row = Row
    vasComList.Col = 1
    If vasComList.Value = 1 Then
        chkUse.Value = 1
    Else
        chkUse.Value = 0
    End If
    
    txtGubun = Trim(GetText(vasComList, Row, 2))
    
    Ret = -1
    For i = 0 To Combo_Port.ListCount - 1
        If Trim(GetText(vasComList, Row, 3)) = Trim(Combo_Port.List(i)) Then
            Combo_Port.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
    Ret = -1
    For i = 0 To Combo_BPS.ListCount - 1
        If Trim(GetText(vasComList, Row, 4)) = Trim(Combo_BPS.List(i)) Then
            Combo_BPS.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
    Ret = -1
    For i = 0 To Combo_Parity.ListCount - 1
        If Trim(GetText(vasComList, Row, 5)) = Trim(Combo_Parity.List(i)) Then
            Combo_Parity.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
    Ret = -1
    For i = 0 To Combo_Databit.ListCount - 1
        If Trim(GetText(vasComList, Row, 6)) = Trim(Combo_Databit.List(i)) Then
            Combo_Databit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
    Ret = -1
    For i = 0 To Combo_Stopbit.ListCount - 1
        If Trim(GetText(vasComList, Row, 7)) = Trim(Combo_Stopbit.List(i)) Then
            Combo_Stopbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
    If Trim(GetText(vasComList, Row, 8)) = True Then
        chkRTS.Value = 1
    Else
        chkRTS.Value = 0
    End If
    
    If Trim(GetText(vasComList, Row, 9)) = True Then
        chkDTR.Value = 1
    Else
        chkDTR.Value = 0
    End If

End Sub

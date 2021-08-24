VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZipCdFind 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "우편번호 찾기"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraFirst 
      BackColor       =   &H00DBE6E6&
      Height          =   3255
      Left            =   135
      TabIndex        =   3
      Top             =   660
      Width           =   6945
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Text            =   "양재동"
         Top             =   630
         Width           =   1995
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F7F3F8&
         Caption         =   "검색"
         Height          =   300
         Left            =   2205
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   615
         Width           =   600
      End
      Begin MSComctlLib.ListView lvwFind 
         Height          =   1815
         Left            =   135
         TabIndex        =   4
         Top             =   1230
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "우편    번호"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "주                              소"
            Object.Width           =   8467
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   255
         Left            =   3030
         TabIndex        =   5
         Top             =   690
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   450
         BackColor       =   14411494
         ForeColor       =   16711680
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
         Caption         =   "(예) 서울시 서초구 양재동 (X) -> 양재동(O)"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   255
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   450
         BackColor       =   14411494
         ForeColor       =   16711680
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
         Caption         =   "☞ 찾고자 하는 동/읍/면을 입력하시고 [검색] 버튼을 누르십시오."
         Appearance      =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   2
         X1              =   150
         X2              =   6735
         Y1              =   1065
         Y2              =   1080
      End
   End
   Begin VB.Frame fraSecond 
      BackColor       =   &H00DBE6E6&
      Height          =   3255
      Left            =   135
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   6945
      Begin MedControls1.LisLabel lblZipCdSecond 
         Height          =   300
         Left            =   1395
         TabIndex        =   13
         Top             =   495
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         BackColor       =   14411494
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
         Caption         =   "456"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00F7F3F8&
         Caption         =   "이전"
         Height          =   390
         Left            =   2475
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   2625
         Width           =   885
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00F7F3F8&
         Caption         =   "확인"
         Height          =   390
         Left            =   3465
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   2625
         Width           =   885
      End
      Begin VB.TextBox txtAddrNo 
         Height          =   300
         Left            =   570
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1890
         Width           =   5190
      End
      Begin MedControls1.LisLabel lblProvince 
         Height          =   300
         Left            =   570
         TabIndex        =   8
         Top             =   975
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "서울특별시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblZipCdFirst 
         Height          =   300
         Left            =   570
         TabIndex        =   9
         Top             =   495
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         BackColor       =   14411494
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
         Caption         =   "123"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDistrict 
         Height          =   300
         Left            =   2130
         TabIndex        =   10
         Top             =   975
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "서초구"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   300
         Left            =   1110
         TabIndex        =   11
         Top             =   510
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "-"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVillage 
         Height          =   300
         Left            =   3630
         TabIndex        =   12
         Top             =   975
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "중원구 상대원2동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   255
         Left            =   570
         TabIndex        =   17
         Top             =   1455
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   450
         BackColor       =   14411494
         ForeColor       =   16711680
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
         Caption         =   "☞ 나머지 동/번지/호수를 입력하십시오."
         Appearance      =   0
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "우편번호 찾기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E48372&
      Height          =   180
      Left            =   390
      TabIndex        =   1
      Top             =   255
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBF2FD&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   390
      Index           =   4
      Left            =   225
      Shape           =   4  '둥근 사각형
      Top             =   150
      Width           =   1725
   End
End
Attribute VB_Name = "frmZipCdFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objMySQL As clsZipCdFind

Private Sub cmdFind_Click()
    Call ZipFind
End Sub

Private Sub cmdOk_Click()

    'Set objMySQL = New clsZipCdFind
    With objMySQL
        .ZipCd = lblZipCdFirst.Caption & "-" & lblZipCdSecond.Caption
        .Province = lblProvince.Caption
        .District = lblDistrict.Caption
        .Village = lblVillage.Caption
        .AddrNo = txtAddrNo.Text
    End With
    
    Unload Me
End Sub

Private Sub cmdPre_Click()
    fraSecond.Visible = False
    fraFirst.Visible = True
End Sub

Private Sub Form_Load()
    txtFind.Text = ""
    lvwFind.ListItems.Clear
    lblZipCdFirst.Caption = ""
    lblZipCdSecond.Caption = ""
    lblProvince.Caption = ""
    lblDistrict.Caption = ""
    lblVillage.Caption = ""
    txtAddrNo.Text = ""
    
'    objMySQL.setDbConn ZipFindDbConn
End Sub

Private Sub lvwFind_DblClick()

    Dim strTmp As String
    Dim strZipCd As String
    Dim strProvince As String
    Dim strDistrict As String
    Dim strVillage As String
    Dim strAddrNo As String
       
    With lvwFind
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem.Selected Then
            fraFirst.Visible = False
            fraSecond.Visible = True
            
            strTmp = .SelectedItem.Text & vbTab & .ListItems(.SelectedItem.Index).Key
            
            strZipCd = medGetP(strTmp, 1, vbTab)
            strProvince = medGetP(strTmp, 2, vbTab)
            strDistrict = medGetP(strTmp, 3, vbTab)
            strVillage = medGetP(strTmp, 4, vbTab)
            strAddrNo = medGetP(strTmp, 5, vbTab)
            
            lblZipCdFirst.Caption = medGetP(strZipCd, 1, "-")
            lblZipCdSecond.Caption = medGetP(strZipCd, 2, "-")
            lblProvince.Caption = strProvince
            lblDistrict.Caption = strDistrict
            lblVillage.Caption = strVillage
            If strAddrNo <> "" Then strAddrNo = strAddrNo & Space(3)
            txtAddrNo.Text = strAddrNo
            txtAddrNo.SelStart = Len(strAddrNo)
        End If
    End With
End Sub

Private Sub txtFind_Change()
    lvwFind.ListItems.Clear
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call ZipFind
    End If
End Sub

Private Sub ZipFind()
    Dim Rs As New Recordset
    Dim itmx As ListItem
    Dim strTmp As String
    
    If Trim(txtFind.Text) = "" Then Exit Sub
       
    'Set objMySQL = New clsZipCdFind
    With objMySQL
'        .setDbConn ZipFindDbConn
        Rs.Open .GetZipCd(Trim(txtFind.Text)), dbconn
        
    End With
    
    If Rs.EOF Then
        MsgBox "존재하지 않는 지역이거나 잘못된 문장입니다.", vbInformation, "정보확인"
        With txtFind
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        Set Rs = Nothing
        'Set objMySQL = Nothing
        Exit Sub
    Else
        lvwFind.ListItems.Clear
        Do Until Rs.EOF
            strTmp = Rs.Fields("province").Value & vbTab & _
                                   Rs.Fields("district").Value & vbTab & _
                                   Rs.Fields("village").Value & vbTab & _
                                   Rs.Fields("addrno").Value
                                   
            Set itmx = lvwFind.ListItems.Add(, strTmp, Rs.Fields("zipcd").Value)
                itmx.SubItems(1) = Rs.Fields("province").Value & Space(1) & _
                                   Rs.Fields("district").Value & Space(1) & _
                                   Rs.Fields("village").Value & Space(1) & _
                                   Rs.Fields("addrno").Value
            Rs.MoveNext
        Loop
        
    End If
        
    Set Rs = Nothing
    'Set objMySQL = Nothing

End Sub

Private Sub txtFind_LostFocus()
'    Call ZipFind
End Sub

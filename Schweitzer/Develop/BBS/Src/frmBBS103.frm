VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmBBS103 
   BackColor       =   &H00DBE6E6&
   Caption         =   "����ȯ�� �ϰ� ä��"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   Icon            =   "frmBBS103.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14535
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   4830
      Left            =   8895
      ScaleHeight     =   4770
      ScaleWidth      =   5325
      TabIndex        =   25
      Top             =   3240
      Width           =   5385
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   345
         TabIndex        =   26
         Top             =   555
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin MedControls1.LisLabel lblPtCount 
         Height          =   330
         Left            =   345
         TabIndex        =   27
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin FPSpread.vaSpread tblCount 
         Height          =   4770
         Left            =   2175
         TabIndex        =   28
         Tag             =   "15109"
         Top             =   0
         Width           =   3150
         _Version        =   196608
         _ExtentX        =   5556
         _ExtentY        =   8414
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15003117
         GridColor       =   14737632
         MaxCols         =   3
         MaxRows         =   18
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS103.frx":076A
         VisibleCols     =   3
         VisibleRows     =   15
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��"
         Height          =   255
         Left            =   1620
         TabIndex        =   31
         Tag             =   "20104"
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblBuildCnt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ä����"
         Height          =   210
         Left            =   345
         TabIndex        =   30
         Tag             =   "20104"
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ȯ�ڼ�"
         Height          =   210
         Left            =   345
         TabIndex        =   29
         Tag             =   "20104"
         Top             =   1170
         Width           =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   4770
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   6600
      Left            =   300
      ScaleHeight     =   6540
      ScaleWidth      =   8295
      TabIndex        =   23
      Top             =   2205
      Width           =   8355
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6540
         Left            =   0
         TabIndex        =   24
         Tag             =   "15109"
         Top             =   0
         Width           =   8280
         _Version        =   196608
         _ExtentX        =   14605
         _ExtentY        =   11536
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15003117
         MaxCols         =   15
         MaxRows         =   25
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS103.frx":0AF6
         VisibleCols     =   3
         VisibleRows     =   25
      End
   End
   Begin VB.Frame fraOption 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Print Option"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   8880
      TabIndex        =   12
      Tag             =   "15102"
      Top             =   180
      Width           =   5355
      Begin VB.CheckBox chkPrintFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   705
         TabIndex        =   21
         Top             =   375
         Width           =   1470
      End
      Begin VB.Frame fraPrtOption 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   1485
         Left            =   630
         TabIndex        =   13
         Top             =   780
         Width           =   4215
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "���ڵ� Only"
            Height          =   330
            Index           =   1
            Left            =   300
            TabIndex        =   17
            Tag             =   "15107"
            Top             =   360
            Width           =   3210
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "���ڵ� Label && ä�� ����Ʈ"
            Height          =   330
            Index           =   0
            Left            =   300
            TabIndex        =   16
            Tag             =   "15106"
            Top             =   60
            Width           =   3210
         End
         Begin VB.TextBox txtCopy 
            Alignment       =   1  '������ ����
            Appearance      =   0  '���
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2445
            TabIndex        =   15
            Text            =   "2"
            Top             =   1050
            Width           =   525
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "ä������Ʈ Only"
            Height          =   330
            Index           =   2
            Left            =   300
            TabIndex        =   14
            Tag             =   "15107"
            Top             =   660
            Width           =   3210
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   315
            Left            =   2970
            TabIndex        =   18
            Top             =   1050
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            OrigLeft        =   3645
            OrigTop         =   1590
            OrigRight       =   3885
            OrigBottom      =   1980
            Enabled         =   -1  'True
         End
         Begin VB.Label capPrint 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "ä������Ʈ ��� ���"
            Height          =   180
            Left            =   360
            TabIndex        =   20
            Tag             =   "15105"
            Top             =   1140
            Width           =   1740
         End
         Begin VB.Label lblCopy 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3405
            TabIndex        =   19
            Tag             =   "15103"
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&S)"
      Height          =   480
      Left            =   10020
      Style           =   1  '�׷���
      TabIndex        =   11
      Tag             =   "15101"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   480
      Left            =   11505
      Style           =   1  '�׷���
      TabIndex        =   10
      Tag             =   "124"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   480
      Left            =   12945
      Style           =   1  '�׷���
      TabIndex        =   9
      Tag             =   "128"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdWardList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2475
      MousePointer    =   14  'ȭ��ǥ�� ����ǥ
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   630
      Width           =   270
   End
   Begin VB.TextBox txtWardId 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1395
      TabIndex        =   6
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdGetOrders 
      BackColor       =   &H00F4F0F2&
      Caption         =   "��ȸ(&Q)"
      Height          =   405
      Left            =   7500
      Style           =   1  '�׷���
      TabIndex        =   4
      Tag             =   "15101"
      Top             =   1140
      Width           =   1020
   End
   Begin VB.ListBox lstBuilding 
      BackColor       =   &H00F1F5F4&
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   510
      Left            =   5520
      TabIndex        =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ħä��"
         Height          =   270
         Index           =   0
         Left            =   405
         TabIndex        =   2
         Top             =   195
         Width           =   1215
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ϰ�ä��"
         Height          =   270
         Index           =   1
         Left            =   1650
         TabIndex        =   1
         Top             =   195
         Width           =   1215
      End
   End
   Begin MedControls1.LisLabel lblWardNm 
      Height          =   315
      Left            =   2820
      TabIndex        =   5
      Top             =   660
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   556
      BackColor       =   13622494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin MSComCtl2.DTPicker dtpToTime 
      Height          =   315
      Left            =   1395
      TabIndex        =   8
      Top             =   1110
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
      Format          =   24510464
      UpDown          =   -1  'True
      CurrentDate     =   36342.5951388889
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   300
      Left            =   8880
      TabIndex        =   22
      Top             =   2820
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDt 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ó����"
      Height          =   225
      Left            =   690
      TabIndex        =   34
      Tag             =   "15104"
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���� ID"
      Height          =   225
      Left            =   705
      TabIndex        =   33
      Tag             =   "15105"
      Top             =   660
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȯ�� ����Ʈ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   32
      Tag             =   "15105"
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Label lblWardLine 
      BackStyle       =   0  '����
      BorderStyle     =   1  '���� ����
      Height          =   1320
      Left            =   315
      TabIndex        =   35
      Top             =   300
      Width           =   8340
   End
End
Attribute VB_Name = "frmBBS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strBlgCd As String      '������ �ǹ� �ڵ�
Private strErbldcd As String    '�����ϰ�� �˻��� �ǹ��ڵ�
Private strGbldcd As String     '�Ϲ��ϰ�� �˻��� �ǹ��ڵ�
Private Bussdiv As String

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    lblColNm.Caption = objMyUser.EmpLngNm
    dtpToTime.value = Format(DbConn.GetSysDate, "yyyy-MM-dd  H:mm:ss")
    cmdGenerate.Enabled = False
End Sub

Private Sub UpDown1_DownClick() '����������
    txtCopy = CInt(txtCopy) - 1
    If CInt(txtCopy) < 1 Then txtCopy = 0
End Sub
Private Sub UpDown1_UpClick()   '����������
    txtCopy = CInt(txtCopy) + 1
End Sub
Private Sub chkPrintFg_Click()      '���â ó��...
    If chkPrintFg.value = 1 Then
        fraPrtOption.Enabled = False
    Else
        fraPrtOption.Enabled = True
    End If
End Sub
Private Sub cmdClear_Click()    'ȭ������
    Clear
    cmdGenerate.Enabled = False
End Sub
Private Sub cmdExit_Click()     '����
    Unload Me
End Sub
Private Sub Clear()
    txtWardId = ""
    lblWardNm.Caption = ""
    lblPtCount.Caption = ""
    tblPtList.MaxRows = 0: tblPtList.MaxRows = 20
    tblCount.MaxRows = 0: tblCount.MaxRows = 20
End Sub
Private Sub BarCode_Print(objdic As clsDictionary)
    Dim objSql As New clsGetSqlStatement
    Dim strBuildNm As String        '�ǹ��̸�
    Dim strPtid As String
    Dim strptnm As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strSpcNo As String
    Dim strAccSeq As String         'SpcYy-SpcNo ������ ��ü��ȣ
    
    objSql.setDbConn DbConn
    strBuildNm = objSql.TestBldNm(strBlgCd)
        
    objdic.MoveFirst
    
    Do Until objdic.EOF
        strPtid = medGetP(objdic.GetString, 1, COL_DIV)
        strptnm = medGetP(objdic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objdic.GetString, 3, COL_DIV)
        strColDt = medGetP(objdic.GetString, 4, COL_DIV)
        strColTm = Mid(medGetP(objdic.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        
        '��ü��ȣ ��� : 2001.2.8 �߰�
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '���ڵ� ���
        objBBSComCode.BarInfo.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                            strptnm, "", "", "", txtWardId, _
                                            strColDt, strColTm, "", CLng(txtCopy)
        objdic.MoveNext
    Loop
    
    'Form Feed : 2001.2.8 �߰�
    objBBSComCode.BarInfo.Label_FormFeed
    Set objSql = Nothing
        
End Sub
Private Sub ColList_Print()
'ä������Ʈ ������
End Sub
Private Function Redim_Ary() As Long
'���ڵ� ��½� �迭�� ������ ���Ѵ�.
    Dim ii As Integer
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii: .Col = 1
            If .value = 0 Then
                Redim_Ary = Redim_Ary + 1
            End If
        Next
    End With
End Function
Private Sub cmdGenerate_Click() '����ä�� ����

    Dim strPtid As String       'ȯ��id
    Dim strptnm As String       'ȯ�ڸ�
    Dim strColID As String      'ä����
    Dim strColDt As String      'ä����
    Dim strColTm As String      'ä���Ͻ�
    Dim lngErCnt As Long
    Dim lngGcnt As Long
    
    Dim ii As Long
    
    If Redim_Ary = 0 Then Exit Sub
    strColID = objMyUser.EmpId
    
    Dim objCollect As New clsSpcAddPaper
    Dim objdic     As New clsDictionary
    
    objCollect.setDbConn DbConn
    
    objdic.Clear
    objdic.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd"
    
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1
            If .value = "0" Then
                .Col = 3: strPtid = .value
                .Col = 4: strptnm = .value
                .Col = 5
                If .value = "����" Then
                    lngErCnt = lngErCnt + 1
                Else
                    lngGcnt = lngGcnt + 1
                End If
                .Col = 7:  strColDt = Format(.Text, "YYYYMMDD")
                .Col = 8:  strColTm = Format(.Text, "HHMMss")
                objdic.AddNew strPtid, Join(Array(strptnm, strColDt, strColTm, strColID, BBSBUSSDIV.stsBed, strBlgCd), COL_DIV)
            End If
        Next
    End With
    If objdic.RecordCount > 0 Then
        If objCollect.Set_Collect(objdic) Then
            With tblCount
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = 1
                    If .value = strErbldcd Then
                        .Col = 3: .value = lngErCnt
                    ElseIf .value = strGbldcd Then
                        .Col = 3: .value = lngGcnt
                    ElseIf .value = "" Then
                        Exit For
                    End If
                Next
                lblPtCount.Caption = lngErCnt + lngGcnt
            End With
            Dim objBar As New clsDictionary
            
            Set objBar = objCollect.BldDic
            If objBar.RecordCount > 0 Then
                BarCode_Print objBar
            Else
                MsgBox "��ü�� �̹� �����ϹǷ� ���ڵ尡 ��µ��� �ʽ��ϴ�.", vbInformation + vbOKOnly, "���ڵ����"
            End If
            cmdGenerate.Enabled = False
        End If
    End If
    Set objCollect = Nothing
    Set objdic = Nothing
    Set objBar = Nothing
    
 
    
End Sub
Private Sub TestBuilding_Search()
    Dim objSql As New clsGetSqlStatement
    Dim strTmp As String
    
    objSql.setDbConn DbConn
    
    With objSql
        If txtWardId = "" Then
            strBlgCd = objSysInfo.BuildingCd
        Else
            strBlgCd = .Get_BuildingCd(UCase(txtWardId))
        End If
        strTmp = .TestBuildCd(strBlgCd)
        strErbldcd = medGetP(strTmp, 1, COL_DIV)
        strGbldcd = medGetP(strTmp, 2, COL_DIV)
    End With
    
    With tblCount
        .Row = 1: .Col = 1: .value = strErbldcd
        .Row = 1: .Col = 2: .value = objSql.TestBldNm(strErbldcd)
        .Row = 2: .Col = 1: .value = strGbldcd
        .Row = 2: .Col = 2: .value = objSql.TestBldNm(strGbldcd)
    End With
    
    Set objSql = Nothing
End Sub
Private Sub cmdGetOrders_Click()
    '������ ä������� ��ȸ
    'ó�����̺�(lab101)���� BussDiv=B ,DoneFg=0 �ΰ� ��ȸ�ؿ´�.
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New DrRecordSet
    Dim strErChk As String
    Dim strOrdDt As String
    Dim strPtid As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strOrdNo As String
    Dim blnSearch As Boolean
    Dim i As Integer
    
    blnSearch = True
    strOrdDt = Format(dtpToTime.value, "yyyyMMdd")
    strColDt = Format(DbConn.GetSysDate, "yyyy-mm-dd")
    strColTm = Format(DbConn.GetSysDate, "HH:mm")
    If txtWardId = "" Then
        MsgBox "������ �Է����� ��ȸ�Ͻʽÿ�.", vbInformation + vbOKOnly, Me.Caption
        txtWardId.SetFocus
        Exit Sub
    End If
    TestBuilding_Search
    i = 1
    
    objGetSql.setDbConn DbConn
    
    Set DrRS = objGetSql.Get_ORDER_103(strOrdDt, UCase(txtWardId))
    
    If Not DrRS.EOF = True Then
        Do Until DrRS.EOF = True
            With tblPtList
                .MaxRows = i
                .Row = .MaxRows
                .Col = 2:  .value = lblWardNm.Caption
                .Col = 3:  .value = DrRS.Fields("ptid").value: strPtid = Trim(.value)
                .Col = 4:  .value = DrRS.Fields("ptnm").value
                strErChk = objGetSql.ER_Chk(strPtid, strOrdDt)
                .Col = 5:  .value = IIf(strErChk = "1", "����", "�Ϲ�")
                If objGetSql.Blood_Existence(strPtid, Format(DbConn.GetSysDate, "yyyyMMdd"), Format(DbConn.GetSysDate, "HHmm")) = True Then
                    .Col = 6: .value = "�ű԰�ü"
                Else
                    .Col = 6: .value = "��ü����"
                End If
                .Col = 7:  .Text = strColDt
                .Col = 8:  .Text = strColTm
                .Col = 9:  .value = strOrdDt
                .Col = 10: .value = IIf(strErChk = "1", strErbldcd, strGbldcd)
                .Col = 11: .value = DrRS.Fields("bedindt").value
                .Col = 12: .value = DrRS.Fields("bussdiv").value
                .Col = 13: .value = DrRS.Fields("reqdt").value
                i = i + 1
            End With
            DrRS.MoveNext
        Loop
    Else
        blnSearch = False
        tblPtList.MaxRows = 0
    End If
    
    If Get_SpcAdd(strOrdDt, txtWardId) = False And blnSearch = False Then
        MsgBox "���ǿ� �ش�Ǵ� ó�渮��Ʈ�� �����ϴ�.Ȯ���� ó���ϼ���.", vbInformation + vbOKOnly, Me.Caption
        cmdGenerate.Enabled = False
        tblPtList.MaxRows = 0: tblPtList.MaxRows = 25
    Else
        cmdGenerate.Enabled = True
    End If
    DrRS.RsClose:   Set DrRS = Nothing
    Set objGetSql = Nothing
    
End Sub
Private Function Get_SpcAdd(ByVal orddt As String, wardid As String) As Boolean
'���������� ä��������߿� ��ü �߰� ����ڰ� ���ԵǾ� �ִ��� �Ǵ��ؼ� �����ش�.
'��ü �߰� ����ڴ� �̹� ������ ȯ�ڸ� �������� �ҷ��´�.
'�߰���û���� ������ ���� ��¥�� �������� �۰ų� ���� �͸��� ������� �Ѵ�.
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New DrRecordSet
    Dim strErChk As String
    Dim strPtid As String
    Dim strColDt As String
    Dim strColTm As String
    Dim cnt As Integer
    
    Get_SpcAdd = True
    strColDt = Format(DbConn.GetSysDate, "yyyy-mm-dd")
    strColTm = Format(DbConn.GetSysDate, "HH:mm")

    
    objGetSql.setDbConn DbConn
    
    Set DrRS = objGetSql.Get_SpcAdd(UCase(wardid))
    
    If Not DrRS.EOF Then
        With tblPtList
            Do Until DrRS.EOF
                If DupCheck(DrRS.Fields("ptid").value) = False Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .ForeColor = vbBlue
                    .Col = 2: .value = lblWardNm.Caption
                    .Col = 3: .value = DrRS.Fields("ptid").value: strPtid = Trim(.value)
                    .Col = 4: .value = DrRS.Fields("ptnm").value
                    strErChk = objGetSql.ER_Chk(strPtid, DrRS.Fields("orddt").value)
                    .Col = 5: .value = IIf(strErChk = "1", "����", "�Ϲ�")
                    .Col = 6: .value = "�߰���û"
                    .Col = 7: .Text = strColDt
                    .Col = 8: .Text = strColTm
                    .Col = 9: .value = DrRS.Fields("orddt").value
                    .Col = 10: .value = IIf(strErChk = "1", strErbldcd, strGbldcd)
                    .Col = 11: .value = DrRS.Fields("bedindt").value
                    .Col = 12: .value = DrRS.Fields("bussdiv").value
                    .Col = 13: .value = DrRS.Fields("reqdt").value
                    cnt = cnt + 1
                Else
                    '�߰�ä����, �Ϲ�ä���� ���ÿ� �߻��Ѱ��
                    .Col = 14: .value = "*"
                End If
                DrRS.MoveNext
            Loop
        End With
    Else
        Get_SpcAdd = False
    End If
    
    If cnt = 0 Then Get_SpcAdd = False
    
    Set objGetSql = Nothing

End Function
Private Function DupCheck(ByVal pBldNo As String) As Boolean
'�ߺ����� üũ�Ѵ�.

    Dim strClip As String
    
    With tblPtList
        .Row = 1: .Row2 = .MaxRows
        .Col = 3: .Col2 = 3
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False
        
        If InStr(strClip, pBldNo) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With
    
End Function

Private Sub cmdWardList_Click()

    Dim objLPF As New clsListPopFactory
    Dim SelString As String
    
    objLPF.ListType = TypeWard
    objLPF.ShowListPop
    SelString = objLPF.SelString
    If SelString <> "" Then
        txtWardId = medGetP(SelString, 1, ";")
        lblWardNm.Caption = medGetP(SelString, 2, ";")
    End If
    
    Set objLPF = Nothing
End Sub
Private Sub txtWardId_LostFocus()
    If txtWardId = "" Then
        lblWardNm.Caption = ""
    Else
        Search_Ward
    End If
    
End Sub
Private Sub txtWardID_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Search_Ward
    End If
End Sub
Private Sub Search_Ward()
    Dim objSql As New clsGetSqlStatement
    Dim Rs     As New DrRecordSet
    
    With objSql
        .setDbConn DbConn
        Set Rs = objSql.Get_WardDept_Pop(UCase(txtWardId))
        If Not Rs.EOF Then
            txtWardId = Rs.Fields("deptcd").value
            lblWardNm.Caption = Rs.Fields("deptnm").value
        Else
            MsgBox "�ش�Ǵ� �ڷᰡ �����ϴ�. Ȯ���� �Է��ϼ���.", vbInformation + vbOKOnly, "�����Է�"
            txtWardId = ""
            lblWardNm.Caption = ""
        End If
    End With

    Set Rs = Nothing
    Set objSql = Nothing
    
End Sub


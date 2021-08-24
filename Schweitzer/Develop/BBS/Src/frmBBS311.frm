VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS311 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Local 병원 의뢰 혈액"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   -225
   ClientWidth     =   14685
   Icon            =   "frmBBS311.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   14685
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   1530
      ScaleHeight     =   495
      ScaleWidth      =   10860
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   750
      Width           =   10920
      Begin VB.TextBox txtLocalCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdLocalCd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   120
         Width           =   300
      End
      Begin MedControls1.LisLabel lblLocalNm 
         Height          =   315
         Left            =   3345
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Local 병원"
         Height          =   180
         Left            =   360
         TabIndex        =   24
         Top             =   180
         Width           =   885
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   1500
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   420
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   14411494
      TabCaption(0)   =   "혈액 출고"
      TabPicture(0)   =   "frmBBS311.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "출고 혈액 조회"
      TabPicture(1)   =   "frmBBS311.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   6735
         Left            =   -74970
         TabIndex        =   19
         Top             =   810
         Width           =   10905
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00F4F0F2&
            Caption         =   "조회(&Q)"
            Height          =   420
            Left            =   4980
            Style           =   1  '그래픽
            TabIndex        =   11
            Tag             =   "128"
            Top             =   600
            Width           =   1230
         End
         Begin MSComCtl2.DTPicker dtpFr 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   660
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   62849027
            CurrentDate     =   36859
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   3060
            TabIndex        =   10
            Top             =   660
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   62849027
            CurrentDate     =   36859
         End
         Begin FPSpread.vaSpread tblResult 
            Height          =   4905
            Left            =   270
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1485
            Width           =   10425
            _Version        =   196608
            _ExtentX        =   18389
            _ExtentY        =   8652
            _StockProps     =   64
            BackColorStyle  =   1
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   14411494
            GridShowVert    =   0   'False
            MaxCols         =   7
            MaxRows         =   20
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            ShadowText      =   0
            SpreadDesigner  =   "frmBBS311.frx":07A2
            StartingColNumber=   0
            UserResize      =   0
            TextTip         =   4
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "출고일자"
            Height          =   180
            Left            =   480
            TabIndex        =   21
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   180
            Left            =   2820
            TabIndex        =   20
            Top             =   720
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   6735
         Left            =   30
         TabIndex        =   15
         Top             =   810
         Width           =   10935
         Begin VB.TextBox txtABO 
            Height          =   330
            Left            =   7410
            ScrollBars      =   2  '수직
            TabIndex        =   27
            Top             =   1080
            Width           =   1800
         End
         Begin VB.TextBox txtPtNm 
            Height          =   330
            Left            =   7410
            ScrollBars      =   2  '수직
            TabIndex        =   25
            Top             =   735
            Width           =   1800
         End
         Begin VB.CommandButton cmdDelivery 
            BackColor       =   &H00F4F0F2&
            Caption         =   "출고"
            Height          =   420
            Left            =   9660
            Style           =   1  '그래픽
            TabIndex        =   8
            Tag             =   "128"
            Top             =   990
            Width           =   1110
         End
         Begin VB.TextBox txtBldNo 
            Alignment       =   2  '가운데 맞춤
            Height          =   300
            Left            =   1125
            MaxLength       =   12
            TabIndex        =   3
            Text            =   "20-99-042006"
            Top             =   750
            Width           =   1455
         End
         Begin VB.ComboBox cboCompo 
            Height          =   300
            ItemData        =   "frmBBS311.frx":0CFF
            Left            =   3990
            List            =   "frmBBS311.frx":0D01
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   750
            Width           =   2430
         End
         Begin VB.CheckBox chkBar 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드입력"
            Height          =   195
            Left            =   420
            TabIndex        =   2
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox txtRemark 
            Height          =   330
            Left            =   1125
            ScrollBars      =   2  '수직
            TabIndex        =   5
            Top             =   1080
            Width           =   5295
         End
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00F4F0F2&
            Caption         =   "적용"
            Height          =   420
            Left            =   9660
            Style           =   1  '그래픽
            TabIndex        =   6
            Tag             =   "128"
            Top             =   540
            Width           =   1110
         End
         Begin FPSpread.vaSpread tblDelivery 
            Height          =   4905
            Left            =   360
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1500
            Width           =   10425
            _Version        =   196608
            _ExtentX        =   18389
            _ExtentY        =   8652
            _StockProps     =   64
            BackColorStyle  =   1
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   14411494
            GridShowVert    =   0   'False
            MaxCols         =   9
            MaxRows         =   20
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            ShadowText      =   0
            SpreadDesigner  =   "frmBBS311.frx":0D03
            StartingColNumber=   0
            UserResize      =   0
            TextTip         =   4
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "혈액형"
            Height          =   180
            Left            =   6825
            TabIndex        =   28
            Top             =   1170
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "환자명"
            Height          =   180
            Left            =   6825
            TabIndex        =   26
            Top             =   825
            Width           =   540
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "혈액번호"
            Height          =   210
            Left            =   390
            TabIndex        =   18
            Top             =   810
            Width           =   720
         End
         Begin VB.Label Label17 
            BackStyle       =   0  '투명
            Caption         =   "제제"
            Height          =   210
            Left            =   3540
            TabIndex        =   17
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Remark"
            Height          =   180
            Left            =   420
            TabIndex        =   16
            Top             =   1140
            Width           =   645
         End
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9840
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "128"
      Top             =   8340
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   11145
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "128"
      Top             =   8340
      Width           =   1320
   End
End
Attribute VB_Name = "frmBBS311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'혈액출고
'Coding By Legends

Private objMySql As clsBBSSQLStatement
'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private Function BldDupCheck(ByVal pFindStr As String) As Boolean
    Dim strTmp As String
    
    With tblDelivery
        .Col = 6: .COL2 = 6: .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        strTmp = .ClipValue
        .BlockMode = False
    End With
    
    If InStr(strTmp, pFindStr) Then
        BldDupCheck = True
    Else
        BldDupCheck = False
    End If
    
End Function

Private Sub chkBar_Click()
    txtBldNo = ""
    cboCompo.Clear
End Sub

Private Sub cmdApply_Click()
    Dim Row         As Long
    Dim strBldSrc   As String
    Dim strBldYY    As String
    Dim strBldno    As String
    Dim strBldNoAll As String
    Dim strCompocd  As String
    Dim strCompoNm  As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo, 1, 2)
        strBldYY = Mid(txtBldNo, 3, 2)
        strBldno = Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
    Else
        strBldSrc = medGetP(txtBldNo, 1, "-")
        strBldYY = medGetP(txtBldNo, 2, "-")
        strBldno = medGetP(txtBldNo, 3, "-")
    End If
    
    If strBldSrc = "" Or strBldYY = "" Or strBldno = "" Then
        MsgBox "혈액번호가 완전하지 않습니다.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If cboCompo.ListIndex < 0 Then Exit Sub
    
    strBldNoAll = strBldSrc & "-" & strBldYY & "-" & Format(strBldno, "00000#")
    
    strCompocd = Trim(Mid(cboCompo.Text, 1, 5))
    strCompoNm = Trim(Mid(cboCompo.Text, 6))
    
    '출고여부 확인---------------------------------------------------------
    If BldDeliveryChk(strBldSrc, strBldYY, Val(strBldno), strCompocd) Then
        MsgBox "이미 출고된 혈액입니다.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    '중복값 체크-----------------------------------------------------------
    If BldDupCheck(strBldNoAll & ";" & strCompocd) Then
        MsgBox "이미 입력된 Component입니다.", vbExclamation, "정보확인"
        Exit Sub
    End If
    
    '리스트에 보낸다.------------------------------------------------------
    Call SetTblDeliveryList(1, Format(GetSystemDate, "yyyy-MM-dd"))
    Call SetTblDeliveryList(2, strBldNoAll)
    Call SetTblDeliveryList(3, strCompoNm)    'Component Nm
    Call SetTblDeliveryList(5, strCompocd)    'Component Cd
    Call SetTblDeliveryList(4, GetABO(strBldSrc, strBldYY, Val(strBldno), strCompocd))   '혈액형
    Call SetTblDeliveryList(6, strBldNoAll & ";" & strCompocd)
    Call SetTblDeliveryList(7, txtRemark)
    Call SetTblDeliveryList(8, txtPtNm.Text)
    Call SetTblDeliveryList(9, txtABO.Text)
    
    cmdDelivery.Enabled = True
End Sub

Private Sub cmdClear_Click()
'    ClearAll
'    txtLocalCd.Enabled = True
'    cmdLocalCd.Enabled = True
    Call FormInitialize
End Sub
Private Function Delivery_Check() As Boolean
    '정보 누락 체크
    If Trim(txtLocalCd.Text) = "" Then
        MsgBox "병원코드를 입력하세요.", vbInformation, "정보확인"
        txtLocalCd.SetFocus
        Exit Function
    End If
    
    If Trim(txtBldNo.Text) = "" Then
        MsgBox "혈액 번호를 입력하세요.", vbInformation, "정보확인"
        txtBldNo.SetFocus
        Exit Function
    End If
    
    If cboCompo.Text = "" Then
        MsgBox "Component를 선택하세요.", vbInformation, "정보확인"
        cboCompo.SetFocus
        Exit Function
    End If
    Delivery_Check = True
End Function
Private Sub cmdDelivery_Click()
    
    Dim objBg           As clsBeginTrans
    Dim strBldSrc       As String
    Dim strBldYY        As String
    Dim lngBldNo        As String
    Dim strCompocd      As String
    Dim strDeliveryDt   As String
    Dim lngDeliverySeq  As Long
    Dim strDeliveryTm   As String
    Dim lngDeliveryID   As Long
    
    Dim strPtNm         As String
    Dim strABO          As String
    Dim strRemark       As String
    
    Dim i               As Long
    
    Dim SSQL            As String
    
    If Delivery_Check = False Then Exit Sub
    
    strDeliveryTm = Format(GetSystemDate, "hhMMss")
    lngDeliveryID = Val(ObjMyUser.EmpId)

On Error GoTo SAVE_ERROR
    
    Set objBg = New clsBeginTrans
        
    With tblDelivery
                
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: strDeliveryDt = Format(.value, "YYYYMMDD")
            .Col = 2: strBldSrc = Mid(.Text, 1, 2)
                      strBldYY = Mid(.Text, 4, 2)
                      lngBldNo = Format(Val(Mid(.Text, 7)), "00000#")
            .Col = 5: strCompocd = .Text
            .Col = 7: strRemark = .Text
            .Col = 8: strPtNm = .Text
            .Col = 9: strABO = .Text
            
            '출고 Seq를 가지고온다.............
            lngDeliverySeq = GetDeliverySeq(strBldSrc, strBldYY, lngBldNo, strCompocd)
            
            '출고내역 저장.........
            SSQL = objBg.SetLocalDelivery(strBldSrc, strBldYY, lngBldNo, strCompocd, strDeliveryDt, _
                                        lngDeliverySeq, strDeliveryTm, lngDeliveryID, 0, "", "", 0, _
                                        "", Trim(txtLocalCd.Text), strRemark, strPtNm, strABO)
            DBConn.Execute SSQL
            
            '입고내역 update.......
            SSQL = objBg.SetBldStorageUpdateByStsCd(strBldSrc, strBldYY, lngBldNo, strCompocd, _
                                                    BBSBloodStatus.stsDELIVERY)
            DBConn.Execute SSQL
            
        Next i
    End With
    
    DBConn.CommitTrans
    Call FormInitialize
    Set objBg = Nothing
    Exit Sub
    
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리 되지 않았습니다.", vbInformation, "정보확인"
    Set objBg = Nothing
    
End Sub

Private Function GetDeliverySeq(ByVal BldSrc As String, ByVal BldYY As String, _
                                ByVal BldNo As Long, ByVal CompoCd As String) As Long
'
'
    Dim objMyMaxSeq As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Set objMyMaxSeq = New clsBBSSQLStatement
    With objMyMaxSeq
'        .setDbConn DBConn
        RS.Open .GetBldDeliveryMaxSeq(BldSrc, BldYY, BldNo, CompoCd), DBConn
    End With
    
    If RS.EOF Then
        GetDeliverySeq = 1
    Else
        GetDeliverySeq = Val(RS.Fields("maxseq").value & "") + 1
    End If
    
    Set RS = Nothing
    Set objMyMaxSeq = Nothing
                               
End Function


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLocalCd_Click()

    Set objMyList = New clsPopUpList
    Set objMySql = New clsBBSSQLStatement
    
    With objMyList
        .Connection = DBConn
'        .BackColor = Me.BackColor
        .FormCaption = "LOCAL 병원조회": .ColumnHeaderText = "코드;코드명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Call .LoadPopUp(objMySql.GetLocalHp) ', 2350, 7650)
        txtLocalCd.Text = "": lblLocalNm.Caption = ""
        If .SelectedString <> "" Then
            txtLocalCd.Text = medGetP(.SelectedString, 1, ";")
            lblLocalNm.Caption = medGetP(.SelectedString, 2, ";")
            cmdQuery.Enabled = True
        End If
    End With
    
    Set objMySql = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim RS As New Recordset
    Dim i As Long
    Dim strBldno As String
    Dim strDeliveryFrom As String
    Dim strDeliveryTo As String
        
    tblResult.MaxRows = 0
    
    If Trim(txtLocalCd.Text) = "" Then Exit Sub
    
    
    strDeliveryFrom = Format(dtpFr.value, "yyyyMMdd")
    strDeliveryTo = Format(dtpTo.value, "yyyyMMdd")
    
    Set objMySql = New clsBBSSQLStatement
    
    With objMySql
'        .setDbConn DBConn
        RS.Open .GetBldDeliveryByLocalCd(Trim(txtLocalCd.Text), strDeliveryFrom, strDeliveryTo), DBConn
    End With
            
        If RS.EOF Then
            MsgBox "조회할 내용이 없습니다.", vbInformation, "정보확인"
            Set RS = Nothing
            Set objMySql = Nothing
            Exit Sub
        End If
    
    With tblResult
        .ReDraw = False
        .MaxRows = 1
        .Row = 0
        .Col = 0

        Do Until RS.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows - 1
            
            .Col = 1: .Text = Format(Trim(RS.Fields("deliverydt").value & ""), "####-##-##")
            .Col = 2: .Text = Trim(RS.Fields("bldsrc").value & "") & "-" & Trim(RS.Fields("bldyy").value & "") & "-" & Format(Trim(RS.Fields("bldno").value & ""), "00000#")
            .Col = 3: .Text = GetCompoCd(Trim(RS.Fields("compocd").value & ""))   'Component Nm
            .Col = 4: .Text = GetABO(Trim(RS.Fields("bldsrc").value & ""), Trim(RS.Fields("bldyy").value & ""), Format(Trim(RS.Fields("bldno").value & ""), "00000#"), Trim(RS.Fields("compocd").value & ""))
            .Col = 5: .Text = RS.Fields("rmk").value & ""
            .Col = 6: .Text = RS.Fields("ptnm").value & ""
            .Col = 7: .Text = RS.Fields("abo").value & ""
            
            RS.MoveNext
        Loop
        .MaxRows = .MaxRows - 1
    End With
    
    Set RS = Nothing
    Set objMySql = Nothing
End Sub

Private Function GetCompoCd(ByVal CompoCd As String) As String
    Dim objCompoCd As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Set objCompoCd = New clsBBSSQLStatement
    
    With objCompoCd
'        .setDbConn DBConn
        RS.Open .GetCompCdForCboBox(CompoCd), DBConn
    End With
    
    GetCompoCd = Trim(RS.Fields("field1").value & "")
    
    Set RS = Nothing
    Set objCompoCd = Nothing
End Function

Private Sub dtpFr_Change()
    tblResult.MaxRows = 0
End Sub

Private Sub dtpTo_Change()
    tblResult.MaxRows = 0
    
    If DateDiff("d", dtpFr.value, dtpTo.value) < 0 Then
        MsgBox "조회기간을 다시 설정하세요.", vbInformation, "정보확인"
        dtpTo.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim sysdate As Date

    txtLocalCd.Text = ""
    lblLocalNm.Caption = ""
    
    sysdate = GetSystemDate
    
    dtpFr.value = Format(sysdate, "YYYY-MM-") & "01"
    dtpTo.value = sysdate
    
    Call FormInitialize
End Sub

Private Sub FormInitialize()
    Dim sysdate As Date
    
    tblResult.MaxRows = 0
    
    txtBldNo.Text = ""
    cboCompo.Clear
    txtRemark = ""
    tblDelivery.MaxRows = 0
    cmdDelivery.Enabled = False
    txtPtNm.Text = ""
    txtABO.Text = ""
    
End Sub


Private Sub ClearAll()
    txtLocalCd = ""
    lblLocalNm.Caption = ""
    
    Clear
End Sub

Private Sub Clear()
    txtBldNo.Text = ""
    cboCompo.Clear
    tblResult.MaxRows = 0
    tblDelivery.MaxRows = 0
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblDelivery
                .Col = .ActiveCol
                .Row = .ActiveRow
                If Trim(.Text) = "" Then Exit Sub
                
                .Col = 1: .COL2 = .MaxCols: .Row = .ActiveRow: .Row2 = .ActiveRow
                .BlockMode = True
                .Action = ActionDeleteRow
                .BlockMode = False
                
                .MaxRows = .MaxRows - 1
            End With
    End Select
End Sub

'Private Sub mnuDelete_Click()
'    With tblDelivery
'        .Col = .ActiveCol
'        .Row = .ActiveRow
'        If Trim(.Text) = "" Then Exit Sub
'
'        .Col = 1: .COL2 = .MaxCols: .Row = .ActiveRow: .Row2 = .ActiveRow
'        .BlockMode = True
'        .Action = ActionDeleteRow
'        .BlockMode = False
'
'        .MaxRows = .MaxRows - 1
'    End With
'End Sub

Private Sub tblDelivery_RightClick(ByVal ClickType As Integer, _
                                   ByVal Col As Long, ByVal Row As Long, _
                                   ByVal MouseX As Long, ByVal MouseY As Long)

    With tblDelivery
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
    End With
    
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing

End Sub

Private Sub ShowCompoCdForCombo(ByVal BldNo As String)
'
    Dim objCompoCd As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim lngBldNo As Long
    Dim strTmp As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(BldNo, 1, 2)
        strBldYY = Mid(BldNo, 3, 2)
        lngBldNo = Mid(Mid(BldNo, 5), 1, Len(Mid(BldNo, 5)) - 2)
    Else
        strBldSrc = Mid(BldNo, 1, 2)
        strBldYY = Mid(BldNo, 4, 2)
        lngBldNo = Mid(BldNo, 7)
    End If
    Set objCompoCd = New clsBBSSQLStatement
    
    With objCompoCd
'        .setDbConn DBConn
    RS.Open .GetCompoCd(strBldSrc, strBldYY, lngBldNo), DBConn
    End With
    
    cboCompo.Clear
    With RS
        Do Until .EOF
            strTmp = .Fields("compocd").value & "" & Space(10)
            strTmp = Mid(strTmp, 1, 5) & .Fields("field1").value & ""
            
            cboCompo.AddItem strTmp
            .MoveNext
        Loop
    End With
        
    Set RS = Nothing
    Set objCompoCd = Nothing
End Sub

Private Sub txtABO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBldNo_Change()
    Dim lngLen As Long
    
    If chkBar.value = 1 Then Exit Sub
    
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strBldNoAll As String
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtBldNo.Text) = "" Then Exit Sub
        
        If chkBar.value = 1 Then
            strBldSrc = Mid(txtBldNo, 1, 2)
            strBldYY = Mid(txtBldNo, 3, 2)
            strBldno = Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
        Else
            strBldSrc = Mid(txtBldNo.Text, 1, 2)
            strBldYY = Mid(txtBldNo.Text, 4, 2)
            strBldno = Mid(txtBldNo.Text, 7)
        End If
        strBldNoAll = strBldSrc & "-" & strBldYY & "-" & Format(strBldno, "00000#")
        
        '혈액번호 존재 체크
        If BldNoExistChk(strBldSrc, strBldYY, Val(strBldno)) = False Then
            '혈액번호이 존재하지 않는 경우
            MsgBox "모두 출고됐거나 존재하지 않는 혈액번호입니다.", vbCritical, Me.Caption
            cboCompo.Clear
            With txtBldNo
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
        Call ShowCompoCdForCombo(txtBldNo.Text)
        cboCompo.SetFocus
    End If
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Function BldDeliveryChk(ByVal BldSrc As String, ByVal BldYY As String, _
                               ByVal BldNo As Long, ByVal CompoCd As String) As Boolean
'혈액 출고 체크
'출고된 경우 True

    Dim objDeliveryChk As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Set objDeliveryChk = New clsBBSSQLStatement
    With objDeliveryChk
'        .setDbConn DBConn
        RS.Open .GetBldStorageDeliveryChk(BldSrc, BldYY, BldNo, CompoCd), DBConn
    End With
    
    If RS.EOF Then
        BldDeliveryChk = False
    Else
        BldDeliveryChk = True
    End If
    
    Set RS = Nothing
    Set objDeliveryChk = Nothing

End Function

Private Function BldNoExistChk(ByVal BldSrc As String, ByVal BldYY As String, _
                               ByVal BldNo As Long) As Boolean
'혈액번호 존재 체크
'존재하면 True, 존재하지 않으면 False 반환

    Dim objExistChk As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Set objExistChk = New clsBBSSQLStatement
    With objExistChk
'        .setDbConn DBConn
        RS.Open .GetBldStorageBldNo(BldSrc, BldYY, BldNo), DBConn
    End With
    
    If RS.EOF Then
        BldNoExistChk = False
    Else
        BldNoExistChk = True
    End If
    
    Set RS = Nothing
    Set objExistChk = Nothing
End Function


Private Sub SetTblDeliveryList(ByVal pColPos As Long, ByVal pValue As String)
'pColPos = 1:출고일자 2: 혈액번호 3: Component 4: 혈액형 5:CompoCd 6: BldNo ; CompoCd

    With tblDelivery
        If pColPos = 1 Then .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = pColPos
        .Text = pValue
    End With
End Sub

Private Function GetABO(ByVal BldSrc As String, ByVal BldYY As String, _
                        ByVal BldNo As Long, ByVal CompoCd As String) As String
'
'
    Dim ObjABO As clsBBSSQLStatement
    Dim RS As New Recordset
    
    Set ObjABO = New clsBBSSQLStatement
        
    With ObjABO
'        .setDbConn DBConn
        RS.Open .GetABORh(BldSrc, BldYY, Format(BldNo, "00000#"), CompoCd), DBConn
        If Not RS.EOF Then
            GetABO = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
        End If
    End With
    
    Set RS = Nothing
    Set ObjABO = Nothing
End Function

Private Sub txtLocalCd_Change()
    lblLocalNm.Caption = ""
End Sub

Private Sub txtLocalCd_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyReturn Then
        If txtLocalCd.Text = "" Then Exit Sub
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub GetLocalCd()

    Dim RS As New Recordset
    
    Set objMySql = New clsBBSSQLStatement

    RS.Open objMySql.GetLocalHp(Trim(txtLocalCd.Text)), DBConn

    
    If RS.EOF Then
        With txtLocalCd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    
    lblLocalNm.Caption = RS.Fields("field1").value & ""
    
    cmdQuery.Enabled = True
    Set objMySql = Nothing
End Sub

Private Sub txtLocalCd_LostFocus()
    
    If txtLocalCd.Text = "" Then Exit Sub

    Call GetLocalCd
End Sub

VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm발주서업체 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "발주서작성(일반구매)"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   16350
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   16350
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16350
      _ExtentX        =   28840
      _ExtentY        =   16484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm발주서업체.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   915
         Left            =   12750
         TabIndex        =   5
         Top             =   30
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1614
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
            TabIndex        =   6
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "저장(&S)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   2340
            TabIndex        =   7
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "닫기(&X)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClear 
            Height          =   420
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "화면지움"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   915
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   1614
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtRemark 
            Height          =   300
            Left            =   3930
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   480
            Width           =   8475
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   61210625
            CurrentDate     =   41078
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   90
            TabIndex        =   2
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   90
            TabIndex        =   4
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매업체"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblSumAmt 
            Height          =   300
            Left            =   10350
            TabIndex        =   9
            Top             =   150
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   300
            Left            =   9120
            TabIndex        =   10
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "합계금액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdNo 
            Height          =   300
            Left            =   7740
            TabIndex        =   11
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   6510
            TabIndex        =   12
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   300
            Left            =   2700
            TabIndex        =   14
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PVCOMBOLibCtl.PVComboBox cboCust 
            Height          =   300
            Left            =   1320
            TabIndex        =   15
            Top             =   150
            Width           =   975
            _Version        =   524288
            _cx             =   1720
            _cy             =   529
            Appearance      =   1
            Enabled         =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            Locked          =   0   'False
            Style           =   2
            Sorted          =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowPictures    =   0   'False
            ColumnHeaders   =   -1  'True
            PrimaryColumn   =   0
            VisibleItems    =   10
            ColumnHeaderHeight=   20
            ListMember      =   ""
            ColumnHeaderForeColor=   0
            ColumnHeaderBackColor=   14215660
            SelectedForeColor=   16777215
            SelectedBackColor=   12937777
            AlternateBackColor=   16777215
            ItemLabelStyle  =   1
            ItemLabelType   =   0
            ItemLabelWidth  =   40
            ItemLabelForeColor=   0
            ItemLabelBackColor=   14215660
            ColumnHeaderStyle=   1
            VerticalGridLines=   0   'False
            HorizontalGridLines=   0   'False
            ColumnResize    =   0   'False
            ItemLabelResize =   0   'False
            AllowDBAutoConfig=   -1  'True
            GridLineColor   =   13421772
            List            =   ""
            NullString      =   "[NULL]"
            DropShadow      =   -1  'True
            Text            =   ""
            SortOnColumnHeaderClick=   0   'False
            DropEffect      =   1
            ColumnCount     =   2
            Column0.Heading =   "업체코드"
            Column0.Width   =   40
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   ""
            Column0.Format  =   ""
            Column0.Bound   =   0   'False
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "업체명"
            Column1.Width   =   80
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   ""
            Column1.Format  =   ""
            Column1.Bound   =   0   'False
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            SortKey1.Column =   -1
            SortKey1.Ascending=   -1  'True
            SortKey1.CaseInsensitive=   -1  'True
            SortKey2.Column =   -1
            SortKey2.Ascending=   -1  'True
            SortKey2.CaseInsensitive=   -1  'True
            SortKey3.Column =   -1
            SortKey3.Ascending=   -1  'True
            SortKey3.CaseInsensitive=   -1  'True
            BoundColumn     =   ""
            Border          =   -1  'True
            VertAlign       =   1
            Format          =   ""
         End
         Begin Threed.SSPanel lblCustNm 
            Height          =   300
            Left            =   2310
            TabIndex        =   16
            Top             =   150
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483634
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdYm 
            Height          =   300
            Left            =   7830
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdSeq 
            Height          =   300
            Left            =   8760
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   7785
         Left            =   30
         TabIndex        =   21
         Top             =   1050
         Width           =   16290
         _Version        =   524288
         _ExtentX        =   28734
         _ExtentY        =   13732
         _StockProps     =   64
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         MaxCols         =   12
         SpreadDesigner  =   "frm발주서업체.frx":0092
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   30
         TabIndex        =   22
         Top             =   8940
         Width           =   16290
         _ExtentX        =   28734
         _ExtentY        =   661
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "※ 구매업체 단위의 물품 발주서 등록만 가능합니다. 수정 및 삭제는 [발주서 일반]메뉴에서 작업하시기 바랍니다. !!!"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlOrdList 
      Height          =   5295
      Left            =   5700
      TabIndex        =   17
      Top             =   510
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
      _Version        =   262144
      BackColor       =   8438015
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPUSpreadADO.fpSpread spOrdList 
         Height          =   5040
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   6960
         _Version        =   524288
         _ExtentX        =   12277
         _ExtentY        =   8890
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483634
         MaxCols         =   6
         OperationMode   =   3
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frm발주서업체.frx":1042
         UserResize      =   0
      End
   End
End
Attribute VB_Name = "frm발주서업체"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCust_Click()
Dim sRow As Long, sDate As String

    If cboCust.ListIndex > 0 Then
        lblCustNm.Caption = cboCust.SubItem(cboCust.ListIndex, 1)
        
        sDate = Format(dtpDate.Value, "yyyy-MM-dd")
        
        gSql = "select stkcd, stknm, stkspec, buyunit, buyamt, buyday from mstSTK where custcd = " & Val(cboCust.Text) & " order by stknm"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    Call gsSpreadClear(spList, .RecordCount, True)
                    While (Not .EOF)
                        sRow = sRow + 1
                        
                        spList.SetText 1, sRow, ""
                        spList.SetText 2, sRow, "" & .Fields("stkcd").Value
                        spList.SetText 3, sRow, "" & .Fields("stknm").Value
                        spList.SetText 4, sRow, "" & .Fields("stkspec").Value
                        spList.SetText 5, sRow, "" & .Fields("buyunit").Value
                        spList.SetText 6, sRow, gfPresentStkRmd(.Fields("stkcd").Value, sDate, False)
                        spList.SetText 7, sRow, "" & .Fields("buyamt").Value
                        spList.SetText 10, sRow, DateAdd("d", Val("" & .Fields("buyday").Value), sDate)
                        
                        .MoveNext
                    Wend
                Else
                    Call gsSpreadClear(spList, 0, True)
                End If
                .Close
            End If
        End With
        
        cmdSave.Enabled = True
    Else
        lblCustNm.Caption = ""
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()
    
    pnlOrdList.Visible = False
    
    Call gsSpreadClear(spList, , True)

    dtpDate.Value = gfSystemDate
    Call gsSetCustComboPV(cboCust)
    
    lblOrdYm.Caption = ""
    lblOrdSeq.Caption = ""
    lblOrdNo.Caption = ""
    
    lblSumAmt.Caption = ""
    lblCustNm.Caption = ""
    txtRemark.Text = ""
    
    cmdSave.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()
Dim cOrdH As clsOrdHead, cOrdL As clsOrdList
Dim sRow As Long, sData As Variant, sCode As Long, sDate As String, sReturn As Boolean
Dim sReqDuty As String, sReqDt As String, sReqSeq As Integer, sAmt As Currency

    MousePointer = vbHourglass
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    If Len(lblOrdYm.Caption) = 0 Then
        lblOrdYm.Caption = Format(dtpDate.Value, "yyyyMM")
    End If
    
    Call cDb.csBegin
    
    Set cOrdH = New clsOrdHead
    With cOrdH
        .ordym = lblOrdYm.Caption
        .ordno = Val(lblOrdSeq.Caption)
        .custcd = Val(cboCust.Text)
        .orddt = sDate
'        .ordamt = Val(Str(lblSumAmt.Caption))
        .stat = gOrderStsWrt
        .ordtype = gOrderNormal
        .remark = Trim(txtRemark.Text)
        
        sReturn = .cfSave
        
        If sReturn Then
            lblOrdNo.Caption = .ordym & "-" & Trim(.ordno)
            lblOrdYm.Caption = .ordym
            lblOrdSeq.Caption = .ordno
        Else
            lblOrdYm.Caption = ""
            lblOrdSeq.Caption = ""
        End If
    End With
    
    If sReturn Then
        Set cOrdL = New clsOrdList
        cOrdL.ordym = lblOrdYm.Caption
        cOrdL.ordno = Val(lblOrdSeq.Caption)
        
        With spList
            For sRow = 1 To .MaxRows
                .GetText 2, sRow, sData:    sCode = Val(sData)
                .GetText 1, sRow, sData
                If Val(sData) > 0 And sCode > 0 Then
                    cOrdL.stkcd = sCode
                    .GetText .MaxCols, sRow, sData:     cOrdL.ordseq = Val(sData)
                    .GetText 7, sRow, sData:            cOrdL.amt = Val(sData)
                    .GetText 8, sRow, sData:            cOrdL.qty = Val(sData)
                    .GetText 9, sRow, sData:            cOrdL.sumamt = Val(sData)
                    .GetText 10, sRow, sData:           cOrdL.lastdt = Trim(sData)
                    .GetText 11, sRow, sData:           cOrdL.remark = Trim(sData)
                    
                    sReturn = cOrdL.cfSave
                    If sReturn = False Then
                        Exit For
                    Else
                        .SetText 1, sRow, ""
                        .SetText .MaxCols, sRow, cOrdL.ordseq
                    End If
                End If
            Next sRow
        End With
    End If
    
    If sReturn Then
        sReturn = cOrdH.cfAmtSumUpdate(lblOrdYm.Caption, Val(lblOrdSeq.Caption), sAmt)
    End If
    
    If sReturn Then
        Call cDb.csCommit
        lblSumAmt.Caption = Format(sAmt, "#,##0")
    Else
        Call cDb.csRollback
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    Me.Show
    
    Call cmdClear_Click
    
End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String, sAmt As Currency, sQty As Single

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
        
        If Col = 7 Or Col = 8 Then
            With spList
                .GetText 7, Row, sData:     sAmt = Val(sData)
                .GetText 8, Row, sData:     sQty = Val(sData)
                .SetText 9, Row, sAmt * sQty
            End With
        End If
    End If

End Sub

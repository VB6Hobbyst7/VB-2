VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm발주서요청 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "발주서작성(요청구매)"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18915
      _ExtentX        =   33364
      _ExtentY        =   17224
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm발주서요청.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   945
         Left            =   15390
         TabIndex        =   1
         Top             =   30
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1667
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
            TabIndex        =   2
            Top             =   90
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
            TabIndex        =   3
            Top             =   90
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
            TabIndex        =   4
            Top             =   90
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
         Height          =   945
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1667
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtRemark 
            Height          =   300
            Left            =   1320
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   480
            Width           =   13785
         End
         Begin VB.ComboBox cboCust 
            Height          =   300
            Left            =   3930
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   150
            Width           =   4035
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   7
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   21430273
            CurrentDate     =   41078
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   90
            TabIndex        =   8
            Top             =   150
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
            Left            =   2700
            TabIndex        =   9
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
         Begin Threed.SSPanel SSPanel1 
            Height          =   300
            Left            =   90
            TabIndex        =   10
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
         Begin Threed.SSPanel lblOrdNo 
            Height          =   300
            Left            =   9240
            TabIndex        =   12
            Top             =   150
            Width           =   2295
            _ExtentX        =   4048
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
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdYm 
            Height          =   300
            Left            =   9240
            TabIndex        =   13
            Top             =   -60
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
            Left            =   10170
            TabIndex        =   14
            Top             =   -60
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   300
            Left            =   8010
            TabIndex        =   15
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
         Begin Threed.SSPanel lblSumAmt 
            Height          =   300
            Left            =   12810
            TabIndex        =   16
            Top             =   150
            Width           =   2295
            _ExtentX        =   4048
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
            Left            =   11580
            TabIndex        =   17
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
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   8175
         Left            =   30
         TabIndex        =   18
         Top             =   1080
         Width           =   18855
         _Version        =   524288
         _ExtentX        =   33258
         _ExtentY        =   14420
         _StockProps     =   64
         EditEnterAction =   2
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
         MaxCols         =   18
         SpreadDesigner  =   "frm발주서요청.frx":0092
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Left            =   30
         TabIndex        =   19
         Top             =   9360
         Width           =   18855
         _ExtentX        =   33258
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
         Caption         =   "※ 구매요청서 기준의 발주처리만 가능합니다. 수정 및 삭제는 [발주서 일반]메뉴에서 작업하시기 바랍니다. !!!"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frm발주서요청"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCust_Click()
Dim sRow As Long, sDate As String
    
    MousePointer = vbHourglass
    If cboCust.ListIndex > 0 Then
        sDate = Format(dtpDate.Value, "yyyy-MM-dd")
        gSql = "select a.*, b.stknm, b.stkspec, b.buyunit, b.buyamt, c.dutynm, d.usernm from reqL a with (nolock), mstSTK b, mstDUTY c, mstUSER d" & _
             " where a.reqdt <= '" & sDate & "' and a.stat = " & gReqStatWrt & " and a.stkcd = b.stkcd and b.custcd = " & cboCust.ItemData(cboCust.ListIndex) & _
             " and a.dutycd = c.dutycd and a.usercd = d.usercd order by a.stkcd, a.reqdt"
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
                        spList.SetText 6, sRow, gfPresentStkRmd(Val("" & .Fields("stkcd").Value), sDate, False)
                        spList.SetText 7, sRow, Val("" & .Fields("buyamt").Value)
                        spList.SetText 8, sRow, Val("" & .Fields("qty").Value)
                        spList.SetText 9, sRow, Val("" & .Fields("buyamt").Value) * Val("" & .Fields("qty").Value)
                        spList.SetText 10, sRow, "" & .Fields("lastdt").Value
                        spList.SetText 11, sRow, ""
                        spList.SetText 12, sRow, "" & .Fields("dutynm").Value
                        spList.SetText 13, sRow, "" & .Fields("reqdt").Value
                        spList.SetText 14, sRow, "" & .Fields("usernm").Value
    
                        spList.SetText 15, sRow, "" & .Fields("dutycd").Value
                        spList.SetText 16, sRow, "" & .Fields("reqdt").Value
                        spList.SetText 17, sRow, "" & .Fields("reqseq").Value
    
                        spList.SetText 18, sRow, ""
                        
                        .MoveNext
                    Wend
                Else
                    Call gsSpreadClear(spList, 0, True)
                    MsgBox "해당업체 물품요청서가 없습니다.!", vbCritical
                End If
                .Close
            End If
        End With
    Else
        Call gsSpreadClear(spList, 0, True)
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdClear_Click()

    lblOrdNo.Caption = ""
    lblOrdYm.Caption = ""
    lblOrdSeq.Caption = ""
    
    lblSumAmt.Caption = ""
    
    dtpDate.Value = gfSystemDate
    txtRemark.Text = ""
    
    cboCust.Clear
    cboCust.AddItem ""
    gSql = "select b.custcd, max(c.custnm) as custnm from reqL a with (nolock), mstSTK b, mstCUST c" & _
           " where a.stat = " & gReqStatWrt & " and a.stkcd = b.stkcd and b.custcd = c.custcd " & _
           " group by b.custcd order by custnm"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                cboCust.AddItem .Fields("custnm").Value
                cboCust.ItemData(cboCust.NewIndex) = .Fields("custcd").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With
    
    Call gsSpreadClear(spList, , True)

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()
Dim cOrdH As clsOrdHead, cOrdL As clsOrdList
Dim sRow As Long, sData As Variant, sCode As Long, sDate As String, sReturn As Boolean
Dim sReqDuty As String, sReqDt As String, sReqSeq As Integer

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
'        .custcd = cboCust.ItemData(cboCust.ListIndex)
        .orddt = sDate
        .ordamt = Val(Str(lblSumAmt.Caption))
        .stat = gOrderStsWrt
        .ordtype = gOrderReq
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
                        
                        ' 요청서 발주번호 등록
                        .GetText 15, sRow, sData:       sReqDuty = Trim(sData)
                        .GetText 16, sRow, sData:       sReqDt = Trim(sData)
                        .GetText 17, sRow, sData:       sReqSeq = Val(sData)
                        
                        gSql = "update reqL set ordym = '" & cOrdL.ordym & "', ordno = " & cOrdL.ordno & ", ordseq = " & cOrdL.ordseq & ", stat = " & gReqStatOrder & _
                               " where dutycd = '" & sReqDuty & "' and reqdt = '" & sReqDt & "' and reqseq = " & sReqSeq
                        sReturn = cDb.cfExecute(gSql)
                        If sReturn = False Then Exit For
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
        lblSumAmt.Caption = Format(cOrdH.ordamt, "#,##0")
    Else
        Call cDb.csRollback
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()

    Me.Width = 19000
    Me.Height = 10120

    Me.Show

    Call cmdClear_Click

End Sub

Private Sub spList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim sRow As Long, sData As Variant, sAmt As Currency
    
    If Col = 1 Then
        With spList
            For sRow = 1 To .MaxRows
                .GetText 1, sRow, sData
                If Val(sData) > 0 Then
                    .GetText 9, sRow, sData
                    sAmt = sAmt + Val(sData)
                End If
            Next sRow
            lblSumAmt.Caption = Format(sAmt, "#,##0")
        End With
    End If
    
End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sAmt As Currency, sQty As Single

    If Row > 0 And Col >= 1 Then
        With spList
            If Col = 7 Or Col = 8 Then
                .GetText 7, Row, sData:   sAmt = Val(sData)
                .GetText 8, Row, sData:   sQty = Val(sData)
                .SetText 9, Row, sAmt * sQty
            End If
        End With
    End If

End Sub

VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm구매서일반 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "구매서작성(일반구매)"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18915
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
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18915
      _ExtentX        =   33364
      _ExtentY        =   16484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm구매서일반.frx":0000
      Begin SSActiveTreeView.SSTree trvStkList 
         Height          =   8565
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   15108
         _Version        =   65538
         Appearance      =   0
         Indentation     =   569.764
         PictureBackgroundUseMask=   0   'False
         HasFont         =   0   'False
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   615
         Left            =   14235
         TabIndex        =   6
         Top             =   30
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
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
            Caption         =   "저장(&S)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   420
            Left            =   2340
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
            Caption         =   "삭제(&D)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   3450
            TabIndex        =   9
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
            TabIndex        =   11
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
         Height          =   615
         Left            =   3015
         TabIndex        =   2
         Top             =   30
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60948481
            CurrentDate     =   41078
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   90
            TabIndex        =   3
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   2700
            TabIndex        =   5
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
            Left            =   9300
            TabIndex        =   13
            Top             =   150
            Width           =   1665
            _ExtentX        =   2937
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
            Left            =   8070
            TabIndex        =   14
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
         Begin PVCOMBOLibCtl.PVComboBox cboCust 
            Height          =   300
            Left            =   3930
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
            Left            =   4920
            TabIndex        =   16
            Top             =   150
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1085
         _Version        =   262144
         Font3D          =   5
         ForeColor       =   65535
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "▒ 물품현황 ▒"
         BevelOuter      =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   8565
         Left            =   3015
         TabIndex        =   12
         Top             =   750
         Width           =   15870
         _Version        =   524288
         _ExtentX        =   27993
         _ExtentY        =   15108
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
         MaxCols         =   15
         SpreadDesigner  =   "frm구매서일반.frx":00B2
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm구매서일반"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long, sDate As String, sAmt As Currency

    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    gSql = "select a.*, b.stknm, b.stkspec, b.buyunit from buyL a with (nolock), mstSTK b" & _
           " where a.buydt = '" & sDate & "' and a.custcd = " & Val(cboCust.Text) & _
           "   and a.stkcd = b.stkcd order by a.buyseq"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("stkcd").Value
                    spList.SetText 3, sRow, "" & .Fields("stknm").Value
                    spList.SetText 4, sRow, "" & .Fields("stkspec").Value
                    spList.SetText 5, sRow, "" & .Fields("buyunit").Value
                    spList.SetText 6, sRow, Val("" & .Fields("qtyrate").Value)
                    spList.SetText 7, sRow, Val("" & .Fields("amt").Value)
                    spList.SetText 8, sRow, Val("" & .Fields("buyqty").Value)
                    spList.SetText 9, sRow, Val("" & .Fields("sumamt").Value)
                    spList.SetText 10, sRow, "" & .Fields("maxdt").Value
                    spList.SetText 11, sRow, "" & .Fields("makeno").Value
                    spList.SetText 12, sRow, "" & .Fields("ordym").Value
                    spList.SetText 13, sRow, Val("" & .Fields("ordno").Value)
                    spList.SetText 14, sRow, Val("" & .Fields("ordseq").Value)
                    spList.SetText 15, sRow, Val("" & .Fields("buyseq").Value)
                    
                    sAmt = sAmt + Val("" & .Fields("sumamt").Value)
                    
                    .MoveNext
                Wend
                cmdDelete.Enabled = True
            Else
                Call gsSpreadClear(spList, , True)
                cmdDelete.Enabled = False
            End If
            .Close
        End If
    End With
    
    lblSumAmt.Caption = Format(sAmt, "#,##0")

End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim cBuy As clsBuyList
Dim sRow As Long, sData As Variant, sDate As String, sReturn As Boolean, sCode As Long, sSeq As Integer
    
    sReturn = True
    
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    
    Set cBuy = New clsBuyList
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 2, sRow, sData:            sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                Call cDb.csBegin
                If brJob Then
                    cBuy.buydt = sDate
                    cBuy.buyseq = sSeq
                    cBuy.stkcd = sCode
                    cBuy.custcd = Val(cboCust.Text)
                    cBuy.usercd = gUserId
                    cBuy.buytype = gBuyIoNormal
                    .GetText 6, sRow, sData:            cBuy.qtyrate = Val(sData)
                    .GetText 7, sRow, sData:            cBuy.amt = Val(sData)
                    .GetText 8, sRow, sData:            cBuy.buyqty = Trim(sData)
                    .GetText 9, sRow, sData:            cBuy.sumamt = Trim(sData)
                    .GetText 10, sRow, sData:           cBuy.maxdt = Trim(sData)
                    .GetText 11, sRow, sData:           cBuy.makeno = Trim(sData)
                    .GetText 12, sRow, sData:           cBuy.ordym = Trim(sData)
                    .GetText 13, sRow, sData:           cBuy.ordno = Val(sData)
                    .GetText 14, sRow, sData:           cBuy.ordseq = Val(sData)
                    sReturn = cBuy.cfSave
                    If sReturn Then
                        .SetText .MaxCols, sRow, cBuy.buyseq
                    End If
                Else
                    If sSeq > 0 Then
                        sReturn = cBuy.cfDelete(sDate, sSeq)
                    End If
                End If
                
                If sReturn = False Then
                    Call cDb.csRollback
                    Exit For
                Else
                    Call cDb.csCommit
                    .SetText 1, sRow, ""
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psListRefresh
    End If

End Sub

Private Sub cboCust_Click()

    If cboCust.ListIndex > 0 Then
        lblCustNm.Caption = cboCust.SubItem(cboCust.ListIndex, 1)
        
        Call psListRefresh
        
        cmdSave.Enabled = True
    Else
        lblCustNm.Caption = ""
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()
    
    Call gsSpreadClear(spList, , True)
    Call gsSetStkTree(trvStkList)

    dtpDate.Value = gfSystemDate
    Call gsSetCustComboPV(cboCust)
    
    cboCust.Text = ""
    lblSumAmt.Caption = ""
    lblCustNm.Caption = ""
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 구매자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psDataProcess(False)
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Call psDataProcess(True)
    MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
Dim sCol As Integer

    Me.Width = 19000
    Me.Height = 10120
    
    Me.KeyPreview = True
    Me.Show
    
    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String, sAmt As Currency, sQty As Single

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
        
        Select Case Col
            Case 2
                    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
                    spList.GetText 2, Row, sData
                    gSql = "select stkcd, stknm, stkspec, buyunit, buyioqty, buyamt from mstSTK where stkcd = " & Val(sData)
                    With cDb.cfRecordSet(gSql)
                        If .State = adStateOpen Then
                            If Not .EOF Then
                                spList.SetText 3, Row, "" & .Fields("stknm").Value
                                spList.SetText 4, Row, "" & .Fields("stkspec").Value
                                spList.SetText 5, Row, "" & .Fields("buyunit").Value
                                spList.SetText 6, Row, Val("" & .Fields("buyioqty").Value)
                                spList.SetText 7, Row, Val("" & .Fields("buyamt").Value)
                                
                                spList.Row = Row
                                spList.Col = 7
                                spList.Action = ActionActiveCell
                            Else
                                MsgBox "등록되지 않은 물품입니다.!", vbCritical
                                
                                spList.SetText 2, Row, ""
                                spList.SetText 3, Row, ""
                                spList.SetText 4, Row, ""
                                spList.SetText 5, Row, ""
                                spList.SetText 6, Row, ""
                                spList.SetText 7, Row, ""
                            
                                spList.Row = Row
                                spList.Col = 2
                                spList.Action = ActionActiveCell
                                spList.SetFocus
                            End If
                            .Close
                        End If
                    End With
            Case 7, 8
                    With spList
                        .GetText 7, Row, sData:     sAmt = Val(sData)
                        .GetText 8, Row, sData:     sQty = Val(sData)
                        .SetText 9, Row, sAmt * sQty
                    End With
        End Select
    End If

End Sub

Private Sub trvStkList_Collapse(Node As SSActiveTreeView.SSNode)

    If Node.Level = 1 Then Node.Image = "close"

End Sub

Private Sub trvStkList_DblClick()

    If trvStkList.SelectedNodes.Item(1).Level > 1 Then
        Call psStkAppendCheck
    End If

End Sub

Private Sub trvStkList_Expand(Node As SSActiveTreeView.SSNode)

    If Node.Level = 1 Then Node.Image = "open"

End Sub

Private Sub psStkAppendCheck()
Dim sRow As Long, sData As Variant

    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData
            If Len(sData) = 0 Then
                .SetText 2, sRow, trvStkList.SelectedItem.Key
                Call spList_Change(2, sRow)
                
                spList.Row = sRow
                spList.Col = 8
                spList.Action = ActionActiveCell
                Exit For
'            ElseIf Trim(sData) = trvStkList.SelectedItem.Key Then
'                MsgBox "등록된 품번입니다.!", vbCritical
'                Exit For
            End If
        Next sRow
    End With

End Sub

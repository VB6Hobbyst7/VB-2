VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm출고서일반 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "출고서작성(일반출고)"
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
      PaneTree        =   "frm출고서일반.frx":0000
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
         Begin VB.ComboBox cboDuty 
            Height          =   300
            Left            =   3930
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   150
            Width           =   3045
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   21364737
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
            Caption         =   "출고일자"
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
            Caption         =   "출고부서"
            BevelOuter      =   1
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
         MaxCols         =   11
         SpreadDesigner  =   "frm출고서일반.frx":00B2
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm출고서일반"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long, sDate As String

    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    gSql = "select a.*, b.stknm, b.stkspec, b.iounit, c.usernm from outL a with (nolock), mstSTK b, mstUSER c" & _
           " where a.outdt = '" & sDate & "' and a.outfg = " & gOutNormal & " and a.dutycd = '" & cboDuty.ItemData(cboDuty.ListIndex) & "'" & _
           "   and a.stkcd = b.stkcd and a.usercd = c.usercd order by a.outseq"
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
                    spList.SetText 5, sRow, "" & .Fields("iounit").Value
                    spList.SetText 6, sRow, ""
'                    spList.SetText 6, sRow, gfPresentStkRmd(.Fields("stkcd").Value, sDate, False)
                    spList.SetText 7, sRow, Val("" & .Fields("qty").Value)
                    spList.SetText 8, sRow, "" & .Fields("reason").Value
                    spList.SetText 9, sRow, "" & .Fields("usernm").Value
                    spList.SetText 10, sRow, "" & .Fields("moddt").Value
                    spList.SetText 11, sRow, Val("" & .Fields("outseq").Value)
                    
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

End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim cOut As clsOutList
Dim sRow As Long, sData As Variant, sDate As String, sReturn As Boolean, sCode As Long, sSeq As Integer
    
    sReturn = True
    
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    
    Set cOut = New clsOutList
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 2, sRow, sData:            sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                Call cDb.csBegin
                If brJob Then
                    cOut.outdt = sDate
                    cOut.outfg = gOutNormal
                    cOut.outseq = sSeq
                    cOut.stkcd = sCode
                    cOut.dutycd = cboDuty.ItemData(cboDuty.ListIndex)
                    cOut.usercd = gUserId
                    .GetText 7, sRow, sData:            cOut.qty = Val(sData)
                    .GetText 8, sRow, sData:            cOut.reason = Trim(sData)
                    sReturn = cOut.cfSave
                    If sReturn Then
                        .SetText .MaxCols, sRow, cOut.outseq
                    End If
                Else
                    If sSeq > 0 Then
                        sReturn = cOut.cfDelete(sDate, gOutNormal, sSeq)
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

Private Sub cboDuty_Click()

    If cboDuty.ListIndex >= 0 Then
        Call psListRefresh
        
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()
    
    Call gsSpreadClear(spList, , True)
    Call gsSetStkTree(trvStkList)

    dtpDate.Value = gfSystemDate
    Call gsSetDutyCombo(cboDuty)
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 출고자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
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
        
        If Col = 2 Then
            sDate = Format(dtpDate.Value, "yyyy-MM-dd")
            
            spList.GetText 2, Row, sData
            gSql = "select stkcd, stknm, stkspec, iounit from mstSTK where stkcd = " & Val(sData)
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        spList.SetText 3, Row, "" & .Fields("stknm").Value
                        spList.SetText 4, Row, "" & .Fields("stkspec").Value
                        spList.SetText 5, Row, "" & .Fields("iounit").Value
                        spList.SetText 6, Row, gfPresentStkRmd(.Fields("stkcd").Value, sDate)
                        
                        spList.Row = Row
                        spList.Col = 6
                        spList.Action = ActionActiveCell
                    Else
                        MsgBox "등록되지 않은 물품입니다.!", vbCritical
                        
                        spList.SetText 2, Row, ""
                        spList.SetText 3, Row, ""
                        spList.SetText 4, Row, ""
                        spList.SetText 5, Row, ""
                        spList.SetText 6, Row, ""
                        spList.SetText 7, Row, ""
                        spList.SetText 8, Row, ""
                    
                        spList.Row = Row
                        spList.Col = 2
                        spList.Action = ActionActiveCell
                        spList.SetFocus
                    End If
                    .Close
                End If
            End With
        End If
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
                spList.Col = 7
                spList.Action = ActionActiveCell
                Exit For
'            ElseIf Trim(sData) = trvStkList.SelectedItem.Key Then
'                MsgBox "등록된 품번입니다.!", vbCritical
'                Exit For
            End If
        Next sRow
    End With

End Sub

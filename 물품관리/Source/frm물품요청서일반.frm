VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm물품요청서일반 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "물품요청서(일반)"
   ClientHeight    =   9765
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
      PaneTree        =   "frm물품요청서일반.frx":0000
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   9360
         Width           =   14445
         _ExtentX        =   25479
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
         Caption         =   "※ 발주 및 입고상태의 자료는 수정/삭제 하실수 없습니다 !!!"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin SSActiveTreeView.SSTree trvStkList 
         Height          =   8955
         Left            =   30
         TabIndex        =   13
         Top             =   780
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   15796
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
         Height          =   645
         Left            =   13110
         TabIndex        =   7
         Top             =   30
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1138
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
            TabIndex        =   8
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
         Begin Threed.SSCommand cmdDelete 
            Height          =   420
            Left            =   2340
            TabIndex        =   9
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
            Caption         =   "삭제(&D)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdDelay 
            Height          =   420
            Left            =   3450
            TabIndex        =   10
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
            Caption         =   "보류처리"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   4560
            TabIndex        =   11
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
            TabIndex        =   14
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
         Height          =   645
         Left            =   4440
         TabIndex        =   2
         Top             =   30
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   1138
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkUser 
            Caption         =   "사용자자료"
            Height          =   420
            Left            =   7530
            TabIndex        =   12
            Top             =   90
            Width           =   975
         End
         Begin VB.ComboBox cboDuty 
            Height          =   300
            Left            =   3930
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
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
            Format          =   60882945
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
            Caption         =   "요청일자"
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
            Caption         =   "요청부서"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   1138
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
         Height          =   8475
         Left            =   4440
         TabIndex        =   15
         Top             =   780
         Width           =   14445
         _Version        =   524288
         _ExtentX        =   25479
         _ExtentY        =   14949
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
         MaxCols         =   13
         MaxRows         =   1000000
         SpreadDesigner  =   "frm물품요청서일반.frx":00D2
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm물품요청서일반"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long, sDate As String, sDutyCd As String
    
    If cboDuty.ListIndex >= 0 Then
        sDate = Format(dtpDate.Value, "yyyy-MM-dd")
        sDutyCd = cboDuty.ItemData(cboDuty.ListIndex)
        
        gSql = "select a.*, b.stknm, b.stkspec, b.buyunit, c.usernm from reqL a, mstSTK b, mstUSER c" & _
               " where a.dutycd = '" & sDutyCd & "' and a.reqdt = '" & sDate & "' and a.stkcd = b.stkcd and a.usercd = c.usercd"
        If chkUser.Value > 0 Then
            gSql = gSql & " and a.usercd = '" & gUserId & "'"
        End If
        gSql = gSql & " order by a.reqseq"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    Call gsSpreadClear(spList, .RecordCount + 1000, True)
                    sRow = 0
                    While (Not .EOF)
                        sRow = sRow + 1
                        
                        spList.SetText 1, sRow, ""
                        spList.SetText 2, sRow, "" & .Fields("stkcd").Value
                        spList.SetText 3, sRow, "" & .Fields("stknm").Value
                        spList.SetText 4, sRow, "" & .Fields("stkspec").Value
                        spList.SetText 5, sRow, "" & .Fields("buyunit").Value
        '                spList.SetText 6, sRow, gfPresentStkRmd(.Fields("stkcd").Value, sDate, False)
                        spList.SetText 6, sRow, ""
                        spList.SetText 7, sRow, Val("" & .Fields("qty").Value)
                        spList.SetText 8, sRow, "" & .Fields("lastdt").Value
                        spList.SetText 9, sRow, gReqStatus(.Fields("stat").Value)
                        spList.Row = sRow
                        spList.Col = 9
                        Select Case Val("" & .Fields("stat").Value)
                            Case gReqStatWrt
                                    spList.ForeColor = vbBlack
                            Case gReqStatHold
                                    spList.ForeColor = vbRed
                            Case gReqStatOrder, gReqStatBuy
                                    spList.ForeColor = vbBlue
                        End Select
                        spList.SetText 10, sRow, "" & .Fields("usernm").Value
                        spList.SetText 11, sRow, "" & .Fields("remark").Value
                        spList.SetText 12, sRow, Val("" & .Fields("stat").Value)
                        spList.SetText 13, sRow, Val("" & .Fields("reqseq").Value)
                        
                        .MoveNext
                    Wend
                    cmdDelete.Enabled = True
                    cmdDelay.Enabled = True
                Else
                    cmdDelete.Enabled = False
                    cmdDelay.Enabled = False
                    Call gsSpreadClear(spList, , True)
                End If
                .Close
            End If
        End With
        cmdSave.Enabled = True
    End If
    
End Sub

Private Sub psDataProcess(ByVal brJob As Integer)
Dim cReq As clsReqList
Dim sRow As Long, sData As Variant, sReturn As Boolean, sCode As Long, sDate As String, sSeq As Integer, sStat As Integer

    Set cReq = New clsReqList
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 12, sRow, sData:           sStat = Val(sData)
            .GetText 2, sRow, sData:            sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 And sStat < gReqStatOrder Then
                Select Case brJob
                    Case 0
                            cReq.dutycd = cboDuty.ItemData(cboDuty.ListIndex)
                            cReq.reqdt = sDate
                            cReq.reqseq = sSeq
                            cReq.stkcd = sCode
                            .GetText 7, sRow, sData:            cReq.qty = Val(sData)
                            .GetText 8, sRow, sData:            cReq.lastdt = Trim(sData)
                            .GetText 11, sRow, sData:           cReq.remark = Trim(sData)
                            cReq.usercd = gUserId
                            
                            sReturn = cReq.cfSave
                            If sReturn Then
                                .SetText .MaxCols, sRow, cReq.reqseq
                            End If
                    Case 1
                            If sSeq > 0 Then
                                sReturn = cReq.cfDelete(cboDuty.ItemData(cboDuty.ListIndex), sDate, sSeq)
                            Else
                                sReturn = True
                            End If
                    Case 2
                            If sSeq > 0 Then
                                sReturn = cReq.cfHoldSet(cboDuty.ItemData(cboDuty.ListIndex), sDate, sSeq)
                            Else
                                sReturn = True
                            End If
                End Select
                
                If sReturn = False Then
                    Exit For
                Else
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
        spList.SetFocus
    End If
    
End Sub

Private Sub chkUser_Click()

    Call psListRefresh
    
End Sub

Private Sub cmdClear_Click()
    
    Call gsSpreadClear(spList, , True)
    Call gsSetStkTree(trvStkList)

    dtpDate.Value = gfSystemDate
    Call gsSetDutyCombo(cboDuty)
    
    With spList
        .Row = -1
        .Col = 7
        .ForeColor = vbRed
    End With
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdDelay.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelay_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 요청자료를 보류처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psDataProcess(2)
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 요청자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psDataProcess(1)
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Call psDataProcess(0)
    MousePointer = vbDefault
    
End Sub

Private Sub dtpDate_Change()

    Call psListRefresh
    
End Sub

Private Sub Form_Load()

    Me.Width = 19000
    Me.Height = 10120
    
    Me.Show
    
    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
        
        If Col = 2 Then
            sDate = Format(dtpDate.Value, "yyyy-MM-dd")
            spList.GetText 2, Row, sData
            gSql = "select stkcd, stknm, stkspec, buyunit, buyday from mstSTK where stkcd = " & Val(sData)
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        spList.SetText 3, Row, "" & .Fields("stknm").Value
                        spList.SetText 4, Row, "" & .Fields("stkspec").Value
                        spList.SetText 5, Row, "" & .Fields("buyunit").Value
                        spList.SetText 6, Row, gfPresentStkRmd(Val(sData), sDate, False)
                        spList.SetText 8, Row, DateAdd("d", Val("" & .Fields("buyday").Value), sDate)
                        
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
                        spList.SetText 8, Row, ""
                    
                        spList.Row = Row
                        spList.Col = 1
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


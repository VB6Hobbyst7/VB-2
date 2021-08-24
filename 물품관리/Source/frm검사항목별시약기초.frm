VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm검사항목별시약기초 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장비별 운영시약 기초자료"
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
   Icon            =   "frm검사항목별시약기초.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
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
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm검사항목별시약기초.frx":000C
      Begin SSActiveTreeView.SSTree trvStkList 
         Height          =   9000
         Left            =   14895
         TabIndex        =   2
         Top             =   735
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   15875
         _Version        =   65538
         Appearance      =   0
         LabelEdit       =   1
         Indentation     =   569.764
         PictureBackgroundUseMask=   0   'False
         HasFont         =   0   'False
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   600
         Left            =   14895
         TabIndex        =   1
         Top             =   30
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   1058
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
         Caption         =   " ▒ 시약리스트"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   555
         Left            =   4455
         TabIndex        =   3
         Top             =   750
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   979
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtTestCd 
            Height          =   300
            Left            =   1350
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   120
            Width           =   975
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   300
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "검사항목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblTestNm 
            Height          =   300
            Left            =   2340
            TabIndex        =   5
            Top             =   120
            Width           =   3465
            _ExtentX        =   6112
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   600
         Left            =   4455
         TabIndex        =   6
         Top             =   9135
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1058
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   6780
            TabIndex        =   7
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
            Left            =   7890
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
            Caption         =   "삭제(&D)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClear 
            Height          =   420
            Left            =   5670
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
            Caption         =   "화면지움"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   9000
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
            Caption         =   "닫기(&X)"
            ButtonStyle     =   2
         End
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   7620
         Left            =   4455
         TabIndex        =   11
         Top             =   1410
         Width           =   10335
         _Version        =   524288
         _ExtentX        =   18230
         _ExtentY        =   13441
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
         GrayAreaBackColor=   -2147483633
         MaxCols         =   6
         SpreadDesigner  =   "frm검사항목별시약기초.frx":011E
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   4455
         TabIndex        =   12
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
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
         Caption         =   " ▒ 검사항목별 소요량"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   615
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   4320
         _ExtentX        =   7620
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
         Caption         =   " ▒ 검사항목"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin SSActiveTreeView.SSTree trvTestList 
         Height          =   8985
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   15849
         _Version        =   65538
         Appearance      =   0
         LabelEdit       =   1
         Indentation     =   569.764
         PictureBackgroundUseMask=   0   'False
         HasFont         =   0   'False
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
      End
   End
End
Attribute VB_Name = "frm검사항목별시약기초"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTstk As clsTestStk

Private Sub psListRefresh()
Dim sRow As Long

    gSql = "select a.*, b.stknm, b.stkspec, b.iounit from testSTK a with (nolock), mstSTK b" & _
           " where a.testcd = '" & Trim(txtTestCd.Text) & "' and a.stkcd = b.stkcd order by a.stkcd"
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
                    spList.SetText 6, sRow, "" & .Fields("qty").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, , True)
            End If
            .Close
        End If
    End With
        
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    
End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim sRow As Long, sData As Variant, sCode As Long, sReturn As Boolean

    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData:        sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                If brJob Then
                    cTstk.testcd = Trim(txtTestCd.Text)
                    cTstk.stkcd = sCode
                    .GetText 6, sRow, sData:    cTstk.qty = Val(sData)
                    
                    sReturn = cTstk.cfSave
                Else
                    sReturn = cTstk.cfDelete(Trim(txtTestCd.Text), sCode)
                End If
                
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

Private Sub cmdClear_Click()

    Call gsSetStkTree(trvStkList, gDelYes)
    Call gsSetTestTree(trvTestList, gDelYes)
    
    Call gsSpreadClear(spList, , True)
    txtTestCd.Text = ""
    lblTestNm.Caption = ""
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    
    txtTestCd.SetFocus

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
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

    Me.Width = 19000
    Me.Height = 10120
    
    Set cTstk = New clsTestStk
   
    Me.Show
    
    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
    
        If Col = 2 Then
            spList.GetText 2, Row, sData
            gSql = "select stkcd, stknm, stkspec, iounit from mstSTK where stkcd = " & Val(sData)
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        spList.SetText 3, Row, "" & .Fields("stknm").Value
                        spList.SetText 4, Row, "" & .Fields("stkspec").Value
                        spList.SetText 5, Row, "" & .Fields("iounit").Value
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

Private Sub trvTestList_Collapse(Node As SSActiveTreeView.SSNode)

    If Node.Level = 1 Then Node.Image = "close"

End Sub

Private Sub trvTestList_DblClick()

    If trvTestList.SelectedNodes.Item(1).Level > 1 Then
        txtTestCd.Text = trvTestList.SelectedItem.Key
        lblTestNm.Caption = trvTestList.SelectedItem.Text
        
        Call psListRefresh
    End If

End Sub

Private Sub trvTestList_Expand(Node As SSActiveTreeView.SSNode)

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
                Exit For
            ElseIf Trim(sData) = trvStkList.SelectedItem.Key Then
                MsgBox "등록된 품번입니다.!", vbCritical
                Exit For
            End If
        Next sRow
    End With

End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Len(txtTestCd.Text) > 0 Then
        lblTestNm.Caption = "[" & Trim(txtTestCd.Text) & "] " & gfTestName(txtTestCd.Text)
    End If

End Sub

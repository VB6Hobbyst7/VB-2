VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm장비별운영시약기초 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장비별 운영시약 기초자료"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15600
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm장비별운영시약기초.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   15600
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   13520
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm장비별운영시약기초.frx":000C
      Begin Threed.SSPanel SSPanel3 
         Height          =   525
         Left            =   5235
         TabIndex        =   4
         Top             =   750
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   926
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboOper 
            Height          =   300
            Left            =   7080
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   120
            Width           =   3135
         End
         Begin PVCOMBOLibCtl.PVComboBox cboMach 
            Height          =   300
            Left            =   1350
            TabIndex        =   5
            Top             =   120
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
            Column0.Heading =   "분류코드"
            Column0.Width   =   30
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   ""
            Column0.Format  =   ""
            Column0.Bound   =   0   'False
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "분류명"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "장비번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblMachNm 
            Height          =   300
            Left            =   2340
            TabIndex        =   7
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   300
            Left            =   5850
            TabIndex        =   9
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "운영내역"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin SSActiveTreeView.SSTree trvStkList 
         Height          =   6885
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   12144
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
         Height          =   615
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   5100
         _ExtentX        =   8996
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
         Caption         =   " ▒ 시약리스트"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   5235
         TabIndex        =   2
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
         Caption         =   " ▒ 운영시약 소요량"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   5565
         Left            =   5235
         TabIndex        =   10
         Top             =   1380
         Width           =   10335
         _Version        =   524288
         _ExtentX        =   18230
         _ExtentY        =   9816
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
         SpreadDesigner  =   "frm장비별운영시약기초.frx":00DE
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   585
         Left            =   5235
         TabIndex        =   11
         Top             =   7050
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1032
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   6780
            TabIndex        =   12
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
            TabIndex        =   13
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
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   9000
            TabIndex        =   15
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
   End
End
Attribute VB_Name = "frm장비별운영시약기초"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cMstk As clsMachStk

Private Sub psListRefresh()
Dim sRow As Long

    If cboMach.ListIndex > 0 And cboOper.ListIndex > 0 Then
        With cMstk.cfList(cboMach.Text, cboOper.ItemData(cboOper.ListIndex))
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
    End If
    
End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim sRow As Long, sData As Variant, sCode As Long, sReturn As Boolean

    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData:        sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                If brJob Then
                    cMstk.machcd = Trim(cboMach.Text)
                    cMstk.opercd = cboOper.ItemData(cboOper.ListIndex)
                    cMstk.stkcd = sCode
                    .GetText 6, sRow, sData:    cMstk.qty = Val(sData)
                    
                    sReturn = cMstk.cfSave
                Else
                    sReturn = cMstk.cfDelete(Trim(cboMach.Text), cboOper.ItemData(cboOper.ListIndex), sCode)
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

Private Sub cboMach_Click()

    If cboMach.ListIndex >= 0 Then
        lblMachNm.Caption = cboMach.SubItem(cboMach.ListIndex, 1)
        Call psListRefresh
    End If

End Sub

Private Sub cboOper_Click()

    If cboOper.ListIndex > 0 Then
        Call psListRefresh
    End If

End Sub

Private Sub cmdClear_Click()

    Call gsSetStkTree(trvStkList, gDelYes)
    
    cboMach.ListIndex = -1
    cboOper.ListIndex = -1
    
    Call gsSpreadClear(spList, , True)
    lblMachNm.Caption = ""
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    
    cboMach.SetFocus

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

    Set cMstk = New clsMachStk
    
    Me.Show
    
    Call gsSetOperCombo(cboOper, False)
    Call gsSetMachComboPV(cboMach, False)
    
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

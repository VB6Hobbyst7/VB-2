VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm장비운영내역서 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장비운영내역서"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13170
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
   ScaleHeight     =   9015
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   15901
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm장비운영내역서.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   615
         Left            =   8475
         TabIndex        =   1
         Top             =   30
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
            TabIndex        =   2
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
            TabIndex        =   3
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
            TabIndex        =   4
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
            TabIndex        =   5
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
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "운영일자"
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
            Caption         =   "운영장비"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PVCOMBOLibCtl.PVComboBox cboMach 
            Height          =   300
            Left            =   3930
            TabIndex        =   11
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
         Begin Threed.SSPanel lblMachNm 
            Height          =   300
            Left            =   4920
            TabIndex        =   12
            Top             =   150
            Width           =   3285
            _ExtentX        =   5794
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
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   8235
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   13110
         _Version        =   524288
         _ExtentX        =   23125
         _ExtentY        =   14526
         _StockProps     =   64
         ButtonDrawMode  =   4
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
         MaxCols         =   8
         SpreadDesigner  =   "frm장비운영내역서.frx":0072
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm장비운영내역서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long, sDate As String

    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    gSql = "select a.*, b.opernm, c.usernm from operL a with (nolock), mstOPER b, mstUSER c" & _
           " where a.machcd = '" & cboMach.Text & "' and a.operdt = '" & sDate & "' and a.opercd = b.opercd and a.usercd = c.usercd" & _
           " order by a.operseq"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, .Fields("opernm").Value & Space(.Fields("opernm").DefinedSize - HLen(.Fields("opernm").Value) + 30) & " | " & .Fields("opercd").Value
                    spList.SetText 3, sRow, "" & .Fields("opercnt").Value
                    spList.SetText 4, sRow, "" & .Fields("reason").Value
                    spList.SetText 5, sRow, gEndFlag("" & .Fields("endfg").Value)
                    spList.SetText 6, sRow, "" & .Fields("usernm").Value
                    spList.SetText 7, sRow, "" & .Fields("moddt").Value
                    spList.SetText 8, sRow, Val("" & .Fields("operseq").Value)
                    
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
Dim cOpl As clsOperList
Dim sRow As Long, sData As Variant, sDate As String, sReturn As Boolean, sCode() As String, sSeq As Integer
    
    sReturn = True
    
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    
    Set cOpl = New clsOperList
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 2, sRow, sData:            sCode = Split(sData, "|")
            .GetText 1, sRow, sData
            If Val(sData) > 0 And UBound(sCode) > 0 Then
                Call cDb.csBegin
                If brJob Then
                    cOpl.operdt = sDate
                    cOpl.machcd = cboMach.Text
                    cOpl.operseq = sSeq
                    cOpl.opercd = Trim(sCode(1))
                    cOpl.usercd = gUserId
                    .GetText 3, sRow, sData:            cOpl.opercnt = Val(sData)
                    .GetText 4, sRow, sData:            cOpl.reason = Trim(sData)
                    sReturn = cOpl.cfSave
                    If sReturn Then
                        .SetText .MaxCols, sRow, cOpl.operseq
                    End If
                Else
                    If sSeq > 0 Then
                        sReturn = cOpl.cfDelete(cboMach.Text, sDate, sSeq)
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

Private Sub cboMach_Click()

    If cboMach.ListIndex >= 0 Then
        lblMachNm.Caption = cboMach.SubItem(cboMach.ListIndex, 1)
        
        Call psListRefresh
        
        cmdSave.Enabled = True
    Else
        lblMachNm.Caption = ""
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()
Dim cOper As clsMstOper, sStr As String

    Call gsSpreadClear(spList, , True)

    dtpDate.Value = gfSystemDate
    Call gsSetMachComboPV(cboMach)
    
    cboMach.ListIndex = -1
    lblMachNm.Caption = ""
    
    Set cOper = New clsMstOper
    With cOper.cfList
        If .State = adStateOpen Then
            While (Not .EOF)
                sStr = sStr & .Fields("opernm").Value & Space(.Fields("opernm").DefinedSize - HLen(.Fields("opernm").Value) + 30) & " | " & .Fields("opercd").Value
                
                .MoveNext
                If Not .EOF Then
                    sStr = sStr & vbTab
                End If
            Wend
            .Close
        End If
    End With
    
    spList.Row = -1
    spList.Col = 2
    spList.TypeComboBoxList = sStr
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 장비운영자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
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
    
    Me.KeyPreview = True
    Me.Show
    
    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String, sAmt As Currency, sQty As Single

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
    End If

End Sub


VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm물품요청서분류 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "물품요청서(분류별)"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15840
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15840
      _ExtentX        =   27940
      _ExtentY        =   16695
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm물품요청서분류.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   630
         Left            =   12285
         TabIndex        =   1
         Top             =   30
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1111
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
         Height          =   630
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboKind 
            Height          =   300
            Left            =   8250
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   150
            Width           =   3045
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
            Caption         =   "요청일자"
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
            Caption         =   "요청부서"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   300
            Left            =   7020
            TabIndex        =   12
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "물품분류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   8670
         Left            =   30
         TabIndex        =   10
         Top             =   765
         Width           =   15780
         _Version        =   524288
         _ExtentX        =   27834
         _ExtentY        =   15293
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
         MaxCols         =   11
         SpreadDesigner  =   "frm물품요청서분류.frx":0072
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm물품요청서분류"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboKind_Click()
Dim cStk As clsMstStk, sRow As Long, sDate As String, sRmdQty As Double

    If cboKind.ListIndex > 0 Then
        Set cStk = New clsMstStk
        
        sDate = Format(dtpDate.Value, "yyyy-MM-dd")
        
        With cStk.cfList(gDelNo, cboKind.ItemData(cboKind.ListIndex), , False)
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
                        sRmdQty = gfPresentStkRmd(.Fields("stkcd").Value, sDate, False)
                        spList.SetText 6, sRow, sRmdQty
                        If Val("" & .Fields("safeqty").Value) > sRmdQty Then
                            spList.SetText 7, sRow, (Val("" & .Fields("safeqty").Value) - sRmdQty)
                        Else
                            spList.SetText 7, sRow, ""
                        End If
                        spList.SetText 8, sRow, ""
                        spList.SetText 9, sRow, DateAdd("d", Val("" & .Fields("buyday").Value), sDate)
                        spList.SetText 10, sRow, ""
                        
                        .MoveNext
                    Wend
                    
                    cmdSave.Enabled = True
                Else
                    Call gsSpreadClear(spList, 0, True)
                    cmdSave.Enabled = False
                End If
                .Close
            End If
        End With
    Else
        Call gsSpreadClear(spList, 0, True)
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()

    Call gsSpreadClear(spList, , True)

    dtpDate.Value = gfSystemDate
    Call gsSetDutyCombo(cboDuty)
    Call gsSetKindCombo(cboKind, False)
    
    With spList
        .Row = -1
        .Col = 8
        .ForeColor = vbRed
    End With
    
    cmdSave.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()
Dim cReq As clsReqList
Dim sRow As Long, sData As Variant, sDate As String, sSeq As Integer, sCode As Long, sReturn As Boolean

    MousePointer = vbHourglass
    If cboDuty.ListIndex >= 0 Then
        Set cReq = New clsReqList
        sDate = Format(dtpDate.Value, "yyyy-MM-dd")
        With spList
            For sRow = 1 To .MaxRows
                .GetText 2, sRow, sData:            sCode = Val(sData)
                .GetText 1, sRow, sData
                If Val(sData) > 0 And sCode > 0 Then
                    cReq.dutycd = cboDuty.ItemData(cboDuty.ListIndex)
                    cReq.reqdt = sDate
                    cReq.reqseq = sSeq
                    cReq.stkcd = sCode
                    .GetText 8, sRow, sData:            cReq.qty = Val(sData)
                    .GetText 9, sRow, sData:            cReq.lastdt = Trim(sData)
                    .GetText 10, sRow, sData:           cReq.remark = Trim(sData)
                    cReq.usercd = gUserId
                    
                    sReturn = cReq.cfSave
                    If sReturn Then
                        .SetText .MaxCols, sRow, cReq.reqseq
                    End If
                    
                    If sReturn = False Then
                        Exit For
                    Else
                        .SetText 1, sRow, ""
                    End If
                End If
            Next sRow
        End With
    Else
        MsgBox "요청부서를 선택하세요.!", vbCritical
    End If
    MousePointer = vbDefault

End Sub

Private Sub Form_Load()

    Me.Show

    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
    End If
    
End Sub

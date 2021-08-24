VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm품목별재고현황 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "품목별 재고현황"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15765
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
   ScaleHeight     =   9465
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   16695
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm품목별재고현황.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   630
         Left            =   12150
         TabIndex        =   1
         Top             =   30
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   2340
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
            Caption         =   "닫기(&X)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClear 
            Height          =   420
            Left            =   120
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
            Caption         =   "화면지움"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdExport 
            Height          =   420
            Left            =   1230
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
            Caption         =   "자료변환"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   630
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComDlg.CommonDialog cdgExport 
            Left            =   11100
            Top             =   60
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox cboKind 
            Height          =   300
            Left            =   3630
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   150
            Width           =   3765
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   150
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   60882947
            CurrentDate     =   41061
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   90
            TabIndex        =   7
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "재고년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   300
            Left            =   2400
            TabIndex        =   8
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
         TabIndex        =   9
         Top             =   765
         Width           =   15705
         _Version        =   524288
         _ExtentX        =   27702
         _ExtentY        =   15293
         _StockProps     =   64
         ColHeaderDisplay=   0
         EditEnterAction =   2
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
         SpreadDesigner  =   "frm품목별재고현황.frx":0072
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "frm품목별재고현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboKind_Click()
Dim sRow As Long, sDate As String, sQty As Double, sRmdQty As Double, sRateQty As Integer

    MousePointer = vbHourglass
    If cboKind.ListIndex > 0 Then
        sDate = Format(dtpDate.Value, "yyyy-MM")
        
        gSql = "select b.*, a.stkcd, a.stknm, a.stkspec, a.buyunit, a.iounit, a.buyioqty from stkRMD b with (nolock), mstSTK a" & _
               " where a.kindcd = '" & cboKind.ItemData(cboKind.ListIndex) & "' and a.stkcd *= b.stkcd and b.rmdym = '" & sDate & "'" & _
               " order by a.stknm"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    Call gsSpreadClear(spList, .RecordCount, True)
                    While (Not .EOF)
                        sRow = sRow + 1
                        
                        sRmdQty = Val("" & .Fields("prevqty").Value) + Val("" & .Fields("inqty").Value) - Val("" & .Fields("outqty").Value)
                        
                        spList.SetText 1, sRow, "" & .Fields("stkcd").Value
                        spList.SetText 2, sRow, "" & .Fields("stknm").Value
                        spList.SetText 3, sRow, "" & .Fields("stkspec").Value
                        spList.SetText 4, sRow, "" & .Fields("buyunit").Value
                        
                        sRateQty = Val("" & .Fields("buyioqty").Value)
                        If sRateQty = 0 Then sRateQty = 1
                        
                        sQty = Val("" & .Fields("prevqty").Value) / sRateQty
                        spList.SetText 5, sRow, Format(sQty, "#,##0.0")
                        sQty = sRmdQty / sRateQty
                        spList.SetText 6, sRow, Format(sQty, "#,##0.0")
                        
                        spList.SetText 8, sRow, "" & .Fields("iounit").Value
                        spList.SetText 9, sRow, Format(Val("" & .Fields("prevqty").Value), "#,##0.0")
                        spList.SetText 10, sRow, Format(Val("" & .Fields("inqty").Value), "#,##0.0")
                        spList.SetText 11, sRow, Format(Val("" & .Fields("outqty").Value), "#,##0.0")
                        spList.SetText 12, sRow, Format(sRmdQty, "#,##0.0")
                        
                        .MoveNext
                    Wend
                    cmdExport.Enabled = True
                Else
                    Call gsSpreadClear(spList, 0, True)
                    cmdExport.Enabled = False
                End If
                .Close
            End If
        End With
    Else
        Call gsSpreadClear(spList, , True)
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Call cboKind_Click
    End If

End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    Me.Show

    dtpDate.Value = gfSystemDate
    Call gsSetKindCombo(cboKind, False)
    
    cmdExport.Enabled = False

End Sub

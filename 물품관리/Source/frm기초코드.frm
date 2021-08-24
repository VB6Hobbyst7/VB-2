VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Begin VB.Form frm기초코드 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "기초코드등록"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7335
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
   ScaleHeight     =   7155
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12621
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm기초코드.frx":0000
      Begin ActiveTabs.SSActiveTabs sstMain 
         Height          =   6405
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   11298
         _Version        =   262144
         TabCount        =   3
         BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tabs            =   "frm기초코드.frx":0052
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
            Height          =   6015
            Left            =   30
            TabIndex        =   4
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   10610
            _Version        =   262144
            TabGuid         =   "frm기초코드.frx":0103
            Begin FPUSpreadADO.fpSpread spOper 
               Height          =   5745
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   6945
               _Version        =   524288
               _ExtentX        =   12250
               _ExtentY        =   10134
               _StockProps     =   64
               ButtonDrawMode  =   4
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
               MaxCols         =   4
               SpreadDesigner  =   "frm기초코드.frx":012B
               AppearanceStyle =   0
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   6015
            Left            =   30
            TabIndex        =   3
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   10610
            _Version        =   262144
            TabGuid         =   "frm기초코드.frx":0D92
            Begin FPUSpreadADO.fpSpread spKind 
               Height          =   5745
               Left            =   120
               TabIndex        =   6
               Top             =   120
               Width           =   6945
               _Version        =   524288
               _ExtentX        =   12250
               _ExtentY        =   10134
               _StockProps     =   64
               ButtonDrawMode  =   4
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
               MaxCols         =   4
               SpreadDesigner  =   "frm기초코드.frx":0DBA
               AppearanceStyle =   0
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   6015
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   10610
            _Version        =   262144
            TabGuid         =   "frm기초코드.frx":1A1B
            Begin FPUSpreadADO.fpSpread spDuty 
               Height          =   5745
               Left            =   120
               TabIndex        =   5
               Top             =   120
               Width           =   6945
               _Version        =   524288
               _ExtentX        =   12250
               _ExtentY        =   10134
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
               MaxCols         =   3
               SpreadDesigner  =   "frm기초코드.frx":1A43
               AppearanceStyle =   0
            End
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   585
         Left            =   30
         TabIndex        =   7
         Top             =   6540
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   1032
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   3780
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
            Left            =   4890
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
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   6000
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
   End
End
Attribute VB_Name = "frm기초코드"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cDuty As clsMstDuty, cKind As clsMstStkkind, cOper As clsMstOper

Private Sub psDutyRefresh()
Dim sRow As Long

    With cDuty.cfList
        If .State = adStateOpen Then
            sRow = 0
            If Not .EOF Then
                Call gsSpreadClear(spDuty, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spDuty.SetText 1, sRow, ""
                    spDuty.SetText 2, sRow, "" & .Fields("dutycd").Value
                    spDuty.SetText 3, sRow, "" & .Fields("dutynm").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spDuty, , True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psKindRefresh()
Dim sRow As Long

    With cKind.cfList
        If .State = adStateOpen Then
            sRow = 0
            If Not .EOF Then
                Call gsSpreadClear(spKind, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spKind.SetText 1, sRow, ""
                    spKind.SetText 2, sRow, "" & .Fields("kindcd").Value
                    spKind.SetText 3, sRow, "" & .Fields("kindnm").Value
                    
                    spKind.Row = sRow:      spKind.Col = 4
                    spKind.TypeComboBoxCurSel = Val("" & .Fields("reagentfg").Value)
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spKind, , True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psOperRefresh()
Dim sRow As Long

    With cOper.cfList
        If .State = adStateOpen Then
            sRow = 0
            If Not .EOF Then
                Call gsSpreadClear(spOper, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spOper.SetText 1, sRow, ""
                    spOper.SetText 2, sRow, "" & .Fields("opercd").Value
                    spOper.SetText 3, sRow, "" & .Fields("opernm").Value
                    
                    spOper.Row = sRow:      spOper.Col = 4
                    spOper.TypeComboBoxCurSel = Val("" & .Fields("operfg").Value)
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spOper, , True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psDutyProcess(ByVal brJob As Boolean)
Dim sRow As Long, sData As Variant, sReturn As Boolean, sCode As String

    With spDuty
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData:        sCode = Trim(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And Len(sCode) > 0 Then
                If brJob Then
                    cDuty.dutycd = sCode
                    .GetText 3, sRow, sData
                    cDuty.dutynm = Trim(sData)
                    
                    sReturn = cDuty.cfSave
                Else
                    sReturn = cDuty.cfDelete(sCode)
                End If
                
                If sReturn = False Then
                    Exit For
                End If
                
                .SetText 1, sRow, ""
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psDutyRefresh
    End If

End Sub

Private Sub psKindProcess(ByVal brJob As Boolean)
Dim sRow As Long, sData As Variant, sReturn As Boolean, sCode As String

    With spKind
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData:        sCode = Trim(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And Len(sCode) > 0 Then
                If brJob Then
                    cKind.kindcd = sCode
                    .GetText 3, sRow, sData
                    cKind.kindnm = Trim(sData)
                    
                    .Row = sRow:    .Col = 4
                    cKind.reagentfg = .TypeComboBoxCurSel
                    
                    sReturn = cKind.cfSave
                Else
                    sReturn = cKind.cfDelete(sCode)
                End If
                
                If sReturn = False Then
                    Exit For
                End If
                
                .SetText 1, sRow, ""
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psKindRefresh
    End If

End Sub

Private Sub psOperProcess(ByVal brJob As Boolean)
Dim sRow As Long, sData As Variant, sReturn As Boolean, sCode As String

    With spOper
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData:        sCode = Trim(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And Len(sCode) > 0 Then
                If brJob Then
                    cOper.opercd = sCode
                    .GetText 3, sRow, sData
                    cOper.opernm = Trim(sData)
                    
                    .Row = sRow:    .Col = 4
                    cOper.operfg = .TypeComboBoxCurSel
                    
                    sReturn = cOper.cfSave
                Else
                    sReturn = cOper.cfDelete(sCode)
                End If
                
                If sReturn = False Then
                    Exit For
                End If
                
                .SetText 1, sRow, ""
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psOperRefresh
    End If

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 기초코드를 삭제하시겠습니까 ?", vbQuestion + vbYesNo) = vbYes Then
        Select Case sstMain.SelectedTab.Index
            Case 1:     Call psDutyProcess(False)
            Case 2:     Call psKindProcess(False)
            Case 3:     Call psOperProcess(False)
        End Select
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Select Case sstMain.SelectedTab.Index
        Case 1:     Call psDutyProcess(True)
        Case 2:     Call psKindProcess(True)
        Case 3:     Call psOperProcess(True)
    End Select
    MousePointer = vbDefault

End Sub

Private Sub Form_Load()

    Me.Show

    Set cDuty = New clsMstDuty
    Set cKind = New clsMstStkkind
    Set cOper = New clsMstOper
    
    spOper.Row = -1
    spOper.Col = 4
    spOper.TypeComboBoxList = Replace(gOperFlagStr, "|", vbTab)
    
    Call psDutyRefresh
    Call psKindRefresh
    Call psOperRefresh

End Sub

Private Sub spDuty_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        With spDuty
            .SetText 1, Row, 1
        End With
    End If
    
End Sub

Private Sub spKind_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        With spKind
            .SetText 1, Row, 1
        End With
    End If

End Sub

Private Sub spOper_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        With spOper
            .SetText 1, Row, 1
        End With
    End If
    
End Sub
'
'Private Sub sstMain_TabClick(ByVal NewTab As ActiveTabs.SSTab)
'
'    Select Case NewTab.Index
'        Case 1:     Call psDutyRefresh
'        Case 2:     Call psKindRefresh
'        Case 3:     Call psOperRefresh
'    End Select
'
'End Sub

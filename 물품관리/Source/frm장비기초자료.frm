VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm장비기초자료 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검사장비기초자료"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10020
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
   ScaleHeight     =   8475
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   14949
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm장비기초자료.frx":0000
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   7125
         Left            =   30
         TabIndex        =   1
         Top             =   645
         Width           =   9960
         _Version        =   524288
         _ExtentX        =   17568
         _ExtentY        =   12568
         _StockProps     =   64
         ButtonDrawMode  =   4
         EditEnterAction =   5
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
         MaxCols         =   5
         SpreadDesigner  =   "frm장비기초자료.frx":0072
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   510
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   900
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
         Caption         =   " ▒ 검사장비 기초자료"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   7875
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   6390
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
            Caption         =   "저장(&S)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   420
            Left            =   7500
            TabIndex        =   5
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
            Left            =   8610
            TabIndex        =   6
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
Attribute VB_Name = "frm장비기초자료"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cMch As clsMstMach

Private Sub psListRefresh()
Dim sRow As Long

    With cMch.cfList
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 1000, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, .Fields("machcd").Value
                    spList.SetText 3, sRow, .Fields("machnm").Value
                    spList.SetText 4, sRow, .Fields("dutynm").Value & Space(.Fields("dutynm").DefinedSize - HLen(.Fields("dutynm").Value)) & " | " & .Fields("dutycd").Value
                    spList.SetText 5, sRow, .Fields("delfg").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, , True)
            End If
            .Close
        End If
    End With
        
End Sub

Private Sub psDataProcess(ByVal brJob As Boolean)
Dim sRow As Long, sReturn As Boolean, sData As Variant, sCode As String, sDuty() As String

    With spList
        For sRow = 1 To .MaxRows
            .Row = sRow
            .GetText 2, sRow, sData:        sCode = Trim(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And Len(sCode) > 0 Then
                If brJob Then
                    cMch.machcd = sCode
                    .GetText 3, sRow, sData:    cMch.machnm = Trim(sData)
                    .Col = 4
                    If .TypeComboBoxCurSel >= 0 Then
                        .GetText 4, sRow, sData:    sDuty = Split(sData, "|")
                        cMch.dutycd = Trim(sDuty(1))
                    Else
                        cMch.dutycd = ""
                    End If
                    .GetText 5, sRow, sData:    cMch.delfg = Val(sData)
                    
                    sReturn = cMch.cfSave
                Else
                    sReturn = cMch.cfDelete(sCode)
                End If
                
                If sReturn Then
                    .SetText 1, sRow, ""
                Else
                    Exit For
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psListRefresh
    End If
        
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
Dim cDuty As clsMstDuty, sStr As String

    Set cMch = New clsMstMach
    
    Me.Show
    
    Set cDuty = New clsMstDuty
    With cDuty.cfList
        If .State = adStateOpen Then
            If Not .EOF Then
                sStr = .Fields("dutynm").Value & Space(.Fields("dutynm").DefinedSize - HLen(.Fields("dutynm").Value)) & " | " & .Fields("dutycd").Value
                .MoveNext
                While (Not .EOF)
                    sStr = sStr & vbTab
                    sStr = sStr & .Fields("dutynm").Value & Space(.Fields("dutynm").DefinedSize - HLen(.Fields("dutynm").Value)) & " | " & .Fields("dutycd").Value
                    
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    
    spList.Row = -1
    spList.Col = 4
    spList.TypeComboBoxList = sStr
    
    Call psListRefresh
    
End Sub

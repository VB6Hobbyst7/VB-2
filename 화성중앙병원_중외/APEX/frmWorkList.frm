VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "워크리스트"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   14880
   Icon            =   "frmWorkList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   14880
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      Caption         =   "Hidden"
      Height          =   2085
      Left            =   1860
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Frame fraBIT 
         BackColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   150
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox txtFrNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Text            =   "0000"
            Top             =   180
            Width           =   765
         End
         Begin VB.TextBox txtToNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1080
            TabIndex        =   17
            Text            =   "0999"
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   19
            Top             =   240
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F8E4D8&
         Height          =   555
         Left            =   150
         TabIndex        =   12
         Top             =   810
         Visible         =   0   'False
         Width           =   3195
         Begin VB.CommandButton cmdSeq 
            Caption         =   "Seq 매치"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1830
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtSeqNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1110
            TabIndex        =   13
            Text            =   "1"
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lblSeqNo 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "시작 Seq"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   90
            TabIndex        =   15
            Top             =   210
            Width           =   750
         End
      End
      Begin HSCotrol.CButton cmdSend 
         Height          =   375
         Left            =   2460
         TabIndex        =   25
         Top             =   330
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "화면전송"
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmWorkList.frx":06C2
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14880
      TabIndex        =   0
      Top             =   0
      Width           =   14880
      Begin VB.Frame fraBrain 
         BackColor       =   &H00F8E4D8&
         Height          =   495
         Left            =   12300
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "완료"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   2
            Left            =   1740
            TabIndex        =   23
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "대기"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   1
            Left            =   930
            TabIndex        =   22
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "전체"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         Caption         =   "저장된 결과를 포함하여 조회"
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   9480
         TabIndex        =   1
         Top             =   150
         Width           =   2865
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   125501441
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2340
         TabIndex        =   3
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   125501441
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   375
         Left            =   8250
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "닫기"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmWorkList.frx":081C
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   375
         Left            =   3810
         TabIndex        =   5
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   15698777
         Caption         =   "워크조회"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmWorkList.frx":0DB6
      End
      Begin HSCotrol.CButton cmdSendClose 
         Height          =   375
         Left            =   6030
         TabIndex        =   6
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "전송/닫기"
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmWorkList.frx":0F10
      End
      Begin HSCotrol.CButton cmdWorkPrint 
         Height          =   375
         Left            =   7140
         TabIndex        =   7
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "워크출력"
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmWorkList.frx":106A
      End
      Begin HSCotrol.CButton cmdOrderSend 
         Height          =   435
         Left            =   4920
         TabIndex        =   24
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         BackColor       =   32768
         Caption         =   "오더전송"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":11C4
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2130
         TabIndex        =   9
         Top             =   180
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "기간"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   180
         Width           =   360
      End
   End
   Begin FPSpreadADO.fpSpread spdWork 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   150
      TabIndex        =   10
      Tag             =   "20001"
      Top             =   720
      Width           =   16020
      _Version        =   524288
      _ExtentX        =   28258
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   23
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmWorkList.frx":1420
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
      CellNoteIndicatorColor=   16576
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOrderSend_Click()
    Dim intRow      As Integer
    Dim STM         As ADODB.Stream
    Dim blnSendXml  As Boolean
    Dim strHeader   As String
    Dim strBody     As String
    Dim strFileNm   As String
    Dim strAssayNm  As String
    
    blnSendXml = False
    
    If spdWork.MaxRows < 1 Then
        MsgBox "검사대상자가 없습니다.", vbOKOnly + vbCritical, Me.Caption
    Else
    
        strHeader = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
        strHeader = strHeader & "<GROUP NAME=""WORKLIST"" TYPE=""HOST"" VERSION=""00.00.00.00"">" & vbCrLf
        strBody = ""
        
        With spdWork
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = colCHECKBOX
                If .Value = 1 Then
                    If Trim(GetText(spdWork, intRow, colBARCODE)) <> "" Then
                        blnSendXml = True
                        
                        strBody = strBody & vbTab & "<GROUP TYPE=""Sample"">" & vbCrLf
                        strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Patient"" ID=""" & Trim(GetText(spdWork, intRow, colBARCODE)) & """>" & vbCrLf
                        '-- Group
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Group"" FILTER=""-wwwr"">" & Trim(GetText(spdWork, intRow, colINOUT)) & "</PARAM>" & vbCrLf
                        '-- Name
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Name"" FILTER=""-ooor"">" & Trim(GetText(spdWork, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                        '-- Surname
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Surname"" FILTER=""-ooor"">" & Trim(GetText(spdWork, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                        '-- BirthDate
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""Date"" ID=""BirthDate"" FILTER=""-ooor"">" & Format(Now, "yyyy-mm-dd") & "</PARAM>" & vbCrLf
                        '-- Sex
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""SexType"" ID=""Sex"" FILTER=""-wwwr"">" & IIf(Trim(GetText(spdWork, intRow, colPSEX)) = "M", "M", "F") & "</PARAM>" & vbCrLf
                        '-- Note
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[50]"" ID=""Note"" FILTER=""-wwwr"">" & "Note Test" & "</PARAM>" & vbCrLf
                        '-- Code (장비에서 Patient ID)
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[50]"" ID=""Note"" FILTER=""-wwwr"">" & Trim(GetText(spdWork, intRow, colPID)) & "</PARAM>" & vbCrLf
                        strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[25]"" ID=""Code"" FILTER=""-ooor"">" & "Code Test" & "</PARAM>" & vbCrLf
                        
                        strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
                        strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Assays"" ID="""">" & vbCrLf
                        'Select Case Trim(GetText(spdWork, intRow, colINOUT))
                        '    Case "ATOPY":       strAssayNm = gB4C.ASSAYATNM
                        '    Case "FOOD":        strAssayNm = gB4C.ASSAYFDNM
                        '    Case "INHALANT":    strAssayNm = gB4C.ASSAYINNM
                        ''    Case "96M":         strAssayNm = gB4C.ASSAY96NM
                        'End Select
                        strAssayNm = gB4C.ASSAY96NM
                        strBody = strBody & vbTab & vbTab & vbTab & "<GROUP TYPE=""Assay"" ID=""" & strAssayNm & """ STATE=""Host"" />" & vbCrLf
                        strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
                        strBody = strBody & vbTab & "</GROUP>" & vbCrLf
                        
                        .Row = intRow
                        .Col = colCHECKBOX
                        .Value = "0"
                        
                        Call SetText(spdWork, "오더전송", intRow, colSTATE)
                    End If
                End If
            Next
        End With
        
        If blnSendXml = True Then
            '## 기존에 파일이 있으면 삭제
            strFileNm = gB4C.IMPORT & "\HostIn.xml"
        
            If Dir$(strFileNm, vbNormal) <> "" Then
                Kill strFileNm
            End If
            
             '## 파일오픈
            Set STM = New ADODB.Stream
            
            STM.Open
            STM.Type = adTypeText
            STM.Charset = "utf-8"
            STM.WriteText strHeader & strBody & "</GROUP>" & vbCrLf
                        
            STM.SaveToFile strFileNm, adSaveCreateNotExist
            STM.Close
            Set STM = Nothing
            
            MsgBox "오더 파일이 생성되었습니다", vbOKOnly + vbInformation, Me.Caption
            
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    
    If gDBTYPE <> "99" Then
        Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork, _
                        Format(txtFrNo.Text, "0000"), Format(txtToNo.Text, "0000"), IIf(chkSave.Value = "1", True, False))
        
    Else
        Dim i As Integer
    
        With spdWork
            .MaxRows = 10
            For i = 1 To 10
                Call SetText(spdWork, "1", i, colCHECKBOX)
                Call SetText(spdWork, Format(dtpTo.Value, "yyyy-mm-dd"), i, colHOSPDATE)
                Call SetText(spdWork, Format(dtpTo.Value, "mmddhhmmss") & CStr(i), i, colBARCODE)
                Call SetText(spdWork, "오세원" & CStr(i), i, colPNAME)
                'Call SetText(spdWork, "BLD/BIL/URO/KET/PRO/NIT/GLU/pH/S.G/LEU", i, colITEMS)
                Call SetText(spdWork, "BLD/BIL/URO/PRO/NIT/pH/S.G", i, colITEMS)
            Next
            .RowHeight(-1) = gROWHEIGHT ' = 13
        End With
    
        spdWork.RowHeight(-1) = gROWHEIGHT
        Call Form_Resize
    End If

End Sub

Private Sub cmdSend_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To frmInterface.spdOrder.MaxRows
                    frmInterface.spdOrder.Row = intORow
                    frmInterface.spdOrder.Col = colBARCODE
                    If strBarno = GetText(frmInterface.spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    frmInterface.spdOrder.MaxRows = frmInterface.spdOrder.MaxRows + 1
                    intRow = frmInterface.spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                                frmInterface.spdOrder.Row = 0
                                frmInterface.spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(frmInterface.spdOrder.Text) Then
                                    Call SetSPDOrder(frmInterface.spdOrder, intRow, intRow, intOCol, intOCol)
                                End If
                            Next
                        Next
                    Next
                    
                    'Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, colITEMS), intRow, colITEMS)
                    frmInterface.spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdSendClose_Click()
    
    Call cmdSend_Click
    
    Call cmdClose_Click
    
    frmInterface.ZOrder 0
    
End Sub

Private Sub cmdSeq_Click()
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim strSeq          As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                For intORow = 1 To frmInterface.spdOrder.MaxRows
                    If GetText(spdWork, intWRow, colSEQNO) = GetText(frmInterface.spdOrder, intORow, colSEQNO) Then
                        
                        Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                        
                        DoEvents
                        
                        If GetSampleInfo(intORow, frmInterface.spdOrder) = -1 Then
                            'MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                        Else
                            '정보수정
                            SQL = ""
                            SQL = SQL & "UPDATE PATRESULT SET "
                            SQL = SQL & "  BARCODE       = '" & Trim(GetText(frmInterface.spdOrder, intORow, colBARCODE)) & "'" & vbCrLf
                            SQL = SQL & " ,INOUT         = '" & Trim(GetText(frmInterface.spdOrder, intORow, colINOUT)) & "'" & vbCrLf
                            SQL = SQL & " ,CHARTNO       = '" & Trim(GetText(frmInterface.spdOrder, intORow, colCHARTNO)) & "'" & vbCrLf
                            SQL = SQL & " ,PID           = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPID)) & "'" & vbCrLf
                            SQL = SQL & " ,PNAME         = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPNAME)) & "'" & vbCrLf
                            SQL = SQL & " ,PSEX          = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPSEX)) & "'" & vbCrLf
                            SQL = SQL & " ,PAGE          = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPAGE)) & "'" & vbCrLf
                            SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(frmInterface.spdOrder, intORow, colEXAMDATE)) & "'" & vbCrLf
                            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(frmInterface.spdOrder, intORow, colSAVESEQ)) & vbCrLf
                            SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        Exit For
                    End If
                Next intORow
            End If
        Next intWRow
    End With
    
End Sub

Private Sub cmdWorkPrint_Click()
    
    If spdWork.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        spdWork.PrintOrientation = PrintOrientationPortrait     '세로출력
        'spdWork.PrintOrientation = PrintOrientationLandscape    '가로출력
        spdWork.Action = 13
    End If
    

End Sub

Private Sub Form_Load()
    Dim intCol      As Integer
    Dim intColWidth As Integer
    
    Call CtlInitializing

    '-- 컬럼보이기설정
    Call SetColumnView(spdWork)

    'spdWork.ColWidth(spdWork.MaxCols) = 30

    If gEMR = "BRAIN" Then
        fraBrain.Visible = True
    Else
        fraBrain.Visible = False
    End If
    
'    If gEMR = "JWINFO" Then
'        fraJWINFO.Visible = True
'    Else
'        fraJWINFO.Visible = False
'    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub CtlInitializing()
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    txtSeqNo.Text = "1"
    
    '순번사용
    If gHOSP.RSTTYPE = "1" Then
        lblSeqNo.Visible = True
        txtSeqNo.Visible = True
    Else
        lblSeqNo.Visible = False
        txtSeqNo.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight = 0 Then Exit Sub

    With spdWork
        .Visible = False
        .WIDTH = Me.ScaleWidth - 300
        .HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
        .ColWidth(colITEMS) = 1000
        .ColWidth(colITEMS) = spdWork.MaxTextColWidth(colITEMS) * 1.1
        .Visible = True
    End With
    
End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer

    If Row = 0 And Col <> colCHECKBOX Then
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
    End If
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    If Col <> colBARCODE Then
        Exit Sub
    End If
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With frmInterface.spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            frmInterface.spdOrder.MaxRows = frmInterface.spdOrder.MaxRows + 1
            intRow = frmInterface.spdOrder.MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, i), intRow, i)

                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                        frmInterface.spdOrder.Row = 0
                        frmInterface.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmInterface.spdOrder.Text) Then
                            .Row = intRow
                            Call SetText(frmInterface.spdOrder, "◇", intRow, intOCol)
                        End If
                    Next
                Next
            Next
            
            frmInterface.spdOrder.RowHeight(-1) = 15
        End If
    
    End With
    
End Sub

Private Sub spdWork_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    
    
    sRow = spdWork.ActiveRow
    sCol = colPNAME
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdWork, sRow, sCol)
    
    If KeyCode = vbKeyDelete Then
        
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdWork, sRow, sRow
        spdWork.MaxRows = spdWork.MaxRows - 1
    
    End If

End Sub

Private Sub spdWork_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    Dim strSeq As String
    
    If KeyAscii = vbKeyReturn Then
        With spdWork
            If .ActiveCol = colSEQNO Then
                strSeq = GetText(spdWork, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdWork, strSeq + 1, intRow, colSEQNO)
                Next
            End If
        End With
    End If
    
End Sub


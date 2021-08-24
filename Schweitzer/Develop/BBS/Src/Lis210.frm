VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm210UnverifiedList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C9D6D6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "미확인결과 리스트"
   ClientHeight    =   8130
   ClientLeft      =   10935
   ClientTop       =   855
   ClientWidth     =   10155
   Icon            =   "Lis210.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdUpDown 
      BackColor       =   &H00E7BAB4&
      Caption         =   "▲"
      Height          =   345
      Left            =   9300
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   60
      Width           =   360
   End
   Begin VB.ComboBox cboWorkArea 
      Height          =   300
      Left            =   30
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   75
      Width           =   2100
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3225
      Top             =   8115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lis210.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lis210.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lis210.frx":094A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2805
      Top             =   8085
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FEDECD&
      Caption         =   "Re&fresh"
      Height          =   345
      Left            =   8520
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   60
      Width           =   765
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   344
      BackColor       =   13227734
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "From"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Index           =   0
      Left            =   4110
      TabIndex        =   4
      Top             =   135
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   344
      BackColor       =   13227734
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "To"
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpRcvDt 
      Height          =   300
      Left            =   4530
      TabIndex        =   5
      Top             =   75
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yy-MM-dd"
      Format          =   17039363
      CurrentDate     =   36467
   End
   Begin MSComCtl2.DTPicker dtpTm 
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      Top             =   75
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:mm:ss"
      Format          =   17039363
      CurrentDate     =   37770
   End
   Begin MSComCtl2.DTPicker dtpFRcvDt 
      Height          =   300
      Left            =   2835
      TabIndex        =   7
      Top             =   75
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yy-MM-dd"
      Format          =   17039363
      CurrentDate     =   36467
   End
   Begin FPSpread.vaSpread tblLabList 
      Height          =   7710
      Left            =   0
      TabIndex        =   8
      Top             =   420
      Width           =   10125
      _Version        =   196608
      _ExtentX        =   17859
      _ExtentY        =   13600
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   15
      MaxRows         =   30
      OperationMode   =   1
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis210.frx":0C6E
      TextTip         =   2
   End
End
Attribute VB_Name = "frm210UnverifiedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRst As New clsPatientInfo
Private blnForce As Boolean

'Private Sub cmdAll_Click()
'    If cmdAll.Tag = "1" Then   '전체데이타
'        DoneFg = "1"
'        cmdAll.Caption = "New"
'        cmdAll.Tag = "2'"
'    Else   '새로운 데이타
'        DoneFg = ""
'        cmdAll.Caption = "All"
'        cmdAll.Tag = "1"
'    End If
'    Call Get_Data
'End Sub

Private Sub cboWorkArea_Click()
   
'    tblLabList.Col = 1: tblLabList.COL2 = tblLabList.MaxCols
'    tblLabList.Row = 2: tblLabList.Row2 = tblLabList.MaxRows
'    tblLabList.BlockMode = True
'    tblLabList.Action = ActionClear
    tblLabList.MaxRows = 0
    
    tblLabList.BlockMode = False
    Call Get_Data

End Sub

Private Sub cmdRefresh_Click()
    Call Get_Data
End Sub

Private Sub cmdUpDown_Click()
    If cmdUpDown.tag = "0" Then
        cmdUpDown.tag = "1"
        cmdUpDown.Caption = "▼"
        Me.Height = 810
        blnForce = True
        
        Me.Width = 2685
        cmdUpDown.Left = 2205
    Else
        cmdUpDown.tag = "0"
        cmdUpDown.Caption = "▲"
        Me.Height = 8505
        blnForce = False
        
        Me.Width = 9765
        cmdUpDown.Left = 9300
    End If
End Sub

Private Sub dtpFRcvDt_Change()
    Me.Caption = "미확인결과 리스트 (" & Format(dtpFRcvDt.value, "MM.DD") & ")"
    Call Get_Data
End Sub

Private Sub dtpRcvDt_Change()
    Me.Caption = "미확인결과 리스트 (" & Format(dtpRcvDt.value, "MM.DD") & ")"
    Call Get_Data
End Sub

Private Sub dtpTm_Change()
    Call Get_Data
End Sub

Private Sub Form_Load()
    Dim strWA As String
    
    Me.Top = 600
    Me.Left = 5500
    Me.Show
    
    dtpFRcvDt.value = GetSystemDate
    dtpRcvDt.value = GetSystemDate
    dtpTm.value = GetSystemDate
    
    Call medAlwaysOn(frm210UnverifiedList, 1)
    Call objRst.Load_WorkArea(cboWorkArea)
'    cboWorkArea.ListIndex = 0
    
    '설정된 Workarea 가 있는경우 읽기
    
    strWA = GetSetting("Schweitzer2000 LIS", "Options", "UnvfyForWA", vbNullString)
    
    If strWA <> vbNullString Then
        cboWorkArea.ListIndex = Val(strWA)
    Else
        cboWorkArea.ListIndex = 0
    End If
    
    Me.Caption = "미확인결과 리스트 (" & Format(dtpRcvDt.value, "MM.DD") & ")"
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub


Private Sub Get_Data()

    Dim i As Integer
    Dim tmpWorkArea As String
    Dim tmpRcvDt As String
    
    MouseRunning
    DoEvents
    
    tmpWorkArea = medGetP(cboWorkArea.Text, 1, " ")


    tmpRcvDt = Format(dtpRcvDt.value, CS_DateDbFormat) & Format(dtpTm.value, CS_TimeDbFormat) & _
               Format(dtpFRcvDt.value, CS_DateDbFormat)
    
    Call objRst.LoadUnverifiedList(tmpWorkArea, tmpRcvDt, ObjSysInfo.BuildingCd, tblLabList)
    
    MouseDefault
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If blnForce Then Exit Sub
    If cmdUpDown.tag = "1" Then
        cmdUpDown.tag = "0"
        cmdUpDown.Caption = "▲"
        Me.Height = 8505
        
        Me.Width = 9765
        cmdUpDown.Left = 9300
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting("Schweitzer2000 LIS", "Options", "UnvfyForWA", cboWorkArea.ListIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm210UnverifiedList = Nothing
End Sub


Private Sub tblLabList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strMask As String
    Static iSortOrder As Integer
    
    With tblLabList
        
        Dim tmpLabNo As String
        
        If Row = 0 Then
            .Row = 0: .Col = Col
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Action = ActionSort
            Exit Sub
        End If
        
        .Row = Row
        .Col = 5
        If .value = "0" Then    '일반검사
            .Col = 1: tmpLabNo = .value
            frm202AccDataEntry.WindowState = 2
            frm202AccDataEntry.Show
'            frm202AccDataEntry.chkSpcNo.Value = 0
'            frm202AccDataEntry.chkSpcNo_Click
            DoEvents
            strMask = String(Len(medGetP(tmpLabNo, 1, "-")), "&") & "-"
            strMask = strMask & String(Len(medGetP(tmpLabNo, 2, "-")), "#") & "-"
            strMask = strMask & String(Len(medGetP(tmpLabNo, 3, "-")), "#")
            frm202AccDataEntry.ClearData
            frm202AccDataEntry.mskAccNo.Mask = strMask
            frm202AccDataEntry.mskAccNo.Text = tmpLabNo '& String(14 - Len(tmpLabNo), "_")
            DoEvents
            'Call frm202AccDataEntry.Data_Load
'            frm202AccDataEntry.lvwPatient.SetFocus
            SendKeys "{TAB}"
        ElseIf .value = "1" Then    '기타검사
            .Col = 1: tmpLabNo = .value
            frm293SpecialTest.WindowState = 2
            frm293SpecialTest.Show
            DoEvents
            frm293SpecialTest.optInput(0).value = True
            DoEvents
            frm293SpecialTest.txtWorkArea.Text = medGetP(tmpLabNo, 1, "-")
            frm293SpecialTest.txtAccDt.Text = medGetP(tmpLabNo, 2, "-")
            frm293SpecialTest.txtAccSeq.Text = medGetP(tmpLabNo, 3, "-")
            DoEvents
            frm293SpecialTest.Call_txtAccSeq_KeyPress
            DoEvents
        Else
            MsgBox "미생물 검사입니다. 미생물 결과등록을 이용하세요.", vbInformation, "메세지"
            Exit Sub
        End If
        
        If cmdUpDown.tag = "0" Then
            cmdUpDown.tag = "1"
            cmdUpDown.Caption = "▼"
            Me.Height = 810
            
            Me.Width = 2685
            cmdUpDown.Left = 2205
        End If
        
    End With
    
End Sub

Private Sub tblLabList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'    Dim strToolTip As String    '표시할 ToolTip
'    Dim strColNm   As String    '채혈자
'
'    If Row < 1 Then Exit Sub
'
'    strToolTip = vbCrLf
'    With tblLabList
'        .Row = Row
'        .Col = 12
'            strColNm = GetDoctNm(.value)
'            If Trim$(strColNm) = "" Then strColNm = GetEmpNm(.value)
'            strToolTip = strToolTip & "  채 혈 자 : " & strColNm & vbCrLf
'        .Col = 11: strToolTip = strToolTip & "  채혈일시 : " & .value & vbCrLf
'        .Col = 14: strToolTip = strToolTip & "  접 수 자 : " & GetEmpNm(.value) & vbCrLf
'        .Col = 13: strToolTip = strToolTip & "  접수일시 : " & .value & vbCrLf
'
'        MultiLine = 1
'        TipText = strToolTip
'        TipWidth = 4000
'        .TextTipDelay = 1000
'        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
'        ShowTip = True
'    End With
    
    Dim strToolTip As String    '표시할 ToolTip
    Dim strColNm   As String    '채혈자
    
    If Row < 1 Then Exit Sub
    
    strToolTip = vbCrLf
    With tblLabList
        .Row = Row
        .Col = 13
            strColNm = GetDoctNm(.value)
            If Trim$(strColNm) = "" Then strColNm = GetEmpNm(.value)
            strToolTip = strToolTip & "  채 혈 자 : " & strColNm & vbCrLf
        .Col = 12: strToolTip = strToolTip & "  채혈일시 : " & .value & vbCrLf
        .Col = 15: strToolTip = strToolTip & "  접 수 자 : " & GetEmpNm(.value) & vbCrLf
        .Col = 14: strToolTip = strToolTip & "  접수일시 : " & .value & vbCrLf
        
        MultiLine = 1
        TipText = strToolTip
        TipWidth = 4000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub Timer1_Timer()
    Dim strMin  As String
    Dim strSec  As String
    
    dtpTm.value = Now
    
    Static TimeCount As Long
    Static ImgCount As Integer
    
    ImgCount = ImgCount + 1
    TimeCount = TimeCount + 1
    Me.Icon = ImgList.ListImages(ImgCount).Picture
    If ImgCount = 3 Then ImgCount = 0
    
    strMin = Mid((300 - TimeCount) / 60, 1, 1)
    strSec = ((300 - TimeCount) Mod 60)
    
    Me.Caption = "미확인결과 리스트 (" & Format(dtpRcvDt.value, "MM.DD") & ")" & " " & "남은시간 : " & strMin & " 분 " & strSec & " 초"
    
    If TimeCount = 300 Then
        Call Get_Data: TimeCount = 0 '5분 간격
    End If
    
End Sub

VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm158AccPtList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "접수대기자 명단"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox chkPay 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전체처방조회(미수납포함)"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   3285
   End
   Begin MedControls1.LisLabel lblSortOrder 
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   285
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   556
      BackColor       =   15132390
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "환자 성명 순"
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FEDECD&
      Caption         =   "Re&fresh"
      Height          =   330
      Left            =   2595
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   270
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1425
      Top             =   3750
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1980
      Top             =   2895
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
            Picture         =   "Lis158.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lis158.frx":0324
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lis158.frx":0640
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread tblLabList 
      Height          =   7680
      Left            =   30
      TabIndex        =   1
      Top             =   915
      Width           =   3360
      _Version        =   196608
      _ExtentX        =   5927
      _ExtentY        =   13547
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
      MaxCols         =   4
      MaxRows         =   30
      OperationMode   =   1
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis158.frx":0964
   End
   Begin MSComCtl2.DTPicker dtpFromDt 
      Height          =   285
      Left            =   510
      TabIndex        =   3
      Top             =   630
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   67305473
      CurrentDate     =   37544
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   285
      Left            =   2085
      TabIndex        =   4
      Top             =   630
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   67305473
      CurrentDate     =   37544
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "From : "
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "To : "
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1785
      TabIndex        =   5
      Top             =   675
      Width           =   225
   End
End
Attribute VB_Name = "frm158AccPtList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSQL As New clsLISSqlStatement

Private Sub chkPay_Click()
    Call Get_Data
End Sub

Private Sub cmdRefresh_Click()
    Call Get_Data
    lblSortOrder.Caption = "처방일 순 (내림차순)"
End Sub


Private Sub Form_Load()
    
    Me.Top = 1665
    Me.Left = 0
    Me.Show
    DoEvents
    dtpFromDt.Value = GetSystemDate
    dtpToDt.Value = GetSystemDate
    
    Call medAlwaysOn(frm158AccPtList, 1)
    Call Get_Data
    Timer1.Interval = 1000
    Timer1.Enabled = True
    chkPay.Value = 0
End Sub

Private Sub Get_Data()

    Dim i As Integer
    Dim SqlStmt As String
    Dim tmpRs   As Recordset
    Dim strKey  As String
    Dim j As Integer
    
    Me.Caption = "접수대기자 명단 (데이타 로드중..)"
    MouseRunning
    

    Dim FrDt As String
    Dim ToDt As String
    
    FrDt = Format(dtpFromDt.Value, "yyyymmdd")
    ToDt = Format(dtpToDt.Value, "yyyymmdd")
    objSQL.PayDt = ""
    
    
    If chkPay.Value = 0 Then
        objSQL.PayDt = "PayComplete"
    End If
    
    SqlStmt = objSQL.SqlGetPtForAccess(FrDt, ToDt)
    
    Set tmpRs = New Recordset
    
    tmpRs.Open SqlStmt, DBConn
    
    With tblLabList
        .ReDraw = False
        
        '-- 원본 ==============================================================================
'        .MaxRows = 0
'        .MaxRows = tmpRs.RecordCount
'        For i = 1 To tmpRs.RecordCount
'            .Row = i
'            .Col = 1: .Value = GetPtNm(tmpRs.Fields("PtId").Value)
'            .Col = 2: .Value = Trim(tmpRs.Fields("PtId").Value)
'            .Col = 3: .Value = Format(Mid(tmpRs.Fields("OrdDt").Value, 3, 6), CS_DateShortMask)
'            .Col = 4: .Value = tmpRs.Fields("OrdDt").Value
'            tmpRs.MoveNext
'        Next
        '======================================================================================
        
        '-- 전주예수병원 ======================================================================
        ' - By M.G.Choi  2004.10.25
        .MaxRows = 0
        j = 1
        For i = 1 To tmpRs.RecordCount
            If strKey <> Trim(tmpRs.Fields("PtId").Value) & tmpRs.Fields("OrdDt").Value Then
                .MaxRows = j
                .Row = j
                .Col = 1: .Value = GetPtNm(tmpRs.Fields("PtId").Value)
                .Col = 2: .Value = Trim(tmpRs.Fields("PtId").Value)
                .Col = 3: .Value = Format(Mid(tmpRs.Fields("OrdDt").Value, 3, 6), CS_DateShortMask)
                .Col = 4: .Value = tmpRs.Fields("OrdDt").Value
                
                strKey = Trim(tmpRs.Fields("PtId").Value) & tmpRs.Fields("OrdDt").Value
                
                j = j + 1
                
            End If
                
            tmpRs.MoveNext
        Next
        '======================================================================================
        
        .ReDraw = True
    End With
          
    Set tmpRs = Nothing
    lblSortOrder.Caption = "처방일 순 (내림차순)"
    Me.Caption = "접수대기자 명단"
    MouseDefault
    
End Sub

'Private Function GetPtNm(ByVal vPtId As String) As String
'    Dim objData As New clsBasisData
'
'    GetPtNm = GetPtNm(vPtId)
'    Set objData = Nothing
'End Function

Private Sub tblLabList_Click(ByVal Col As Long, ByVal Row As Long)

    Static iSortOrder As Integer
    Dim tmpColNm As String
    Dim strDt As String
    
    With tblLabList
        
        If Row = 0 Then  'Sort...
            .Row = 0: .Col = Col: tmpColNm = .Text
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            If Col = 3 Then
                .SortKey(1) = 4
            Else
                .SortKey(1) = Col
            End If
            lblSortOrder.Caption = tmpColNm & " 순 ("
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
                lblSortOrder.Caption = lblSortOrder.Caption & "내림차순)"
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
                lblSortOrder.Caption = lblSortOrder.Caption & "오름차순)"
            End If
            .Action = ActionSort
            Exit Sub
        End If
        
        
        .Row = Row
        
        Dim PtId As String
        Dim ordDt As String

        .Col = 2: PtId = .Value
        .Col = 3: ordDt = .Value

        ShowCollectionForm "frm165OutCol", PtId, ordDt

        Exit Sub
'모야.. frm153SendPt 폼이 없잖아..
'        frm153SendPt.WindowState = 2
'        frm153SendPt.Show
'        DoEvents
'        frm153SendPt.Call_cmdClear_Click
'        frm153SendPt.txtPtid.SetFocus
'        '스프레드 데이타 선택시 외래접수화면으로 이동
'        .Col = 2: frm153SendPt.txtPtid.Text = .Value
'        frm153SendPt.cboOrdDate.SetFocus
'        DoEvents
'        'frm153SendPt.cboOrdDate.Clear
'        .Col = 3: strDt = .Value
'        frm153SendPt.cboOrdDate.ListIndex = medComboFind(frm153SendPt.cboOrdDate, Format(strDt, CS_DateLongFormat))
'        DoEvents
'        'Call frm202AccDataEntry.Data_Load
'        'SendKeys "{TAB}"
        
    End With
    
End Sub

Private Sub Timer1_Timer()
    
    Static TimeCount As Long
    Static ImgCount As Integer
    
    ImgCount = ImgCount + 1
    TimeCount = TimeCount + 1
    Me.Icon = ImgList.ListImages(ImgCount).Picture
    If ImgCount = 3 Then ImgCount = 0
    If TimeCount = 300 Then Call Get_Data: TimeCount = 0 '5분 간격
    
End Sub

Private Sub ShowCollectionForm(ByVal pFrmName As String, ByVal PtId As String, ByVal ordDt As String)

    Dim i As Integer
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    frmLisCollection.ButtonKey = "LIS218"
    frmLisCollection.Show
    frmLisCollection.ZOrder 0
    frmLisCollection.ShowThisForm
    'lblSubMenu.Caption = "외래환자LoadOutCollection 접수"
    frmLisCollection.LoadOutCollection PtId, ordDt
'    blnFormShow = True
    
    Exit Sub
    
PermissionDenied:

'    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub


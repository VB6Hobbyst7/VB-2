VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCounsel_4 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmCounsel_4.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin ChartfxLibCtl.ChartFX chtNutrition 
      Height          =   2505
      Left            =   1170
      TabIndex        =   5
      Top             =   5650
      Width           =   7125
      _cx             =   12568
      _cy             =   4419
      Build           =   20
      TypeMask        =   111673346
      LeftGap         =   39
      RightGap        =   9
      BottomGap       =   39
      Volume          =   50
      AxesStyle       =   2
      Axis(0).Scale   =   10
      Axis(0).Max     =   40
      Axis(0).Decimals=   0
      Axis(0).Style   =   14440
      Axis(0).TickMark=   -32767
      Axis(0).Format  =   4
      Axis(0).Format  =   4
      Axis(2).Style   =   14380
      Axis(2).TickMark=   0
      RGBBk           =   16777216
      nColors         =   16
      Colors          =   "frmCounsel_4.frx":904B
      BottomFontMask  =   268435464
      BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      nPts            =   14
      nSer            =   1
      NumPoint        =   14
      NumSer          =   1
      BorderS         =   0
      _Data_          =   "frmCounsel_4.frx":90EB
   End
   Begin VB.ComboBox cmbTopFood 
      Height          =   300
      Left            =   8730
      Style           =   2  '드롭다운 목록
      TabIndex        =   11
      Top             =   7980
      Width           =   1335
   End
   Begin ChartfxLibCtl.ChartFX Chart 
      Height          =   3105
      Left            =   720
      TabIndex        =   10
      Top             =   5160
      Width           =   7905
      _cx             =   13944
      _cy             =   5477
      Build           =   20
      TypeMask        =   111673345
      MarkerShape     =   0
      MarkerSize      =   5
      Axis(0).Max     =   90
      Axis(0).TickMark=   -32767
      RGB2DBk         =   16777215
      nColors         =   16
      Colors          =   "frmCounsel_4.frx":91F8
      nSer            =   1
      NumSer          =   1
      _Data_          =   "frmCounsel_4.frx":9298
   End
   Begin VB.ComboBox cmbDaily 
      Height          =   300
      Left            =   10470
      Style           =   2  '드롭다운 목록
      TabIndex        =   9
      Top             =   2190
      Width           =   1515
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      Left            =   10500
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   1770
      Width           =   1455
   End
   Begin ChartfxLibCtl.ChartFX chtCaloryRate 
      Height          =   3165
      Left            =   6690
      TabIndex        =   1
      Top             =   1680
      Width           =   3465
      _cx             =   6112
      _cy             =   5583
      Build           =   20
      TypeMask        =   111673349
      LeftGap         =   0
      RightGap        =   0
      TopGap          =   9
      BottomGap       =   0
      View3DDepth     =   90
      Volume          =   70
      nColors         =   16
      Pallete         =   "frmCounsel_4.frx":93A5
      Colors          =   "frmCounsel_4.frx":9489
      nPts            =   4
      nSer            =   1
      NumPoint        =   4
      NumSer          =   1
      BorderS         =   0
      ExtCmd          =   30209
      _Data_          =   "frmCounsel_4.frx":9529
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAniVegRate 
      Height          =   3105
      Left            =   720
      TabIndex        =   2
      Top             =   5160
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5477
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd19 
      Height          =   3105
      Left            =   720
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5477
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMealRank 
      Height          =   3105
      Left            =   720
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5477
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   285
      Left            =   10590
      TabIndex        =   6
      Top             =   2100
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   37818
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   10590
      TabIndex        =   7
      Top             =   2370
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   37818
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMealCalory 
      Height          =   3105
      Left            =   720
      TabIndex        =   0
      Top             =   5160
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5477
      _Version        =   393216
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   8400
      Picture         =   "frmCounsel_4.frx":966C
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "(%)"
      Height          =   225
      Left            =   1020
      TabIndex        =   27
      Top             =   5220
      Width           =   285
   End
   Begin VB.Image imgStart 
      Height          =   300
      Left            =   9390
      Picture         =   "frmCounsel_4.frx":9D92
      Top             =   1350
      Width           =   765
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_4.frx":A6B9
      Top             =   50
      Width           =   13140
   End
   Begin VB.Label lblTopFood 
      BackStyle       =   0  '투명
      Caption         =   "피자 > 스파게티 > 떡볶이 > 쌀밥 > 배추김치"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Index           =   4
      Left            =   4110
      TabIndex        =   26
      Top             =   4290
      Width           =   2265
   End
   Begin VB.Label lblTopFood 
      BackStyle       =   0  '투명
      Caption         =   "피자 > 스파게티 > 떡볶이 > 쌀밥 > 배추김치"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Index           =   3
      Left            =   4110
      TabIndex        =   25
      Top             =   3810
      Width           =   2265
   End
   Begin VB.Label lblTopFood 
      BackStyle       =   0  '투명
      Caption         =   "피자 > 스파게티 > 떡볶이 > 쌀밥 > 배추김치"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Index           =   2
      Left            =   4110
      TabIndex        =   24
      Top             =   3360
      Width           =   2265
   End
   Begin VB.Label lblTopFood 
      BackStyle       =   0  '투명
      Caption         =   "피자 > 스파게티 > 떡볶이 > 쌀밥 > 배추김치"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Index           =   1
      Left            =   4110
      TabIndex        =   23
      Top             =   2880
      Width           =   2265
   End
   Begin VB.Label lblTopFood 
      BackStyle       =   0  '투명
      Caption         =   "피자 > 스파게티 > 떡볶이 > 쌀밥 > 배추김치"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Index           =   0
      Left            =   4110
      TabIndex        =   22
      Top             =   2430
      Width           =   2265
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3120
      TabIndex        =   21
      Top             =   4290
      Width           =   855
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   3090
      TabIndex        =   20
      Top             =   3900
      Width           =   915
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3090
      TabIndex        =   19
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,200"
      Height          =   225
      Index           =   1
      Left            =   3090
      TabIndex        =   18
      Top             =   2970
      Width           =   915
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,200"
      Height          =   225
      Index           =   0
      Left            =   3090
      TabIndex        =   17
      Top             =   2490
      Width           =   915
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "4,801"
      Height          =   195
      Index           =   4
      Left            =   2190
      TabIndex        =   16
      Top             =   4350
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "5,230"
      Height          =   195
      Index           =   3
      Left            =   2190
      TabIndex        =   15
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "10.9g(24%)"
      Height          =   195
      Index           =   2
      Left            =   2190
      TabIndex        =   14
      Top             =   3450
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65:20:15"
      Height          =   195
      Index           =   1
      Left            =   2190
      TabIndex        =   13
      Top             =   2970
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,780"
      Height          =   195
      Index           =   0
      Left            =   2190
      TabIndex        =   12
      Top             =   2490
      Width           =   885
   End
   Begin VB.Image imgPrint 
      Height          =   1050
      Left            =   10650
      Picture         =   "frmCounsel_4.frx":C1BB
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   9
      Left            =   10500
      Top             =   6750
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   8
      Left            =   10500
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   7
      Left            =   10500
      Top             =   5970
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   6
      Left            =   10500
      Top             =   5550
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   5
      Left            =   10500
      Top             =   5160
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   4
      Left            =   10500
      Top             =   4770
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   3
      Left            =   10500
      Top             =   4380
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   2
      Left            =   10500
      Top             =   3990
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   1
      Left            =   10500
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   0
      Left            =   10500
      Top             =   3180
      Width           =   1485
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   5
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":D89A
      Top             =   7530
      Width           =   1335
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   4
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":E411
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   3
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":F062
      Top             =   6630
      Width           =   1335
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   2
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":FC86
      Top             =   6210
      Width           =   1335
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   1
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":10929
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Image imgSub 
      Height          =   390
      Index           =   0
      Left            =   8730
      Picture         =   "frmCounsel_4.frx":115B6
      Top             =   5310
      Width           =   1335
   End
End
Attribute VB_Name = "frmCounsel_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 2004-02-10
Option Explicit
Private Type mCustomerInfo
    intState As Integer
    intAge As Integer
    strSex As String
    sngDietCal As Single
End Type
Private typCustomerInfo As mCustomerInfo

Private Const PATH04 As String = "\Back\Counsel\04\"
Private Const IMG_SUB1_ON As String = "1 on.jpg"
Private Const IMG_SUB1_OFF As String = "1 off.jpg"
Private Const IMG_SUB2_ON As String = "2 on.jpg"
Private Const IMG_SUB2_OFF As String = "2 off.jpg"
Private Const IMG_SUB3_ON As String = "3 on.jpg"
Private Const IMG_SUB3_OFF As String = "3 off.jpg"
Private Const IMG_SUB4_ON As String = "4 on.jpg"
Private Const IMG_SUB4_OFF As String = "4 off.jpg"
Private Const IMG_SUB5_ON As String = "5 on.jpg"
Private Const IMG_SUB5_OFF As String = "5 off.jpg"
Private Const IMG_SUB6_ON As String = "6 on.jpg"
Private Const IMG_SUB6_OFF As String = "6 off.jpg"

Private Const IMG_PRINT_ON As String = "상담-영양평가 출력 on.jpg"
Private Const IMG_PRINT_OFF As String = "상담-영양평가 출력 off.jpg"

Dim crxApplication As New CRAXDRT.Application
Public crxReport As CRAXDRT.Report
Public crxReport2 As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxFormula As CRAXDRT.FormulaFieldDefinition
Dim strServer As String, strDBName As String, strUID As String, strPWD As String

Public Sub Form_Load()
    Dim i As Integer
'폼의 뜨는 위치 및 그래픽 ==================================================
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\04\04.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite

    If ExistDiary = False Then
        MsgBox "입력한 식사일기가 없습니다. ", vbOKOnly + vbExclamation
        For i = 0 To 4
            lblIntake(i).Caption = ""
            lblRecommend(i).Caption = ""
            lblTopFood(i).Caption = ""
        Next i
        grdAniVegRate.Visible = False
        grd19.Visible = False
        grdMealCalory.Visible = False
        grdMealRank.Visible = False
        cmbTopFood.Visible = False
        Chart.Visible = False
        chtNutrition.Visible = False
        cmbDaily.Visible = False
        dtpBegin.Visible = False
        dtpEnd.Visible = False
        cmbPeriod.Enabled = False
        For i = 0 To 5
            imgSub(i).Enabled = False
        Next i
        chtCaloryRate.Visible = False
        Exit Sub
    Else
        chtCaloryRate.Visible = True
    End If
    
    Set imgPrint.Picture = LoadPicture(App.Path & PATH04 & IMG_PRINT_OFF)
'입력한 일기 중에 가장 최종일 영양평가를 보여줌
'만약 현재 선택된 일기가 있다면 그것을 보여줌- 현재 몇일치의 영양평가를 보여주고 있는지 보여줄 것
    cmbPeriod.Enabled = True
    dtpBegin.Value = Now()
    dtpEnd.Value = Now()
    
    Call InitialChartRate
    Call InitialControl
    
    For i = 0 To 5
        imgSub(i).Enabled = True
    Next i
    Call imgSub_Click(0)
    Call ShowNutritionInfo
End Sub

Private Function ExistDiary() As Boolean
'해당환자에 입력한 식사일기가 있는지 체크
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT DietDiaryNum FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        ExistDiary = True
    Else
        ExistDiary = False
    End If
End Function

Private Sub InitialControl()
    Dim i As Integer
    
    cmbPeriod.Clear
    cmbPeriod.AddItem "특정일"
    cmbPeriod.AddItem "특정기간"
    cmbPeriod.AddItem "전체"
    cmbPeriod.ListIndex = 0
    
    For i = 0 To 4
        lblIntake(i).Caption = ""
        lblRecommend(i).Caption = ""
        lblTopFood(i).Caption = ""
    Next i

'칼로리,단백질,비타민A,비타민E,비타민C,비타민B1,비타민B2,나이아신,비타민B6,엽산,칼슘,인,철,아연
    With cmbTopFood
        .Clear
        .AddItem "칼로리"
        .AddItem "탄수화물"
        .AddItem "단백질"
        .AddItem "지방"
        .AddItem "비타민A"
        .AddItem "비타민E"
        .AddItem "비타민C"
        .AddItem "비타민B1"
        .AddItem "비타민B2"
        .AddItem "나이아신"
        .AddItem "비타민B6"
        .AddItem "엽산"
        .AddItem "칼슘"
        .AddItem "나트륨"
        .AddItem "인"
        .AddItem "철"
        .AddItem "아연"
        .ListIndex = 0
    End With
End Sub

Private Sub cmbPeriod_Change()
    Call cmbPeriod_Click
End Sub

Private Sub cmbPeriod_Click()
    Select Case cmbPeriod.ListIndex
        Case 0:   '특정일
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            cmbDaily.Visible = True
            Call InitialDailyCombo
            Call imgStart_Click
       Case 1:   '특정기간
            cmbDaily.Visible = False
            dtpBegin.Visible = True
            dtpEnd.Visible = True
        Case 2:   '전체기간
            cmbDaily.Visible = False
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            Call imgStart_Click
    End Select
End Sub

Private Sub cmbDaily_Change()
    'Call cmbDaily_Click
End Sub

'<< 식사일기 평가 >> 페이지를 출력하기 위해 준비하는 함수 /////////////////////////////////////////
Private Sub PrintData()
    Dim strConString As String
    Dim qrySelect As String, rValue As Variant
    Dim strBeginDay As String, strEndDay As String
    Dim i As Integer

On Error GoTo PrintErr
    '출력전에 선택된 기간내에 출력할 내용이 있는지를 먼저 확인할 것
    Set clsSelect = New clsSelect

    strBeginDay = Format(dtpBegin.Value, "YYYYMMDD")
    strEndDay = Format(dtpEnd.Value, "YYYYMMDD")

    qrySelect = "SELECT DISTINCT MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum

    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & strBeginDay
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & strEndDay & "';"
    End If

    rValue = clsSelect.Query(qrySelect)
    If IsNull(rValue) Then
        MsgBox "기간내에 입력된 식사일기가 없습니다.", vbInformation
        Set clsSelect = Nothing
        Exit Sub
    End If
    Set clsSelect = Nothing
    '리포터 연결 설정
    strServer = ServerName
'2005-01-18 류진선 DB정보수정
    strDBName = DBinfo.DBName
    strUID = DBinfo.DBID
    strPWD = DBinfo.DBPWD
'    strDBName = "Body"
'    strUID = "sa"
'    strPWD = "1111"


    Set crxReport = crxApplication.OpenReport(App.Path & "\Report\식사일기평가.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport
        .RecordSelectionFormula = "{CustomerInfo.CustomerNum}=" & glngCustomerNum

        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
'//////////////////////////////////////////  RDC 방식변경
        '+--------------------------------------------------
        '+ 1) 영양소 섭취상태
        '+--------------------------------------------------
        Call LoadCustomerInfo(glngCustomerNum)
        If typCustomerInfo.sngDietCal <> 0 Then
            With typCustomerInfo
                If Calculate_Nut(.sngDietCal, .intState, .intAge, .strSex) = True Then
                End If
            End With
        End If

        '1 : @Sex
        '2 : @FatPercent
        '3 : @Top5_Calory
        '4 : @Top5_Fat
        '5 : @Top5_SFA
        '6 : @Top5_Chol
        '7 : @Top5_Na
        '8 : @TreatCalory
        '9 : @Period
        '    - 섭취량, 가장 많이 포함된 다섯가지 음식
        '    - 열량 / 총지방량 / 포화지방 / 포화,불포화 / 콜레스테롤 / 나트륨
        .FormulaFields(3).Text = "'" & RPT_TopFood("열량") & "'"
        .FormulaFields(4).Text = "'" & RPT_TopFood("지방") & "'"
        .FormulaFields(5).Text = "'" & RPT_TopFood("포화지방") & "'"
        .FormulaFields(6).Text = "'" & RPT_TopFood("콜레스테롤") & "'"
        .FormulaFields(7).Text = "'" & RPT_TopFood("나트륨") & "'"
'        '    - 열량 권장량(선택된 기간내 처방된 칼로리들의 평균값)
        .FormulaFields(8).Text = "'" & Format(WhatTreatCalory, "#,###") & "'"
        '    - 선택된 기간 뿌려줌
        If cmbPeriod.ListIndex = 0 Then   '특정일
            .FormulaFields(9).Text = "'" & Format(cmbDaily.Text, "YYYY.M.D") & "'"
        ElseIf cmbPeriod.ListIndex = 1 Then
            If dtpBegin.Value = dtpEnd.Value Then   '특정기간 선택시
                .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YYYY.M.D") & "'"
            Else
                .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YY.M.D") & " ~ " & Format(dtpEnd.Value, "YY.M.D") & "'"
            End If
        ElseIf cmbPeriod.ListIndex = 2 Then         '전체 선택시
            '초기 방문일부터 ~ ?
            Set clsSelect = New clsSelect

            qrySelect = "SELECT MIN(MealDate) FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum

            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                .FormulaFields(9).Text = "'" & Format(rValue(0, 0), "YYYY.M.D") & " ~'"
            End If
        End If

        '+--------------------------------------------------
        '+ 2) 식습관
        '+--------------------------------------------------
        '    [1] 평균 일별 식사 횟수
        '10 : @Count
        Set clsSelect = New clsSelect

        qrySelect = "SELECT AVG(a) FROM ("
        qrySelect = qrySelect & "SELECT MealDate, COUNT(DietDiaryNum) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If
        qrySelect = qrySelect & " AND MealCalory IS NOT NULL"
        qrySelect = qrySelect & " GROUP BY MealDate) b;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            .FormulaFields(10).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(10).Text = "'0'"
        End If
        '    [2] 식사장소 / 시간
        '11 : @장소아침
        '12 : @장소점심
        '13 : @장소저녁
        '14 : @Time_B
        '15 : @Time_L
        '16 : @Time_D
        qrySelect = "SELECT MealSection, AVG(MealPlace), AVG(CAST(MealTime AS INT)) "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If
        qrySelect = qrySelect & " GROUP BY MealSection;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                Select Case rValue(0, i)
                    Case 1    ' 아침
                        .FormulaFields(11).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 2    ' 점심
                        .FormulaFields(12).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 3    ' 저녁
                        .FormulaFields(13).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                End Select
            Next i
        End If
        '    [3] 걸린시간
        '17 : @NeedTime
        qrySelect = "SELECT AVG(MealNeedTime) FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & "AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(17).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(17).Text = "'0'"
        End If
        '    [4] 기분
        '18 : @Feeling
        qrySelect = "SELECT TOP 1 MealFeeling, COUNT(MealFeeling) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If
        qrySelect = qrySelect & " GROUP BY MealFeeling"
        qrySelect = qrySelect & " ORDER BY a DESC"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(18).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(18).Text = "'0'"
        End If
        '    [5] 식 후 배고픔 정도
        '19 : @Hungry
        qrySelect = "SELECT TOP 1 AfterMealHungry, COUNT(AfterMealHungry) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If
        qrySelect = qrySelect & " GROUP BY AfterMealHungry"
        qrySelect = qrySelect & " ORDER BY a DESC"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(19).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(19).Text = "'0'"
        End If
        '    [6] 외식횟수
        '20 : @EatOut
        qrySelect = "SELECT AVG(a) FROM ("
        qrySelect = qrySelect & "SELECT MealDate, COUNT(DietDiaryNum) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & "AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
        End If
        qrySelect = qrySelect & "AND MealPlace=3 GROUP BY MealDate) b"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(20).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(20).Text = "'0'"
        End If

        .PrintOut
    End With
    '+--------------------------------------------------
    '+ 두번째 장
    '+--------------------------------------------------
    Dim strTemp As String, strBeginDay1 As String, strEndDay1 As String
    Set crxReport2 = crxApplication.OpenReport(App.Path & "\Report\식사일기평가2.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport2
        '1 : @GabCalory
        '2 : @GabMent
        '3 : @Rice
        '4 : @Exercise
        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
        '    [1] 출력할 기간 선택
        If cmbPeriod.ListIndex = 1 Then
            strBeginDay1 = Left(strBeginDay, 4) & "," & Mid(strBeginDay, 5, 2) & "," & Right(strBeginDay, 2)
            strEndDay1 = Left(strEndDay, 4) & "," & Mid(strEndDay, 5, 2) & "," & Right(strEndDay, 2)
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum & " AND {DietDiary.MealDate} IN DateTime (" & strBeginDay1 & ") TO DateTime (" & strEndDay1 & ")"
        Else
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum
        End If
        .RecordSelectionFormula = strTemp

        '    [2] 하단에 요약정보
        Dim sngAvgTreatCal As Single, sngAvgMealCal As Single
        Dim sngAvgWeight As Single
        Dim sngGabCal As Single, sngRice As Single, intExercise As Integer
        '        - 해당기간내에 처방된 칼로리(Treat.TreatCalory)의 평균값
        sngAvgTreatCal = WhatTreatCalory
        If sngAvgTreatCal <> 0 Then
        '        - 해당기간내에 먹은 음식(일별)(DietDiary)의 평균값
            qrySelect = "SELECT AVG(a) FROM ("
            qrySelect = qrySelect & "SELECT MealDate, SUM(MealCalory) AS a "
            qrySelect = qrySelect & "FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            If cmbPeriod.ListIndex = 0 Then
                qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
            ElseIf cmbPeriod.ListIndex = 1 Then
                qrySelect = qrySelect & " AND MealDate BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
            End If
            qrySelect = qrySelect & " GROUP BY MealDate) b"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue(0, 0)) Then
                sngAvgMealCal = CSng(rValue(0, 0))
                sngGabCal = sngAvgTreatCal - sngAvgMealCal
                .FormulaFields(1).Text = "'" & Format(Abs(sngGabCal), "#,###") & "'"
                If sngGabCal > 0 Then
                    .FormulaFields(2).Text = "'부족'"
                ElseIf sngGabCal < 0 Then
                    .FormulaFields(2).Text = "'초과'"
                Else
                    .FormulaFields(2).Text = "'권장량'"
                End If
            '    - 밥 한공기 300kcal
                sngRice = Abs(sngGabCal) / 300
                If sngRice >= 0.6 Then
                    .FormulaFields(3).Text = "'" & Format(sngRice, "#") & "'"
                ElseIf sngRice < 0.6 And sngRice >= 0.4 Then
                    .FormulaFields(3).Text = "'반'"
                Else
                    .FormulaFields(3).Text = "'0'"
                End If
                qrySelect = "SELECT AVG(a) FROM ("
                qrySelect = qrySelect & "SELECT TreatDay, SUM(Weight) AS a "
                qrySelect = qrySelect & "FROM BodyData LEFT JOIN Treat "
                qrySelect = qrySelect & "ON BodyData.TreatNum=Treat.TreatNum "
                qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
                If cmbPeriod.ListIndex = 1 Then
                    qrySelect = qrySelect & " AND TreatDay BETWEEN '" & strBeginDay & "' AND '" & strEndDay & "' "
                End If
                qrySelect = qrySelect & " GROUP BY TreatDay) b"

                rValue = clsSelect.Query(qrySelect)
                If Not IsNull(rValue(0, 0)) Then
                    sngAvgWeight = CSng(rValue(0, 0))
                    intExercise = sngGabCal / (sngAvgWeight * 0.16)
                    .FormulaFields(4).Text = "'" & intExercise & " 분'"
                Else  '기간내 입력된 체중이 없다면 가장 최근 체중
                    qrySelect = "SELECT TOP 1 Weight FROM BodyData LEFT JOIN Treat "
                    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
                    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
                    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

                    rValue = clsSelect.Query(qrySelect)
                    If Not IsNull(rValue(0, 0)) Then
                        sngAvgWeight = CSng(rValue(0, 0))
                        intExercise = sngGabCal / (sngAvgWeight * 0.16)
                        .FormulaFields(4).Text = "'" & intExercise & " 분'"
                    End If
                End If
            Else
                .FormulaFields(1).Text = "'0'"
                .FormulaFields(2).Text = "''"
                .FormulaFields(3).Text = "''"
                .FormulaFields(4).Text = "''"
            End If
        Else
            .FormulaFields(1).Text = "'0'"
            .FormulaFields(2).Text = "''"
            .FormulaFields(3).Text = "''"
            .FormulaFields(4).Text = "''"
        End If

       .PrintOut
    End With

    MsgBox "출력이 완료되었습니다.", vbOKOnly + vbInformation, "출력"
    Set clsSelect = Nothing

    Exit Sub
PrintErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "PrintData", "frmCounsel_4", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Function WhatTreatCalory() As String
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    qrySelect = "SELECT AVG(TreatCalory) FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND TreatDay BETWEEN '" & Format(dtpBegin.Value, "YYYYMMDD") & "' "
        qrySelect = qrySelect & "AND '" & Format(dtpEnd.Value, "YYYYMMDD") & "' "
    End If

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue(0, 0)) Then
        WhatTreatCalory = rValue(0, 0)
    Else
        '만약 기간내에 입력된 처방칼로리가 없다면 가장 최근 것으로 사용한다.
        qrySelect = "SELECT TOP 1 TreatCalory FROM Treat "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            WhatTreatCalory = rValue(0, 0)
        Else
            WhatTreatCalory = "0"   '이 단계까지 오면 안됨
        End If
    End If

    Set clsSelect = Nothing
End Function

Private Function RPT_TopFood(strNutrition) As String
    Dim qrySelect As String, rValue As Variant
    Dim intMealCal As Single, strFldNutrition As String
    Dim strTopFood As String
    Dim i As Integer

    Set clsSelect = New clsSelect
    Select Case strNutrition
        Case "열량"
            strFldNutrition = "tblFood.Energy"
        Case "지방"
            strFldNutrition = "tblFood.Fat"
        Case "포화지방"
            strFldNutrition = "tblFood.SFA"
        Case "콜레스테롤"
            strFldNutrition = "tblFood.Cholesterol"
        Case "나트륨"
            strFldNutrition = "tblFood.Na"
    End Select

    qrySelect = "SELECT DISTINCT(MealName),"
    qrySelect = qrySelect & "SUM((DietFood.FoodWeight*" & strFldNutrition & ")/100) AS a "
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "INNER JOIN tblMeal ON DietMeal.MealCode=tblMeal.MealID "
    qrySelect = qrySelect & "INNER JOIN DietFood ON DietMeal.DietMealNum=DietFood.DietMealNum "
    qrySelect = qrySelect & "INNER JOIN tblFood ON DietFood.FoodCode=tblFood.FoodID "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYYMMDD") & "' AND '" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "GROUP BY MealDate, MealSection, MealName "
    qrySelect = qrySelect & "ORDER BY a DESC;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        strTopFood = ""
        If UBound(rValue, 2) < 4 Then
            For i = 0 To UBound(rValue, 2) - 1
                '음식명이 7자이상이면 ...으로 줄임표현
                If Not IsNull(rValue(0, i)) Then
                    strTopFood = strTopFood & CutString(Trim(rValue(0, i))) & " > "
                End If
            Next i
            If Not IsNull(rValue(0, i)) Then
                strTopFood = strTopFood & CutString(Trim(rValue(0, i))) & " > "
            End If
        Else
            If Not IsNull(rValue(0, 0)) Then
                strTopFood = CutString(Trim(rValue(0, 0)))
            End If
            If Not IsNull(rValue(0, 1)) Then
                strTopFood = strTopFood & " > " & CutString(Trim(rValue(0, 1)))
            End If
            If Not IsNull(rValue(0, 2)) Then
                strTopFood = strTopFood & " > " & CutString(Trim(rValue(0, 2)))
            End If
            If Not IsNull(rValue(0, 3)) Then
                strTopFood = strTopFood & " > " & CutString(Trim(rValue(0, 3)))
            End If
            If Not IsNull(rValue(0, 4)) Then
                strTopFood = strTopFood & " > " & CutString(Trim(rValue(0, 4)))
            End If
        End If
    End If

    Set clsSelect = Nothing
    RPT_TopFood = strTopFood
End Function

Private Function CutString(strOriginal) As String
    If Len(strOriginal) > 5 Then
        CutString = Left(strOriginal, 4) & "..."
    Else
        CutString = strOriginal
    End If
End Function

'<7> 상세평가 - 19가지 식품군별 섭취횟수 //////////////////////////////////////////////////////////
Private Sub NutrionGroup()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect

    Call Initial19

    qrySelect = "SELECT btname,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12,m13,m14,m15,m16,m17,m18,m19 "
    qrySelect = qrySelect & "FROM NutrionGroup "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY bt;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            For j = 0 To 19
                grd19.TextMatrix(i + 1, j) = rValue(j, i)
            Next j
        Next i
        grd19.RowHeight(-1) = 450
        grd19.RowHeight(0) = 570
    End If

    Set clsSelect = Nothing
End Sub

Private Sub Initial19()
    Dim i As Integer

    With grd19
        .Clear
        .BackColorBkg = vbWhite
        .BackColorFixed = FRM_GRAY
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .GridColor = FRM_GRAY
        
        .Rows = 6
        .Cols = 20
        .WordWrap = True
        .ColWidth(-1) = 700
        .ColWidth(0) = 500
        For i = 0 To 19

            .ColAlignmentFixed(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "곡류 및" & vbNewLine & "그 제품"
        .TextMatrix(0, 2) = "감자 및" & vbNewLine & "전분류"
        .TextMatrix(0, 3) = "당류 및" & vbNewLine & "그 제품"
        .TextMatrix(0, 4) = "두류 및" & vbNewLine & "그 제품"
        .TextMatrix(0, 5) = "종실류 및" & vbNewLine & "그 제품"
        .TextMatrix(0, 6) = "채소류"
        .TextMatrix(0, 7) = "버섯류"
        .TextMatrix(0, 8) = "과실류"
        .TextMatrix(0, 9) = "육류 및" & vbNewLine & "그 제품"
        .TextMatrix(0, 10) = "난류"
        .TextMatrix(0, 11) = "어패류"
        .TextMatrix(0, 12) = "해조류"
        .TextMatrix(0, 13) = "우유류" & vbNewLine & "및 유제품"
        .TextMatrix(0, 14) = "유지류"
        .TextMatrix(0, 15) = "음료 및" & vbNewLine & "주류"
        .TextMatrix(0, 16) = "조미료"
        .TextMatrix(0, 17) = "조리가공" & vbNewLine & "식품류"
        .TextMatrix(0, 18) = "이유식류"
        .TextMatrix(0, 19) = "기타"
    End With
End Sub

'<6> 상세평가 - 동물성/식물성 섭취비율 /////////////////////////////////////////////////////////////
Private Sub AniVegRate()
'영앙소별 동,식물성 섭취비율
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    Dim intTemp As Integer, sngTemp2 As Single, sngSum As Single
    Call InitialAniVegRate
    Set clsSelect = New clsSelect
    qrySelect = "SELECT btname,m29,m30,m31,m32,m33,m34,m35,m36 FROM Nutrion "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND bt < 6 ORDER BY bt ASC;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            With grdAniVegRate
                .Rows = .Rows + 1
                .TextMatrix(i + 2, 0) = Trim(rValue(0, i))
                For j = 1 To 7 Step 2
                    sngSum = rValue(j, i) + rValue(j + 1, i)
                    If sngSum = 0 Then
                        .TextMatrix(i + 2, j) = "0.0" & vbNewLine & "(0)"
                        .TextMatrix(i + 2, j + 1) = "0.0" & vbNewLine & "(0)"
                   Else
                        sngTemp2 = rValue(j, i)
                        intTemp = CInt(sngTemp2 / sngSum * 100)
                        .TextMatrix(i + 2, j) = Format(sngTemp2, "0.0") & vbNewLine & "(" & intTemp & ")"
                        sngTemp2 = rValue(j + 1, i)
                        intTemp = CInt(sngTemp2 / sngSum * 100)
                        .TextMatrix(i + 2, j + 1) = Format(sngTemp2, "0.0") & vbNewLine & "(" & intTemp & ")"
                    End If
                Next j
            End With
        Next i
        grdAniVegRate.Rows = grdAniVegRate.Rows - 1
        grdAniVegRate.RowHeight(-1) = 450
        grdAniVegRate.RowHeight(0) = 380
    End If
    Set clsSelect = Nothing
End Sub

Private Sub InitialAniVegRate()
    Dim i As Integer
    With grdAniVegRate
        .Clear
        .BackColorBkg = vbWhite
        .BackColorFixed = FRM_GRAY
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .GridColor = FRM_GRAY
        
        .Rows = 3
        .Cols = 9
        .FixedCols = 1
        .FixedRows = 2
        .WordWrap = True
        .ColWidth(0) = 1020
        .RowHeight(-1) = 450
        .RowHeight(0) = 300
        .MergeCells = flexMergeFree
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        For i = 1 To 8
            .ColWidth(i) = 850
            .ColAlignmentFixed(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
        .MergeRow(0) = True
        .MergeCol(0) = True
        .TextMatrix(0, 1) = "단백질"
        .TextMatrix(0, 2) = "단백질"
        .TextMatrix(0, 3) = "지방"
        .TextMatrix(0, 4) = "지방"
        .TextMatrix(0, 5) = "칼슘"
        .TextMatrix(0, 6) = "칼슘"
        .TextMatrix(0, 7) = "철분"
        .TextMatrix(0, 8) = "철분"

        For i = 1 To 8 Step 2
            .TextMatrix(1, i) = "동물성" & vbNewLine & "g(%)"
            .TextMatrix(1, i + 1) = "식물성" & vbNewLine & "g(%)"
        Next i
    End With
End Sub

'<5> 상세평가 - TOP Food ////////////////////////////////////////////////////////////////
Private Sub TopFood(strNutrition As String)
    Dim qrySelect As String, rValue As Variant
    Dim intMealCal As Single, strFldNutrition As String
    Dim i As Integer

    Set clsSelect = New clsSelect
    Call InitialMealRank
    Select Case strNutrition
        Case "칼로리"
            strFldNutrition = "tblFood.Energy"
            grdMealRank.TextMatrix(0, 3) = "kcal"
        Case "탄수화물"
            strFldNutrition = "tblFood.Carbohy"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "단백질"
            strFldNutrition = "tblFood.Protein"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "지방"
            strFldNutrition = "tblFood.Fat"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "비타민A"
            strFldNutrition = "tblFood.Vitamine_A"
            grdMealRank.TextMatrix(0, 3) = "R.E"
        Case "비타민C"
            strFldNutrition = "tblFood.Vitamine_C"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "비타민B1"
            strFldNutrition = "tblFood.Vitamine_B1"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "비타민B2"
            strFldNutrition = "tblFood.Vitamine_B2"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "비타민E"
            strFldNutrition = "tblFood.Vitamine_E"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "나이아신"
            strFldNutrition = "tblFood.Niacin"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "비타민B6"
            strFldNutrition = "tblFood.Vitamine_B6"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "엽산"
            strFldNutrition = "tblFood.Folic"
            grdMealRank.TextMatrix(0, 3) = "ug"
        Case "칼슘"
            strFldNutrition = "tblFood.Ca"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "인"
            strFldNutrition = "tblFood.P"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "철"
            strFldNutrition = "tblFood.Fe"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "아연"
            strFldNutrition = "tblFood.Zn"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "나트륨"
            strFldNutrition = "tblFood.Na"
            grdMealRank.TextMatrix(0, 3) = "mg"
        '****** 영양소 항목 추가
        Case "콜레스테롤"
            strFldNutrition = "tblFood.Cholesterol"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "포화지방"
            strFldNutrition = "tblFood.SFA"
            grdMealRank.TextMatrix(0, 3) = "g"
    End Select

    qrySelect = "SELECT DISTINCT(MealName),"
    qrySelect = qrySelect & "Portion, Unit,"
    qrySelect = qrySelect & "SUM((DietFood.FoodWeight*" & strFldNutrition & ")/100) AS a "
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "INNER JOIN tblMeal ON DietMeal.MealCode=tblMeal.MealID "
    qrySelect = qrySelect & "INNER JOIN DietFood ON DietMeal.DietMealNum=DietFood.DietMealNum "
    qrySelect = qrySelect & "INNER JOIN tblFood ON DietFood.FoodCode=tblFood.FoodID "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum

    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpEnd.Value, "YYYY-MM-DD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate, MealSection, MealName, Portion, Unit "
    qrySelect = qrySelect & "ORDER BY a DESC;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
        With grdMealRank
            .Rows = .Rows + 1
            .TextMatrix(i + 1, 0) = i + 1
            .TextMatrix(i + 1, 1) = Trim(rValue(0, i))
            .TextMatrix(i + 1, 2) = Trim(rValue(1, i)) & Trim(rValue(2, i))
            .TextMatrix(i + 1, 3) = CInt(rValue(3, i))
        End With
        Next i
        grdMealRank.Rows = grdMealRank.Rows - 1
        grdMealRank.RowHeight(-1) = 300
    End If

    Set clsSelect = Nothing
End Sub

'<4> 영양소별 권장량 대비////////////////////////////////////////////////////////////////
Private Sub NutritionCompare()
    Dim qrySelect As String
    Dim rValue As Variant, rValue2 As Variant
    Dim intDayCount As Integer


On Error GoTo ShowErr
    Call InitialChartNutrition
    '고객정보에서 다이어트열량, 신체상태, 나이, 성별등을 불러옴

    With typCustomerInfo
            Set clsSelect = New clsSelect
            qrySelect = "SELECT m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,"
            qrySelect = qrySelect & "m11,m12,m13,m14,m15,m16,m17,m18,m19,m20,m21,m22 "
            qrySelect = qrySelect & "FROM Nutrion WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND bt=6;"

            rValue = clsSelect.Query(qrySelect)
            '정해진 기간동안의 영양소값이 모두 더해진 것이므로
            '기간내 입력된 일수만큼 나눠서 일별 영양소 평균치를 구한다.
            If Not IsNull(rValue) Then
'칼로리,단백질,비타민A,비타민E,비타민C,비타민B1,비타민B2,나이아신,비타민B6,엽산,칼슘,인,철,아연
'1,2,13,22,20,16,17,19,18,21,7,8,9,12
                With chtNutrition
                    .OpenDataEx COD_VALUES, 1, 1
                    .Axis(AXIS_X).Label(0) = "칼로리"
                    .Axis(AXIS_X).Label(1) = "단백질"
                    .Axis(AXIS_X).Label(2) = "VitA"
                    .Axis(AXIS_X).Label(3) = "VitE"
                    .Axis(AXIS_X).Label(4) = "VitC"
                    .Axis(AXIS_X).Label(5) = "VitB1"
                    .Axis(AXIS_X).Label(6) = "VitB2"
                    .Axis(AXIS_X).Label(7) = "나이아신"
                    .Axis(AXIS_X).Label(8) = "VitB6"
                    .Axis(AXIS_X).Label(9) = "엽산"
                    .Axis(AXIS_X).Label(10) = "칼슘"
                    .Axis(AXIS_X).Label(11) = "인"
                    .Axis(AXIS_X).Label(12) = "철"
                    .Axis(AXIS_X).Label(13) = "아연"
                    
                    .Value(0) = rValue(0, 0) / 10
                    .Value(1) = rValue(1, 0) / 10
                    .Value(2) = rValue(12, 0) / 10
                    .Value(3) = rValue(21, 0) / 10
                    .Value(4) = rValue(19, 0) / 10
                    .Value(5) = rValue(15, 0) / 10
                    .Value(6) = rValue(16, 0) / 10
                    .Value(7) = rValue(18, 0) / 10
                    .Value(8) = rValue(17, 0) / 10
                    .Value(9) = rValue(20, 0) / 10
                    .Value(10) = rValue(6, 0) / 10
                    .Value(11) = rValue(7, 0) / 10
                    .Value(12) = rValue(8, 0) / 10
                    .Value(13) = rValue(11, 0) / 10

                    .CloseData COD_VALUES
                End With
            End If
            Set clsSelect = Nothing
    End With

    Exit Sub
ShowErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "NutritionCompare", "frmCounsel_4", Err.Number, Err.Description
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub LoadCustomerInfo(lngCustomerNum As Long)
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    qrySelect = "SELECT TOP 1 BodyData.BodyStatus, Age, Sex, Treat.TreatCalory "
    qrySelect = qrySelect & "FROM CustomerInfo INNER JOIN BodyData "
    qrySelect = qrySelect & "ON CustomerInfo.CustomerNum=BodyData.CustomerNum INNER JOIN "
    qrySelect = qrySelect & "Treat ON BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE CustomerInfo.CustomerNum=" & lngCustomerNum
    qrySelect = qrySelect & " AND NOT Treat.TreatCalory IS NULL"
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
 
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typCustomerInfo
            .intState = CInt(rValue(0, 0))
            .intAge = CInt(rValue(1, 0))
            .strSex = Trim(rValue(2, 0))
            .sngDietCal = Is_Null(rValue(3, 0), 0)
        End With
    End If

    Set clsSelect = Nothing
End Sub

'<3> 끼니별 비교/////////////////////////////////////////////////////////////////////////
Private Sub MealSectionRate()
    Dim cfxArray As CfxDataArray
    
    Dim qrySelect As String, rValue As Variant
    Dim intTotal As Integer, intBF As Integer, intLunch As Integer, intDinner As Integer, intSnack As Integer
    Dim intMeal(4) As Integer, strMeal(4) As String, i As Integer, sngMeal(4) As Single
    

On Error GoTo ShowErr
    Call InitialChartRate
    Set clsSelect = New clsSelect

    qrySelect = "SELECT MealDate, SUM(a), SUM(b), SUM(c), SUM(d) FROM "
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(d) FROM "
    qrySelect = qrySelect & vbCrLf & "(SELECT MealDate, SUM(Calories) AS a, 0 AS b, 0 AS c, 0 AS d "
    qrySelect = qrySelect & vbCrLf & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & vbCrLf & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & vbCrLf & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & vbCrLf & " AND MealSection=1 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & vbCrLf & "UNION ALL "
    qrySelect = qrySelect & vbCrLf & "SELECT MealDate, 0 AS a, SUM(Calories) AS b, 0 AS c, 0 AS d "
    qrySelect = qrySelect & vbCrLf & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & vbCrLf & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & vbCrLf & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & vbCrLf & " AND MealSection=2 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & vbCrLf & "UNION ALL "
    qrySelect = qrySelect & vbCrLf & "SELECT MealDate, 0 AS a, 0 AS b, SUM(Calories) AS c, 0 AS d "
    qrySelect = qrySelect & vbCrLf & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & vbCrLf & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & vbCrLf & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & vbCrLf & " AND MealSection=3 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & vbCrLf & "UNION ALL "
    qrySelect = qrySelect & vbCrLf & "SELECT MealDate, 0 AS a, 0 AS b, 0 AS c, SUM(Calories) AS d "
    qrySelect = qrySelect & vbCrLf & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & vbCrLf & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & vbCrLf & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & vbCrLf & " AND MealSection=4 GROUP BY MealDate, MealSection) AS park "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & vbCrLf & "WHERE MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & vbCrLf & "WHERE MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYY-MM-DD") & "' AND '" & Format(dtpEnd.Value, "YYYY-MM-DD") & "' "
    End If
     
    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
        For i = 0 To 3
            intMeal(i) = Is_Null(rValue(i, 0), 0)
            intTotal = intTotal + intMeal(i)
        Next i
        For i = 0 To 3
            If intTotal <> 0 Then sngMeal(i) = intMeal(i) / intTotal
        Next i
        strMeal(0) = "아침"
        strMeal(1) = "점심"
        strMeal(2) = "저녁"
        strMeal(3) = "간식"
    End If

    If intTotal = 0 Then
        Exit Sub
    End If
    
    Set cfxArray = CreateObject("cfxData.Array")

    '챠트 보여주기
    cfxArray.AddArray intMeal
    cfxArray.AddArray strMeal
    With chtCaloryRate
        .GetExternalData cfxArray
        
        .OpenDataEx COD_COLORS, 4, 0
        
        .Color(0) = &HEFFBFF
        .Color(1) = &HCEF3E7
        .Color(2) = &H84E3C6
        .Color(3) = &H42CF9C
        
        .CloseData COD_COLORS
        .PointLabels = True
        .PointLabelAlign = LA_BASELINE + LA_RIGHT
        .PointLabelsFont.Size = 8
        .RGBFont(CHART_VALUESFT) = vbBlack
    End With
    Set clsSelect = Nothing
    Exit Sub
ShowErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "MealSectionRate", "frmCounsel_4", Err.Number, Err.Description
    MsgBox Err.Description
End Sub

Private Sub cmbTopFood_Change()
    Call TopFood(Trim(cmbTopFood.Text))
End Sub

Private Sub cmbTopFood_Click()
    Call TopFood(Trim(cmbTopFood.Text))
End Sub

Private Sub ShowNutritionInfo()
' 영양소 섭취 상태 =============///
' 0 : 열량
' 1 : 탄수화물:단백질:지방
' 2 : 총지방량(비율)
' 3 : 콜레스테롤
' 4 : 나트륨
    Dim qrySelect As String, rValue As Variant
    Dim sngTotal As Single, sngC As Single, sngP As Single, sngF As Single
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT m1, m2, m3, m4, m23, m10 FROM Nutrion "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND bt=5;"
    
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
    '1) 섭취량
    lblIntake(0).Caption = Format(Is_Null(rValue(0, 0), 0), "#,###")
    sngTotal = (rValue(1, 0) * 4) + (rValue(2, 0) * 9) + (rValue(3, 0) * 4)
    sngC = (rValue(3, 0) * 4) / sngTotal
    sngP = (rValue(1, 0) * 4) / sngTotal
    sngF = (rValue(2, 0) * 9) / sngTotal
    sngC = sngC * 100
    sngP = sngP * 100
    sngF = sngF * 100
    lblIntake(1).Caption = CInt(sngC) & ":" & CInt(sngP) & ":" & CInt(sngF)
    lblIntake(2).Caption = CInt(Is_Null(rValue(2, 0), 0)) & "g (" & CInt(sngF) & "%)"
    lblIntake(3).Caption = CInt(Is_Null(rValue(4, 0), 0))
    lblIntake(4).Caption = Format(Is_Null(rValue(5, 0), 0), "#,###")
    End If
    
    '2) 권장량
    lblRecommend(0).Caption = Format(WhatTreatCalory, "#,###")
    lblRecommend(1).Caption = "65:15:20"
    lblRecommend(2).Caption = "총열량의15~20%"
    lblRecommend(3).Caption = "300mg이하"
    lblRecommend(4).Caption = "2,400mg" & vbNewLine & "이하"
    
    '3) Top Food
    lblTopFood(0).Caption = RPT_TopFood("열량")
    lblTopFood(1).Caption = ""
    lblTopFood(2).Caption = RPT_TopFood("지방")
    lblTopFood(3).Caption = RPT_TopFood("콜레스테롤")
    lblTopFood(4).Caption = RPT_TopFood("나트륨")
End Sub

Private Sub ShowValuation()
    '고객정보에서 다이어트열량, 신체상태, 나이, 성별등을 불러옴
   Call LoadCustomerInfo(glngCustomerNum)
    If typCustomerInfo.sngDietCal = 0 Then
        Exit Sub
    End If

    With typCustomerInfo
        If Calculate_Nut(.sngDietCal, .intState, .intAge, .strSex) = True Then
            '끼니별평가
            Call MealSectionRate
            '1) 영양소별 권장량대비
            Call NutritionCompare
            '2) 동식물성 섭취비율
            Call AniVegRate
            '3) 19가지 식품군별 섭취횟수
            Call NutrionGroup
            '4) 섭취음식(표)
            Call LoadMealCalory
            '5) 섭취칼로리(그래프)
            Call ShowCaloryChart
        End If
    End With
End Sub

Private Sub ShowCaloryChart()
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant
    Dim sngMin As Single, sngMax As Single
    
    Call InitialChart
    
    qrySelect = "SELECT MAX(a), MIN(a) FROM ( "
    qrySelect = qrySelect & "SELECT SUM(MealCalory) AS a FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " GROUP BY MealDate) total;"
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        sngMax = CInt(rValue(0, 0)) + 10
        sngMin = CInt(rValue(1, 0)) - 10
    Else
        sngMin = 300
        sngMax = 3000
    End If
    Set clsSelect = Nothing
    '식사일기 입력한 날과 당시 총칼로리를 보여줌
    qrySelect = "SELECT MealDate, SUM(MealCalory) FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "'"
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate ORDER BY MealDate ASC;"
    
    Chart.Title(CHART_TOPTIT) = "섭취칼로리 변화"

    Set clsSelect = New clsSelect

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        Chart.OpenDataEx COD_VALUES, 0, COD_UNKNOWN
        Chart.Axis(AXIS_Y).Min = sngMin
        Chart.Axis(AXIS_Y).Max = sngMax
        For i = 0 To UBound(rValue, 2)
            Chart.ValueEx(0, i) = Is_Null(rValue(1, i), 0)
            If cmbPeriod.ListIndex = 0 Then
                Chart.Axis(AXIS_X).Label(i) = Is_Null(rValue(0, i), "")
            Else
                Chart.Axis(AXIS_X).Label(i) = Format(Is_Null(rValue(0, i), ""), "M/D")
            End If
        Next i
        Chart.CloseData COD_VALUES
    Else
        MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
       Chart.Visible = False
    End If

    Set clsSelect = Nothing
End Sub

Private Sub LoadMealCalory()
'현재 해당환자의 저장된 식사일기들을 선택된 기간내에 해당하는 것을 불러서 보여준다
    Dim qrySelect As String
    Dim rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect

    qrySelect = "SELECT DISTINCT MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "';"
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "';"
    End If

    rValue = clsSelect.Query(qrySelect)

    Call InitialMealCalory
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            Call LoadDayMeal(CStr(rValue(0, i)))
        Next i
        grdMealCalory.Rows = grdMealCalory.Rows - 1
    Else
        MsgBox "기간내에 입력된 식사일기가 없습니다.", vbInformation

        Call InitialChartRate
        Call InitialChartNutrition
        Call InitialMealRank
        Call InitialAniVegRate
        Call Initial19
        Call InitialChart
        Exit Sub
    End If

    Set clsSelect = Nothing
End Sub

Private Sub LoadDayMeal(strMealDate As String)
    Dim qrySelect As String
    Dim rValue As Variant
    Dim i As Integer, j As Integer
    Dim row1 As Integer, row2 As Integer, row3 As Integer, row4 As Integer
    Dim intMax As Integer, intStartRow As Integer
    Dim intTotal(4) As Integer

    Set clsSelect = New clsSelect

    qrySelect = "SELECT MAX(a) FROM("
    qrySelect = qrySelect & "SELECT COUNT(DietMeal.DietMealNum) AS a "
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN "
    qrySelect = qrySelect & "DietMeal ON DietDiary.DietDiaryNum = DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "Where DietDiary.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " And DietDiary.MealDate='" & strMealDate
    qrySelect = qrySelect & "' GROUP BY DietDiary.MealSection) b"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        intMax = rValue(0, 0)
    End If

    qrySelect = "SELECT MealSection, MealName, "
    qrySelect = qrySelect & "Portion "
    qrySelect = qrySelect & ", Calories "
    qrySelect = qrySelect & ", Unit "
    qrySelect = qrySelect & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum LEFT JOIN tblMeal "
    qrySelect = qrySelect & "ON tblMeal.MealID=DietMeal.MealCode "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealDate='" & strMealDate & "'"
    qrySelect = qrySelect & " ORDER BY MealSection;"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        With grdMealCalory
            intStartRow = .Rows - 1
            row1 = intStartRow: row2 = intStartRow: row3 = intStartRow: row4 = intStartRow
            .Rows = .Rows + intMax
            For i = intStartRow To .Rows - 2
                .TextMatrix(i, 0) = Format(strMealDate, "M월D일")
            Next i
            For i = 0 To UBound(rValue, 2)
                Select Case rValue(0, i)
                Case 1:
                    .TextMatrix(row1, 1) = Is_Null(rValue(1, i), "") & Is_Null(rValue(2, i), "") & Is_Null(rValue(4, i), "")
                    .TextMatrix(row1, 3) = CInt(Is_Null(rValue(3, i), 0))
                    intTotal(0) = intTotal(0) + Is_Null(rValue(3, i), 0)
                    row1 = row1 + 1
                Case 2:
                    .TextMatrix(row2, 4) = Is_Null(rValue(1, i), "") & Is_Null(rValue(2, i), "") & Is_Null(rValue(4, i), "")
                    .TextMatrix(row2, 6) = CInt(Is_Null(rValue(3, i), 0))
                    intTotal(1) = intTotal(1) + Is_Null(rValue(3, i), 0)
                    row2 = row2 + 1
                Case 3:
                    .TextMatrix(row3, 7) = Is_Null(rValue(1, i), "") & Is_Null(rValue(2, i), "") & Is_Null(rValue(4, i), "")
                    .TextMatrix(row3, 9) = CInt(Is_Null(rValue(3, i), 0))
                    intTotal(2) = intTotal(2) + Is_Null(rValue(3, i), 0)
                    row3 = row3 + 1
                Case 4:
                    .TextMatrix(row4, 10) = Is_Null(rValue(1, i), "") & Is_Null(rValue(2, i), "") & Is_Null(rValue(4, i), "")
                    .TextMatrix(row4, 12) = CInt(Is_Null(rValue(3, i), 0))
                    intTotal(3) = intTotal(3) + Is_Null(rValue(3, i), 0)
                    row4 = row4 + 1
                End Select
            Next i
            '하루치의 총계 보여주기
            .TextMatrix(.Rows - 1, 0) = "총합"
            For i = 1 To .Cols - 1
                .Row = .Rows - 1
                .Col = i
                .CellBackColor = &HC69AA5
            Next i
            For i = 0 To 3
                .TextMatrix(.Rows - 1, (i + 1) * 3) = intTotal(i)
            Next i
            .Rows = .Rows + 1
            .RowHeight(-1) = 300
        End With
    End If
    Set clsSelect = Nothing
End Sub

Private Sub InitialDailyCombo()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT DISTINCT(MealDate) FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    
    rValue = clsSelect.Query(qrySelect)
    cmbDaily.Clear
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            cmbDaily.AddItem Trim(rValue(0, i))
        Next i
        cmbDaily.ListIndex = UBound(rValue, 2)
    End If
    Set clsSelect = Nothing
End Sub

Private Sub InitialMealCalory()
    Dim i As Integer

    With grdMealCalory
        .Clear
        .BackColorBkg = vbWhite
        .BackColorFixed = FRM_GRAY
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .GridColor = FRM_GRAY

        .Rows = 2
        .Cols = 13
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .FixedCols = 1
        .FixedRows = 1
        .MergeCells = flexMergeFree
        For i = 0 To 12
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next i
        .ColAlignment(0) = flexAlignLeftCenter
        .MergeCol(0) = True
        .MergeRow(0) = True

        .TextMatrix(0, 0) = "식사일"

        For i = 1 To 3
            .TextMatrix(0, i) = "아침"
        Next i
        For i = 4 To 6
            .TextMatrix(0, i) = "점심"
        Next i
        For i = 7 To 9
            .TextMatrix(0, i) = "저녁"
        Next i
        For i = 10 To 12
            .TextMatrix(0, i) = "간식"
        Next i
        .ColWidth(0) = 800
        For i = 1 To 12 Step 3
            .ColWidth(i) = 1300
        Next i
        For i = 2 To 12 Step 3
            .ColWidth(i) = 0
        Next i
        For i = 3 To 12 Step 3
            .ColWidth(i) = 400
        Next i
    End With
End Sub

Private Sub InitialChartRate()
    With chtCaloryRate
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = PIE
        .Chart3D = True
        .TypeMask = .TypeMask Or CT_POINTLABELS

        ' Color Settings
        .Border = False
        .RGBBk = &HEFEFEF
        .RGB2DBk = CHART_TRANSPARENT
        .MultipleColors = True

        .ToolBar = False
        .Title(CHART_TOPTIT) = ""
        
        .TopGap = 10
        .BottomGap = 10
        .LeftGap = 50
        .RightGap = 50
        
        .PointLabelAlign = LA_RIGHT Or LA_TOP
        .Axis(AXIS_Y).Decimals = 0

        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialChartNutrition()
    With chtNutrition
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = BAR
        .Chart3D = False
        .Stacked = CHART_NOSTACKED
        .Border = True
        .RGBBk = RGB(255, 255, 255)

        .TopGap = 0
        .LeftGap = 0
        .RightGap = 0
        .BottomGap = 0
       
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""
        .PointLabels = True
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialMealRank()
    With grdMealRank
        .Clear
        .BackColorBkg = vbWhite
        .BackColorFixed = FRM_GRAY
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .GridColor = FRM_GRAY
        
        .Cols = 4
        .Rows = 2
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter

        .FixedRows = 1
        .FixedCols = 1
        .ColWidth(0) = 700
        .ColWidth(1) = 3800
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500

        .TextMatrix(0, 0) = "순위"
        .TextMatrix(0, 1) = "음식명"
        .TextMatrix(0, 2) = "인분"
        .TextMatrix(0, 3) = "열량"   '선택한 항목으로 바뀜
    End With
End Sub

'영양평가를 위한 테이블
Private Function Calculate_Nut(sngDietCal As Single, intState As Integer, intAge As Integer, strSex As String) As Boolean
    Dim rValue As Variant
    Dim rValue2 As Variant
    Dim qrySelect As String

    '   s(영양요소, 끼니코드 1=아침   4=간식, 0=합)
    Dim s(1 To 36, 0 To 4) As Single, temp As Single
    '   각 끼니를 해당기간내에 먹은 횟수 1:아침~4:간식,0은 기간내 입력한 일기수
    '   해당기간내에 평균을 내기 위함..
    Dim intSectionCnt(0 To 4) As Integer

    ' 각 영양소, 끼니별 count
    Dim c(0 To 19, 0 To 4) As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim strEelem(1 To 4) As String
    Dim strTime(1 To 6) As String

    For j = 0 To 4
        For i = 1 To 36
            s(i, j) = 0
        Next i
        For i = 0 To 19
            c(i, j) = 0
        Next i
    Next j
    strTime(1) = "아침"
    strTime(2) = "점심"
    strTime(3) = "저녁"
    strTime(4) = "간식"
    strTime(5) = "1일 합계"
    strTime(6) = "권장량대비%"

    strEelem(1) = "열량"
    strEelem(2) = "단백질"
    strEelem(3) = "지방"
    strEelem(4) = "탄수화물"

    Set clsSelect = New clsSelect

    qrySelect = "SELECT Count(a.MealDate) FROM"
    qrySelect = qrySelect & "(SELECT MealDate FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND '" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate) a"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        intSectionCnt(0) = CInt(rValue(0, 0))
    End If

    For i = 1 To 4
        qrySelect = "SELECT COUNT(DietDiaryNum) FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND '" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " AND MealSection=" & i
        qrySelect = qrySelect & " AND MealCalory IS NOT NULL;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            intSectionCnt(i) = CInt(rValue(0, 0))
        End If
    Next i
    qrySelect = "SELECT MealCode,"                                                       '0
    qrySelect = qrySelect & "Energy,Protein,Fat,Carbohy,Fiber,"                          '1-5
    qrySelect = qrySelect & "Ash,Ca,P,Fe,Na,"                                            '6-10
    qrySelect = qrySelect & "K,Zn,Vitamine_A,Retinol,Carotene,"                          '11-15
    qrySelect = qrySelect & "Vitamine_B1,Vitamine_B2,Vitamine_B6,Niacin,Vitamine_C,"     '16-20
    qrySelect = qrySelect & "Folic,Vitamine_E,Cholesterol,Waste,DietFiberDry,"           '21-25
    qrySelect = qrySelect & "DietFiberWet,Vitamine_B12,Vitamine_D,MealSection,FoodCode," '26-30
    qrySelect = qrySelect & "FoodWeight,FK_PartID,NatureID "                             '31,32,33
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "INNER JOIN DietFood ON DietMeal.DietMealNum=DietFood.DietMealNum "
    qrySelect = qrySelect & "INNER JOIN tblFood ON DietFood.FoodCode=tblFood.FoodID "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "';"
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND MealDate BETWEEN '" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND '" & Format(dtpEnd.Value, "YYYYMMDD") & "';"
    End If
    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            'k is 끼니
            k = rValue(29, i)
            For j = 1 To 28
                If Not IsNull(rValue(j, i)) Then
                    '모든 영양소는 100g에 대한 값임
                    temp = rValue(j, i) * rValue(31, i) / 100
                    s(j, k) = s(j, k) + temp
                    s(j, 0) = s(j, 0) + temp
                Else
                    temp = 0
                End If

                '동물성/식물성 섭취비율
                If j = 2 Or j = 3 Or j = 7 Or j = 9 Then
                    Select Case j
                       Case 2        '단백질
                            l = 29
                       Case 3        '지방
                            l = 31
                       Case 7        '칼슘
                            l = 33
                       Case 9        '철분
                            l = 35
                    End Select

                    If rValue(33, i) = "1" Then      '식물성
                        If Not IsNull(rValue(j, i)) Then
                            s(l, k) = s(l, k) + temp
                            s(l, 0) = s(l, 0) + temp
                        End If
                    ElseIf rValue(33, i) = "2" Then  '동물성
                        If Not IsNull(rValue(j, i)) Then
                            s(l + 1, k) = s(l + 1, k) + temp
                            s(l + 1, 0) = s(l + 1, 0) + temp
                        End If
                    End If
                End If
            Next j
            c(rValue(32, i), k) = c(rValue(32, i), k) + 1
            c(rValue(32, i), 0) = c(rValue(32, i), 0) + 1
        Next i
        Erase rValue
        '평균을 내자 : 동,식물성섭취비율과 19가지 식품군은 평균안냄
        For j = 0 To 4
            For i = 1 To 28
                If intSectionCnt(j) <> 0 Then
                    s(i, j) = s(i, j) / intSectionCnt(j)
                End If
            Next i
       Next j
        '/////////////

        qrySelect = "DELETE FROM Nutrion WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)

        qrySelect = "DELETE FROM NutrionCont WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)

        If strSex = "M" Then
            intState = 1
        End If

        qrySelect = "SELECT ID, m1, m2, m3, m4, m5,m6, m7, m8, m9, m10,"
        qrySelect = qrySelect & "m11, m12, m13, m14, m15,m16, m17, m18, m19, m20,"
        qrySelect = qrySelect & "m21, m22, m23, m24, m25,m26, m27, m28 "
        qrySelect = qrySelect & "FROM Recommand WHERE Gender ='" & strSex & "' AND "
        qrySelect = qrySelect & "BodyState = " & intState
        qrySelect = qrySelect & " AND AgeLow <= " & intAge & " AND AgeHigh > " & intAge
        rValue2 = clsSelect.Query(qrySelect)

        Set clsSelect = Nothing

        Dim qryInsert As String
        For i = 0 To 4
            If i = 0 Then
                l = 5
            Else
                l = i
            End If
            qryInsert = "INSERT INTO Nutrion(CustomerNum, bt, btname, m1, m2, m3, m4, m5"
            qryInsert = qryInsert & ",m6, m7, m8, m9, m10,m11, m12, m13, m14, m15"
            qryInsert = qryInsert & ",m16, m17, m18, m19, m20,m21, m22, m23, m24, m25"
            qryInsert = qryInsert & ",m26, m27, m28, m29, m30,m31, m32, m33, m34, m35,m36) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & "," & l & ", '" & strTime(l) & "',"
            qryInsert = qryInsert & s(1, i) & "," & s(2, i) & "," & s(3, i) & "," & s(4, i) & "," & s(5, i) & "," & s(6, i) & "," & s(7, i) & "," & s(8, i) & "," & s(9, i) & "," & s(10, i) & ","
            qryInsert = qryInsert & s(11, i) & "," & s(12, i) & "," & s(13, i) & "," & s(14, i) & "," & s(15, i) & "," & s(16, i) & "," & s(17, i) & "," & s(18, i) & "," & s(19, i) & "," & s(20, i) & ","
            qryInsert = qryInsert & s(21, i) & "," & s(22, i) & "," & s(23, i) & "," & s(24, i) & "," & s(25, i) & "," & s(26, i) & "," & s(27, i) & "," & s(28, i) & "," & s(29, i) & "," & s(30, i) & ","
            qryInsert = qryInsert & s(31, i) & "," & s(32, i) & "," & s(33, i) & "," & s(34, i) & "," & s(35, i) & "," & s(36, i) & " )"

            Call modSql.AdoExcuteSql(qryInsert)
        Next i
        i = 0
        qryInsert = "INSERT INTO Nutrion (CustomerNum, bt,btname, m1, m2, m3, m4, m5,m6, m7, m8, m9, m10,m11, m12, m13, m14, m15,m16, m17, m18, m19, m20,m21, m22, m23, m24, m25,m26, m27, m28) "
        qryInsert = qryInsert & "VALUES ( " & glngCustomerNum & "," & 6 & ",'" & strTime(6) & "',"
        qryInsert = qryInsert & s(1, i) / sngDietCal * 100 & "," & s(2, i) / rValue2(2, 0) * 100 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & s(7, i) / rValue2(7, 0) * 100 & "," & s(8, i) / rValue2(8, 0) * 100 & "," & s(9, i) / rValue2(9, 0) * 100 & "," & 0 & ","
        qryInsert = qryInsert & 0 & "," & s(12, i) / rValue2(12, 0) * 100 & "," & s(13, i) / rValue2(13, 0) * 100 & "," & 0 & "," & 0 & "," & s(16, i) / rValue2(16, 0) * 100 & "," & s(17, i) / rValue2(17, 0) * 100 & "," & s(18, i) / rValue2(18, 0) * 100 & "," & s(19, i) / rValue2(19, 0) * 100 & "," & s(20, i) / rValue2(20, 0) * 100 & ","
        qryInsert = qryInsert & s(21, i) / rValue2(21, 0) * 100 & "," & s(22, i) / rValue2(22, 0) * 100 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & " )"
        
        Call modSql.AdoExcuteSql(qryInsert)

        qryInsert = "DELETE FROM NutrionGroup WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qryInsert)

        For i = 0 To 4
            If i = 0 Then
               l = 5
            Else
                l = i
            End If
            qryInsert = "INSERT INTO NutrionGroup(CustomerNum,bt,btname, m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13, m14, m15, m16, m17, m18, m19) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & "," & l & ",'" & strTime(l) & "'," & c(1, i) & "," & c(2, i) & "," & c(3, i) & "," & c(4, i) & "," & c(5, i) & "," & c(6, i) & "," & c(7, i) & "," & c(8, i) & "," & c(9, i) & "," & c(10, i) & ","
            qryInsert = qryInsert & c(11, i) & "," & c(12, i) & "," & c(13, i) & "," & c(14, i) & "," & c(15, i) & "," & c(16, i) & "," & c(17, i) & "," & c(18, i) & "," & c(19, i) & ")"
            Call modSql.AdoExcuteSql(qryInsert)
        Next i

        Dim ss(4) As Single
        Dim ssr(4) As Single
        Dim ssc(4) As Single
        Dim fac(2 To 4) As Integer

        fac(2) = 4
        fac(3) = 9
        fac(4) = 4
        For i = 0 To 4
            ss(i) = 0
            ssr(i) = 0
            ssc(i) = 0
        Next i

        '각 영양소의 중량합을 구한다.(각 끼니별)
        For i = 2 To 4 '2=단백질, 3=지방 4=탄수화물
            ss(1) = ss(1) + s(i, 1) * fac(i) '아침
            ss(2) = ss(2) + s(i, 2) * fac(i) '점심
            ss(3) = ss(3) + s(i, 3) * fac(i) '저녁
            ss(4) = ss(4) + s(i, 4) * fac(i) '간식
        Next i
        For i = 1 To 4 '각 영양소
            If i = 1 Then '열량은 각 끼니 비율임
                For j = 1 To 4 '각 끼니

                     ssr(j) = Round(s(i, j) / s(i, 0) * 100, 2)
                Next j
            Else '영양소는 한끼니에서 각 영양소의 열량 비율임.
                For j = 1 To 4 '각 끼니
                    If ss(j) = 0 Then
                        ssr(j) = 0
                    Else
                        ssr(j) = Round(s(i, j) * fac(i) / ss(j) * 100, 2)
                    End If
                Next j
            End If

            qryInsert = "INSERT INTO NutrionCont(CustomerNum, element, m1, m2, m3, m4, m5,m6, m7, m8, m9) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & ",'" & strEelem(i) & "'," & s(i, 1) & "," & ssr(1) & "," & s(i, 2) & "," & ssr(2) & ","
            qryInsert = qryInsert & s(i, 3) & "," & ssr(3) & "," & s(i, 4) & "," & ssr(4) & "," & s(i, 0) & ")"
            Call modSql.AdoExcuteSql(qryInsert)
        Next i
        Calculate_Nut = True
    Else
        Calculate_Nut = False
    End If
End Function

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub InitialChart()
    With Chart
        ' Chart Type Settings
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_RECT
        .MarkerSize = 3
        .AxesStyle = CAS_FLATFRAME
        .Axis(AXIS_X).Grid = True
        .Axis(AXIS_Y).Grid = True

        ' Color Settings
        .Border = False
        .RGBBk = vbWhite

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .PointLabels = True
        .MultipleColors = False
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub imgAppend_Click(Index As Integer)
    frmPop_Additional1.mintNumber = Index + 1
    frmPop_Additional1.Show vbModal
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH04 & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH04 & IMG_PRINT_OFF)
    Call PrintData
End Sub

Private Sub imgStart_Click()
    Dim datTemp As Date
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    '기간별 평가보기
    If cmbPeriod.ListIndex = 1 Then
        If dtpBegin.Value > dtpEnd.Value Then
            datTemp = dtpBegin.Value
            dtpBegin.Value = dtpEnd.Value
            dtpEnd.Value = datTemp
        End If
    End If
    
    qrySelect = "SELECT DISTINCT MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "';"
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "';"
    End If
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)

    Call InitialMealCalory
    If Not IsNull(rValue) Then
        For i = 0 To 5
            imgSub(i).Enabled = True
        Next i
        Call imgSub_Click(0)
        chtCaloryRate.Visible = True
        Call ShowValuation
        Call ShowNutritionInfo
    Else
        MsgBox "기간내에 입력된 식사일기가 없습니다.", vbInformation
        For i = 0 To 4
            lblIntake(i).Caption = ""
            lblRecommend(i).Caption = ""
            lblTopFood(i).Caption = ""
        Next i
        grdAniVegRate.Visible = False
        grd19.Visible = False
        grdMealCalory.Visible = False
        grdMealRank.Visible = False
        cmbTopFood.Visible = False
        Chart.Visible = False
        chtNutrition.Visible = False
        For i = 0 To 5
            imgSub(i).Enabled = False
        Next i
        chtCaloryRate.Visible = False
    End If

    Set clsSelect = Nothing
End Sub

Private Sub imgSub_Click(Index As Integer)
    Select Case Index
        Case 0:  '영양소별 권장량대비
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_ON)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_OFF)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_OFF)
            
            grdAniVegRate.Visible = False
            grd19.Visible = False
            grdMealCalory.Visible = False
            grdMealRank.Visible = False
            cmbTopFood.Visible = False
            Chart.Visible = False
            
            chtNutrition.Visible = True
        Case 1:  '동식물성섭취비율
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_ON)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_OFF)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_OFF)
            
            chtNutrition.Visible = False
            grd19.Visible = False
            grdMealCalory.Visible = False
            grdMealRank.Visible = False
            cmbTopFood.Visible = False
            Chart.Visible = False
            
            grdAniVegRate.Visible = True
        Case 2:  '19가지식품군별섭취횟수
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_ON)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_OFF)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_OFF)
            
            chtNutrition.Visible = False
            grdAniVegRate.Visible = False
            grdMealCalory.Visible = False
            grdMealRank.Visible = False
            cmbTopFood.Visible = False
            Chart.Visible = False
           
            grd19.Visible = True
        Case 3:  '섭취음식(표)
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_ON)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_OFF)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_OFF)
            
           chtNutrition.Visible = False
            grdAniVegRate.Visible = False
            grd19.Visible = False
            grdMealRank.Visible = False
            cmbTopFood.Visible = False
            Chart.Visible = False
            
            grdMealCalory.Visible = True
        Case 4:  '섭취칼로리(그래프)
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_ON)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_OFF)

            chtNutrition.Visible = False
            grdAniVegRate.Visible = False
            grd19.Visible = False
            grdMealCalory.Visible = False
            grdMealRank.Visible = False
            cmbTopFood.Visible = False
            
            Chart.Visible = True
        Case 5:  'Top Food(전체)
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB5_OFF)
            Set imgSub(5).Picture = LoadPicture(App.Path & PATH04 & IMG_SUB6_ON)
            
            chtNutrition.Visible = False
            grdAniVegRate.Visible = False
            grd19.Visible = False
            grdMealCalory.Visible = False
            Chart.Visible = False
            
            grdMealRank.Visible = True
            cmbTopFood.Visible = True
            Call cmbTopFood_Click
        End Select
End Sub

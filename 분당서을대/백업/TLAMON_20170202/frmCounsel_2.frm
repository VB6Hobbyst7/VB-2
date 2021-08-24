VERSION 5.00
Begin VB.Form frmCounsel_2 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmCounsel_2.frx":0000
   ScaleHeight     =   11445
   ScaleWidth      =   22965
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Label lblAfterEx 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "5"
      Height          =   255
      Left            =   3540
      TabIndex        =   0
      Top             =   4470
      Width           =   525
   End
   Begin VB.Label lblAfterMeal 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "5"
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   4470
      Width           =   525
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2760
      TabIndex        =   24
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   5
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":BAF2
      Top             =   3870
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Index           =   5
      Left            =   4200
      TabIndex        =   23
      Top             =   3840
      Width           =   765
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5220
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6360
      TabIndex        =   21
      Top             =   3840
      Width           =   945
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2760
      TabIndex        =   20
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   4
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":BE57
      Top             =   3480
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Index           =   4
      Left            =   4200
      TabIndex        =   19
      Top             =   3450
      Width           =   765
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5220
      TabIndex        =   18
      Top             =   3450
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   17
      Top             =   3450
      Width           =   945
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   16
      Top             =   3090
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   3
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":C1BC
      Top             =   3120
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   3090
      Width           =   765
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5220
      TabIndex        =   14
      Top             =   3090
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   13
      Top             =   3090
      Width           =   945
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   2700
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   2
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":C521
      Top             =   2730
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Index           =   2
      Left            =   4200
      TabIndex        =   11
      Top             =   2700
      Width           =   765
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5220
      TabIndex        =   10
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   9
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   2310
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   1
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":C886
      Top             =   2340
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   7
      Top             =   2310
      Width           =   765
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5220
      TabIndex        =   6
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   5
      Top             =   2310
      Width           =   945
   End
   Begin VB.Label lblMeta 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   0
      Left            =   4000
      Picture         =   "frmCounsel_2.frx":CBEB
      Top             =   1920
      Width           =   150
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
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
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   1890
      Width           =   885
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5220
      TabIndex        =   2
      Top             =   1890
      Width           =   975
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   1890
      Width           =   945
   End
   Begin VB.Image imgPrint 
      Height          =   990
      Left            =   10890
      Picture         =   "frmCounsel_2.frx":CF50
      Top             =   7170
      Width           =   975
   End
End
Attribute VB_Name = "frmCounsel_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 상하화살표(증권식)
'+---------------------------------------------------------------------------------+
Private Const IMG_UP As String = "\Back\Counsel\01\icon-red.jpg"
Private Const IMG_DOWN As String = "\Back\Counsel\01\icon-blue.jpg"

Private Const IMG_PRINT_ON As String = "\Back\Counsel\02\대사량상담 on.jpg"
Private Const IMG_PRINT_OFF As String = "\Back\Counsel\02\대사량상담 off.jpg"

Public Sub Form_Load()
'    Me.Top = FRM_TOP
'    Me.Left = FRM_LEFT
'    Me.Width = FRM_WIDTH
'    Me.Height = FRM_HEIGHT
'    Me.BackColor = vbWhite
'
'    Call InitialLabel
'    Call ShowMetaRecord
'    Call InitialChart
'    Call DrawChart
    
'    Set imgPrint.Picture = LoadPicture(App.Path & IMG_PRINT_OFF)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ShowMetaRecord()
'RMR/TEE
' 휴식대사량 8 : inRMR / 9 : etcRMR / 그밖에 : RMR
    Dim qrySelect As String, rValue As Variant
    Dim sngGab As Single
    Dim strPer As String
    Dim sngTemp As Single
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT TOP 2 isnull(VO2,0), isnull(inRMR,0), ISNULL(RMR,0), ISNULL(etcRMR,0), AfterMeal, AfterEx "
    qrySelect = qrySelect & "FROM Treat RIGHT JOIN BodyData AS b ON b.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        '1) 현재
        For i = 0 To 3
            lblMeta(i).Caption = Is_Null(rValue(i, 0), "-")
        Next i
        If lblMeta(0).Caption <> "-" Then
            lblMeta(0).Caption = lblMeta(0).Caption & " ml/min"
        End If
        If lblMeta(1).Caption <> "-" Then
            lblMeta(1).Caption = Format(lblMeta(1).Caption, "#,###") & " kcal"
            If rValue(2, 0) <> 0 Then sngTemp = (rValue(1, 0) / rValue(2, 0)) * 100
            lblMeta(4).Caption = Format(sngTemp, "#.#") & " %"
        Else
            lblMeta(4).Caption = "-"
        End If
        If lblMeta(2).Caption <> "-" Then
            lblMeta(2).Caption = Format(lblMeta(2).Caption, "#,###") & " kcal"
        End If
        If lblMeta(3).Caption <> "-" Then
            lblMeta(3).Caption = Format(lblMeta(3).Caption, "#,###") & " kcal"
            If rValue(3, 0) <> 0 Then sngTemp = (rValue(1, 0) / rValue(3, 0)) * 100
            lblMeta(5).Caption = Format(sngTemp, "#.#") & " %"
        Else
            lblMeta(5).Caption = "-"
        End If
        
        If UBound(rValue, 2) > 0 Then
            '2) 전회대비
            If Not IsNull(rValue(0, 0)) And Not IsNull(rValue(0, 1)) Then
                sngGab = rValue(0, 0) - rValue(0, 1)
                Call DrawUpDown(sngGab, 0, " ml/min")
            Else
                Set imgUpDown(0).Picture = LoadPicture("")
                lblUpDown(0).Caption = "-"
            End If
            If Not IsNull(rValue(1, 0)) And Not IsNull(rValue(1, 1)) Then
                sngGab = rValue(1, 0) - rValue(1, 1)
                Call DrawUpDown(sngGab, 1, " kcal")
            Else
                Set imgUpDown(1).Picture = LoadPicture("")
                lblUpDown(1).Caption = "-"
            End If
            If Not IsNull(rValue(2, 0)) And Not IsNull(rValue(2, 1)) Then
                sngGab = rValue(2, 0) - rValue(2, 1)
                Call DrawUpDown(sngGab, 2, " kcal")
            Else
                Set imgUpDown(2).Picture = LoadPicture("")
                lblUpDown(2).Caption = "-"
            End If
            If Not IsNull(rValue(3, 0)) And Not IsNull(rValue(3, 1)) Then
                sngGab = rValue(3, 0) - rValue(3, 1)
                Call DrawUpDown(sngGab, 3, " kcal")
            Else
                Set imgUpDown(3).Picture = LoadPicture("")
                lblUpDown(3).Caption = "-"
            End If
            '3) 최고/최저
            'VO2
            lblMax(0).Caption = MaxValue("VO2") & " ml/min"
            lblMin(0).Caption = MinValue("VO2") & " ml/min"
            '메드젬 RMR
            lblMax(1).Caption = Format(MaxValue("inRMR"), "#,###") & " kcal"
            lblMin(1).Caption = Format(MinValue("inRMR"), "#,###") & " kcal"
            'H-B RMR
            lblMax(2).Caption = Format(MaxValue("RMR"), "#,###") & " kcal"
            lblMin(2).Caption = Format(MinValue("RMR"), "#,###") & " kcal"
            '타장비 RMR
            lblMax(3).Caption = Format(MaxValue("etcRMR"), "#,###") & " kcal"
            lblMin(3).Caption = Format(MinValue("etcRMR"), "#,###") & " kcal"
        End If
        
        '4) 식후시간, 운동후시간
        lblAfterMeal.Caption = Is_Null(rValue(4, 0), "-")
        lblAfterEx.Caption = Is_Null(rValue(5, 0), "-")
    Else
        For i = 0 To 5
            lblMeta(i).Caption = ""
            Set imgUpDown(i).Picture = LoadPicture("")
            lblUpDown(i).Caption = ""
            lblMax(i).Caption = ""
            lblMin(i).Caption = ""
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub DrawUpDown(sngGab As Single, i As Integer, strUnit As String)
    If sngGab < 0 Then       '하향 화살표 파란색 글씨
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_DOWN)
        lblUpDown(i).Caption = sngGab & strUnit
    ElseIf sngGab > 0 Then   '상향 화살표 빨간색 글씨
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_UP)
        lblUpDown(i).Caption = sngGab & strUnit
    Else                     '이전과 같음..변동없음
        Set imgUpDown(i).Picture = LoadPicture("")
        lblUpDown(i).Caption = "---"
    End If
End Sub

Private Sub DrawChart()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer

'   세가지(H-B,MedGem,etc)RMR 모두 한 그래프에 보여주기
    qrySelect = "SELECT TreatDay, RMR, inRMR, etcRMR FROM BodyData LEFT JOIN Treat "
    qrySelect = qrySelect & "ON BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay ASC;"
    
    Set clsSelect = New clsSelect

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        Chart.Visible = True
        Chart.OpenDataEx COD_VALUES, 2, COD_UNKNOWN
        
        Chart.Axis(AXIS_Y).Min = MinRMR - 50
        Chart.Axis(AXIS_Y).Max = MaxRMR + 50

        Chart.Axis(AXIS_Y).STEP = 50
        For i = 0 To UBound(rValue, 2)
            If Not IsNull(rValue(1, i)) Then
                Chart.ValueEx(0, i) = rValue(1, i)
            End If
            If Not IsNull(rValue(2, i)) Then
                Chart.ValueEx(1, i) = rValue(2, i)
            End If
            If Not IsNull(rValue(3, i)) Then
                Chart.ValueEx(2, i) = rValue(3, i)
            End If
            Chart.Axis(AXIS_X).Label(i) = Is_Null(rValue(0, i), "")
        Next i
        Chart.CloseData COD_VALUES
    Else
        MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
        Chart.Visible = False
    End If

    Set clsSelect = Nothing
End Sub

Private Sub DrawChart2()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer

'   세가지(H-B,MedGem,etc)RMR 모두 한 그래프에 보여주기
    qrySelect = "SELECT TreatDay, RMR, inRMR, etcRMR FROM BodyData LEFT JOIN Treat "
    qrySelect = qrySelect & "ON BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay ASC;"
    
    Set clsSelect = New clsSelect

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        Chart.Visible = True
        Chart.OpenDataEx COD_VALUES, 2, COD_UNKNOWN
        
        Chart.Axis(AXIS_Y).Min = MinRMR - 50
        Chart.Axis(AXIS_Y).Max = MaxRMR + 50
        Chart.Axis(AXIS_Y).STEP = 50
        For i = 0 To UBound(rValue, 2)
            If Not IsNull(rValue(1, i)) Then
                Chart.ValueEx(0, i) = rValue(1, i)
            End If
            If Not IsNull(rValue(2, i)) Then
                Chart.ValueEx(1, i) = rValue(2, i)
            End If
            If Not IsNull(rValue(3, i)) Then
                Chart.ValueEx(2, i) = rValue(3, i)
            End If
            Chart.Axis(AXIS_X).Label(i) = Is_Null(rValue(0, i), "")
        Next i
        Chart.CloseData COD_VALUES
    Else
        MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
        Chart.Visible = False
    End If

    Set clsSelect = Nothing

End Sub

Private Function MinValue(strField As String) As Single
    Dim qrySelect As String, rMin As Variant

    If strField = "RMR" Then
        qrySelect = "SELECT MIN(rmr) FROM ( "
        If HowOld >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MIN(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    Set clsSelect = New clsSelect

    rMin = clsSelect.Query(qrySelect)
    If Not IsNull(rMin(0, 0)) Then
        MinValue = CSng(rMin(0, 0))
        Erase rMin
    Else
        MinValue = 0
    End If
    Set clsSelect = Nothing
End Function

Private Function MaxValue(strField As String) As Single
    Dim qrySelect As String, rMax As Variant

    If strField = "RMR" Then
        qrySelect = "SELECT MAX(rmr) FROM ( "
        If HowOld >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MAX(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    Set clsSelect = New clsSelect
    
    rMax = clsSelect.Query(qrySelect)
    If Not IsNull(rMax(0, 0)) Then
        MaxValue = CSng(rMax(0, 0))
        Erase rMax
    Else
        MaxValue = 0
    End If
    
    Set clsSelect = Nothing
End Function

Private Function MaxRMR() As Single
    Dim qrySelect As String, rMax As Variant
    Dim sngTemp As Single, sngTemp2 As Single, sngMax As Single
    
    qrySelect = "SELECT MAX(RMR), MAX(inRMR), MAX(etcRMR) FROM BodyData "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rMax = clsSelect.Query(qrySelect)
    If Not IsNull(rMax(0, 0)) Then
        sngTemp = Is_Null(rMax(0, 0), 0)
        sngTemp2 = Is_Null(rMax(1, 0), 0)
        If sngTemp > sngTemp2 Then
            sngMax = sngTemp
        Else
            sngMax = sngTemp2
        End If
        sngTemp = Is_Null(rMax(2, 0), 0)
        If sngMax > sngTemp Then
            sngMax = sngMax
        Else
            sngMax = sngTemp
        End If
    End If
    MaxRMR = sngMax
    Set clsSelect = Nothing
End Function

Private Function MinRMR() As Single
    Dim qrySelect As String, rMin As Variant
    Dim sngTemp As Single, sngTemp2 As Single, sngMin As Single
    
    qrySelect = "SELECT MIN(RMR), MIN(inRMR), MIN(etcRMR) FROM BodyData "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rMin = clsSelect.Query(qrySelect)
    If Not IsNull(rMin(0, 0)) Then
        sngTemp = Is_Null(rMin(0, 0), 0)
        sngTemp2 = Is_Null(rMin(1, 0), 0)
        If sngTemp > sngTemp2 And sngTemp2 > 0 Then
            sngMin = sngTemp2
        Else
            sngMin = sngTemp
        End If
        sngTemp = Is_Null(rMin(2, 0), 0)
        If sngMin > sngTemp And sngTemp > 0 Then
            sngMin = sngTemp
        Else
            sngMin = sngMin
        End If
    End If
    MinRMR = sngMin
    Set clsSelect = Nothing
End Function

Private Sub InitialLabel()
    Dim i As Integer
    
    For i = 0 To 5
        lblMeta(i).Caption = ""
        imgUpDown(i).Picture = LoadPicture("")
        lblUpDown(i).Caption = ""
        lblMax(i).Caption = ""
        lblMin(i).Caption = ""
    Next i
    lblAfterMeal.Caption = ""
    lblAfterEx.Caption = ""
End Sub

Private Sub InitialChart()
    With Chart
        ' Chart Type Settings
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_NONE
        .Axis(0).Grid = True
        .RGBBk = vbWhite
        .BorderStyle = BORDER_NONE
        .AxesStyle = CAS_NONE
        .Axis(AXIS_Y).Decimals = 0
        .MarkerShape = MK_DIAMOND

        ' Color Settings
        .Border = False

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        
        .ToolBar = False
        .PointLabels = True
        .MultipleColors = False
        .Title(CHART_TOPTIT) = ""
        
        .LeftGap = 0
        .TopGap = 0
        .LeftGap = 10
        .RightGap = 10
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & IMG_PRINT_OFF)
End Sub

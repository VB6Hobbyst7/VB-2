VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCounsel_6 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin ChartfxLibCtl.ChartFX chtWeek 
      Height          =   1935
      Left            =   7380
      TabIndex        =   16
      Top             =   2490
      Width           =   2445
      _cx             =   4313
      _cy             =   3413
      Build           =   20
      TypeMask        =   42467330
      LeftGap         =   9
      RightGap        =   0
      TopGap          =   0
      BottomGap       =   0
      Volume          =   40
      AxesStyle       =   0
      Axis(0).Max     =   1000
      Axis(0).Decimals=   0
      Axis(0).Style   =   14440
      Axis(0).Format  =   1
      Axis(0).Format  =   1
      Axis(2).Style   =   14440
      RGBBk           =   16777215
      nColors         =   16
      Colors          =   "frmCounsel_6.frx":0000
      nPts            =   7
      nSer            =   1
      NumPoint        =   7
      NumSer          =   1
      BorderS         =   0
      _Data_          =   "frmCounsel_6.frx":00A0
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      Left            =   10530
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   2070
      Width           =   1455
   End
   Begin ChartfxLibCtl.ChartFX chtExHistory 
      Height          =   3105
      Left            =   720
      TabIndex        =   0
      Top             =   5160
      Width           =   7875
      _cx             =   13891
      _cy             =   5477
      Build           =   20
      TypeMask        =   42467330
      BorderColor     =   16777217
      Axis(0).Decimals=   0
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      nColors         =   16
      Colors          =   "frmCounsel_6.frx":00F8
      nSer            =   1
      NumSer          =   1
      _Data_          =   "frmCounsel_6.frx":0198
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExHistory 
      Height          =   3105
      Left            =   720
      TabIndex        =   1
      Top             =   5160
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5477
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   285
      Left            =   10650
      TabIndex        =   3
      Top             =   2400
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66256897
      CurrentDate     =   37818
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   10650
      TabIndex        =   4
      Top             =   2700
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66256897
      CurrentDate     =   37818
   End
   Begin ChartfxLibCtl.ChartFX chtExTime 
      Height          =   3105
      Left            =   720
      TabIndex        =   17
      Top             =   5160
      Width           =   7875
      _cx             =   13891
      _cy             =   5477
      Build           =   20
      TypeMask        =   42467330
      BorderColor     =   16777217
      Axis(0).Decimals=   0
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      nColors         =   16
      Colors          =   "frmCounsel_6.frx":01F0
      nSer            =   1
      NumSer          =   1
      _Data_          =   "frmCounsel_6.frx":0290
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   10320
      Picture         =   "frmCounsel_6.frx":02E8
      Top             =   3180
      Width           =   915
   End
   Begin VB.Image imgStart 
      Height          =   300
      Left            =   11370
      Picture         =   "frmCounsel_6.frx":0A0E
      Top             =   3150
      Width           =   765
   End
   Begin VB.Image imgPrint 
      Height          =   1065
      Left            =   10650
      Picture         =   "frmCounsel_6.frx":1335
      Top             =   7290
      Width           =   1065
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1350
      X2              =   4020
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1350
      X2              =   4050
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1350
      X2              =   4080
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label lblBody 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "처음 40 kg -> 현재 38 kg (2kg감소)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   3990
      TabIndex        =   15
      Top             =   4500
      Width           =   2775
   End
   Begin VB.Label lblBody 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "처음 40 kg -> 현재 38 kg (2kg감소)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   4470
      TabIndex        =   14
      Top             =   3870
      Width           =   2300
   End
   Begin VB.Label lblBody 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "처음 40 kg -> 현재 38 kg (2kg감소)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   4470
      TabIndex        =   13
      Top             =   3270
      Width           =   2300
   End
   Begin VB.Label lblBody 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "처음 40 kg -> 현재 38 kg (2kg감소)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4470
      TabIndex        =   12
      Top             =   2640
      Width           =   2300
   End
   Begin VB.Label lblCalory 
      BackStyle       =   0  '투명
      Caption         =   "처방 : 400 kcal / 일"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   11
      Top             =   4470
      Width           =   2500
   End
   Begin VB.Label lblCalory 
      BackStyle       =   0  '투명
      Caption         =   "평균 : 400 kcal / 일"
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   10
      Top             =   4140
      Width           =   2500
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  '투명
      Caption         =   "처방 : 1일 / 주"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   9
      Top             =   3390
      Width           =   2500
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  '투명
      Caption         =   "평균 : 1일 / 주"
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   8
      Top             =   3060
      Width           =   2500
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  '투명
      Caption         =   "처방 : 70분 / 일"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   7
      Top             =   2310
      Width           =   2500
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  '투명
      Caption         =   "평균 : 60분 / 일"
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   6
      Top             =   1980
      Width           =   2500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "~"
      Height          =   195
      Index           =   1
      Left            =   10440
      TabIndex        =   5
      Top             =   2760
      Width           =   165
   End
   Begin VB.Image imgSub 
      Height          =   360
      Index           =   2
      Left            =   8730
      Picture         =   "frmCounsel_6.frx":2A38
      Top             =   6300
      Width           =   1125
   End
   Begin VB.Image imgSub 
      Height          =   360
      Index           =   1
      Left            =   8730
      Picture         =   "frmCounsel_6.frx":33A0
      Top             =   5850
      Width           =   1125
   End
   Begin VB.Image imgSub 
      Height          =   360
      Index           =   0
      Left            =   8730
      Picture         =   "frmCounsel_6.frx":3CDC
      Top             =   5370
      Width           =   1125
   End
End
Attribute VB_Name = "frmCounsel_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PATH06 As String = "\Back\Counsel\06\"
Private Const IMG_SUB1_ON As String = "운동일기평가 on.jpg"
Private Const IMG_SUB1_OFF As String = "운동일기평가 off.jpg"
Private Const IMG_SUB2_ON As String = "소모칼로리 on.jpg"
Private Const IMG_SUB2_OFF As String = "소모칼로리 off.jpg"
Private Const IMG_SUB3_ON As String = "운동시간 on.jpg"
Private Const IMG_SUB3_OFF As String = "운동시간 off.jpg"
Private Const IMG_PRINT_ON As String = "운동일기 출력버튼 on.jpg"
Private Const IMG_PRINT_OFF As String = "운동일기 출력버튼 off.jpg"

Public Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\06\06.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite
    
    
    
    Call RPT_ExDiary
    Call InitialLabel
    Call InitialChtWeek
    Call InitialGridHistory
    Call InitialHistoryChart
    Call InitialTimeChart
    
    cmbPeriod.Clear
    cmbPeriod.AddItem "전체"
    cmbPeriod.AddItem "특정기간"
    cmbPeriod.ListIndex = 0
    
    dtpBegin.Value = Now()
    dtpEnd.Value = Now()
    
    Call imgSub_Click(0)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH06 & IMG_PRINT_OFF)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub cmbPeriod_Click()
    If cmbPeriod.ListIndex = 0 Then    '전체 평가보기
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
        '평가보기
        Call ShowVal
    Else
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    End If
End Sub

Private Sub cmbPeriod_Change()
    Call cmbPeriod_Click
End Sub

Private Sub ShowVal()
    Call TopTime
    Call TopCount
    Call TopCalory
    Call ShowWeekChart
    
    Call ShowHistoryGrid
    Call ShowHistoryChart
    Call ShowTimeChart
    
    Call ShowComposition
End Sub

Private Sub TopTime()
'1일에 평균 몇시간(분)씩 운동을 했는지
    Dim qrySelect As String, rValue As Variant
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT AVG(PlayTime) FROM SportsDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblTime(0).Caption = "평균 : " & rValue(0, 0) & " 분 / 일"
    Else
        lblTime(0).Caption = ""
    End If
    
    qrySelect = "SELECT TreatTime FROM RPT_ExDiary WHERE CustomerNum=" & glngCustomerNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblTime(1).Caption = "처방 : " & Is_Null(rValue(0, 0), 0) & " 분 / 일"
    Else
        lblTime(1).Caption = ""
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopCount()
'1주일에 평균 몇회 운동을 했는지
    Dim qrySelect As String, rValue As Variant
    Dim intGabDay As Integer, sngAvg As Single
    
    Set clsSelect = New clsSelect
    
    '평균값
    qrySelect = "SELECT DISTINCT(PlayDay) FROM SportsDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " ORDER BY PlayDay;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        intGabDay = DateDiff("ww", rValue(0, 0), rValue(0, UBound(rValue, 2)))
        If intGabDay = 0 Then
            sngAvg = (UBound(rValue, 2) + 1)
        Else
            sngAvg = (UBound(rValue, 2) + 1) / (intGabDay + 1)
        End If
        lblCount(0).Caption = "평균 : " & Format(sngAvg, "0.0") & " 회 / 주"
    Else
        lblCount(0).Caption = ""
    End If
    
    '처방값(기간내에 처방한 운동횟수의 평균값 ? )
    qrySelect = "SELECT AVG(ExDay) FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND TreatDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND TreatDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblCount(1).Caption = "처방 : " & Is_Null(rValue(0, 0), 0) & " 회 / 주"
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopCalory()
'1일에 운동으로 평균 몇 칼로리씩 소모했는지
    Dim qrySelect As String, rValue As Variant
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT AVG(BurnCalories) FROM SportsDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblCalory(0).Caption = "평균 : " & Is_Null(rValue(0, 0), 0) & " kcal / 일"
    Else
        lblCalory(0).Caption = ""
    End If
    
    '처방값(기간내에 처방한 운동횟수의 평균값 ? )
    qrySelect = "SELECT AVG(ExCalory) FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND TreatDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND TreatDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblCalory(1).Caption = "처방 : " & Is_Null(rValue(0, 0), 0) & " kcal / 일"
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowComposition()
'체지방량 / 체지방률 / 근육량 / RMR
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer, strUnit(3) As String
    Dim strFirst As String, strLast As String, intGab As Integer
    
    strUnit(0) = " kg"
    strUnit(1) = " %"
    strUnit(2) = " kg"
    strUnit(3) = " kcal"
    
    qrySelect = "SELECT ChFat, ChFatRate, Muscle, "
    If HowOld >= ADULT_AGE Then
        qrySelect = qrySelect & " CASE AdBasicDsa "
    Else
        qrySelect = qrySelect & " CASE BaBasicDsa "
    End If
    qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
    qrySelect = qrySelect & "FROM Treat INNER JOIN BodyData "
    qrySelect = qrySelect & "ON BodyData.TreatNum=Treat.TreatNum INNER JOIN CompData "
    qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND TreatDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND TreatDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"
    
    Set clsSelect = New clsSelect
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
            For j = UBound(rValue, 2) To 0 Step -1
                If i = 1 Then  '체지방률인경우
                    strFirst = Format(Is_Null(rValue(i, j), ""), "#.#")
                Else
                    strFirst = Is_Null(rValue(i, j), "")
                End If
                If strFirst = "." Then strFirst = ""
                If strFirst <> "" Then
                    Exit For
                End If
            Next j
            For j = 0 To UBound(rValue, 2)
                If i = 1 Then
                    strLast = Format(Is_Null(rValue(i, j), ""), "#.#")
                Else
                    strLast = Is_Null(rValue(i, j), "")
                End If
                If strLast = "." Then strLast = ""
                If strLast <> "" Then
                    Exit For
                End If
            Next j
            If strFirst <> "" Then
                lblBody(i).Caption = "처음 " & strFirst & strUnit(i)
                If strLast <> "" Then
                    lblBody(i).Caption = lblBody(i).Caption & "-> 현재 " & strLast & strUnit(i)
                    intGab = CInt(strFirst) - CInt(strLast)
                    If intGab > 0 Then
                        lblBody(i).Caption = lblBody(i).Caption & vbNewLine & "(" & intGab & strUnit(i) & " 감소)"
                    ElseIf intGab < 0 Then
                        lblBody(i).Caption = lblBody(i).Caption & vbNewLine & "(" & Abs(intGab) & strUnit(i) & " 증가)"
                    Else
                        lblBody(i).Caption = lblBody(i).Caption & vbNewLine & "(변화없음)"
                    End If
                End If
            Else
                If strLast <> "" Then
                    lblBody(i).Caption = "현재 " & strLast & strUnit(i)
                Else
                    lblBody(i).Caption = "표시할 데이터 없음"
                End If
            End If
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowWeekChart()
'요일별 운동량
    Dim qrySelect As String, rValue As Variant
    Dim sngCalory(7) As Single, intCount(7) As Integer, nWeek As Integer
    Dim cfxArray As Object
    Dim i As Integer
    
    Set cfxArray = CreateObject("cfxdata.array")
    For i = 0 To 6
        sngCalory(i) = 0
        intCount(i) = 0
    Next i
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT PlayDay, SUM(BurnCalories) FROM SportsDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY PlayDay"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            nWeek = Weekday(rValue(0, i), vbMonday) - 1
            sngCalory(nWeek) = sngCalory(nWeek) + CSng(rValue(1, i))
            intCount(nWeek) = intCount(nWeek) + 1
        Next i
        For i = 0 To 6
            If intCount(i) > 0 Then
                sngCalory(i) = sngCalory(i) / intCount(i)
            End If
        Next i
        cfxArray.AddArray sngCalory
        
        chtWeek.GetExternalData cfxArray
    Else
        chtWeek.ClearData CD_VALUES
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowHistoryGrid()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    
    Call InitialGridHistory
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT PlayDay, ExName, PlayTime, BurnCalories "
    qrySelect = qrySelect & "FROM Sportsdiary INNER JOIN tblExercise ON Sportsdiary.SportsCode=tblExercise.ExNo "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " ORDER BY PlayDay ASC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With grdExHistory
        For i = 0 To UBound(rValue, 2)
            .RowS = .RowS + 1
            .TextMatrix(i + 1, 0) = Format(rValue(0, i), "YYYY-MM-DD") '운동일
            .TextMatrix(i + 1, 1) = Trim(rValue(1, i))                '운동종목
            .TextMatrix(i + 1, 2) = Is_Null(rValue(2, i), 0) & " 분"  '운동시간
            .TextMatrix(i + 1, 3) = Is_Null(rValue(3, i), 0) & " kcal" '소모칼로리
            
            .MergeCol(0) = True
        Next i
        .RowS = .RowS - 1
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .RowHeight(-1) = 300
        End With
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowHistoryChart()
'최대값, 최소값 정해줄 것
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant
    Dim colDate As New Collection, colSum As New Collection
    Dim cfxArray As Object
    Dim intMax As Integer
    
    Set clsSelect = New clsSelect
    Set cfxArray = CreateObject("cfxdata.array")
    
    qrySelect = "SELECT MAX(a) FROM ("
    qrySelect = qrySelect & "SELECT Playday, SUM(BurnCalories) AS a FROM Sportsdiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY Playday) b"
        
    rValue = clsSelect.Query(qrySelect)
    intMax = Int(Is_Null(rValue(0, 0), 900) / 100) * 100
    chtExHistory.Axis(AXIS_Y).Max = intMax + 100
    
    qrySelect = "SELECT Playday, SUM(BurnCalories) FROM Sportsdiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY Playday;"
        
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            colDate.Add rValue(0, i)
            colSum.Add rValue(1, i)
        Next i
        cfxArray.AddArray colDate
        cfxArray.AddArray colSum
        
        chtExHistory.GetExternalData cfxArray
    Else
        chtExHistory.ClearData CD_VALUES
    End If
    
    Set colSum = Nothing
    Set colDate = Nothing
    Set clsSelect = Nothing
End Sub

Private Sub ShowTimeChart()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    Dim colDate As New Collection, colSum As New Collection
    Dim cfxArray As Object
    Dim intMax As Integer
    
    Set clsSelect = New clsSelect
    Set cfxArray = CreateObject("cfxdata.array")
    
    qrySelect = "SELECT MAX(a) FROM ("
    qrySelect = qrySelect & "SELECT Playday, SUM(PlayTime) AS a FROM Sportsdiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY Playday) b"
        
    rValue = clsSelect.Query(qrySelect)
    intMax = Int(Is_Null(rValue(0, 0), 180) / 10) * 10
    chtExTime.Axis(AXIS_Y).Max = intMax + 30
    
    qrySelect = "SELECT PlayDay, SUM(PlayTime) FROM Sportsdiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND PlayDay>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND PlayDay<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY PlayDay;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            colDate.Add rValue(0, i)
            colSum.Add rValue(1, i)
        Next i
        cfxArray.AddArray colDate
        cfxArray.AddArray colSum
        
        chtExTime.GetExternalData cfxArray
    Else
        chtExTime.ClearData CD_VALUES
    End If
    
    Set colSum = Nothing
    Set colDate = Nothing
    Set clsSelect = Nothing
End Sub

Private Sub InitialLabel()
    lblTime(0).Caption = ""
    lblTime(1).Caption = ""
        
    lblCount(0).Caption = ""
    lblCount(1).Caption = ""
    
    lblCalory(0).Caption = ""
    lblCalory(1).Caption = ""
    
    lblBody(0).Caption = ""
    lblBody(1).Caption = ""
    lblBody(2).Caption = ""
    lblBody(3).Caption = ""
End Sub

Private Sub InitialHistoryChart()
    With chtExHistory
        .Gallery = BAR
        .Chart3D = False
        .AxesStyle = CAS_FLATFRAME
        .Stacked = CHART_NOSTACKED
        .Axis(2).Grid = False
        .Axis(0).Grid = True
        .Axis(AXIS_Y).Title = "소모칼로리(kcal)"
        .Volume = 30
        
        .Border = True
        .RGBBk = vbWhite
        .ToolBar = False

        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""
        .PointLabels = True
        .Axis(AXIS_Y).STEP = 50
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialTimeChart()
    With chtExTime
        .Gallery = BAR
        .Chart3D = False
        .AxesStyle = CAS_FLATFRAME
        .Axis(AXIS_X).Grid = False
        .Axis(AXIS_Y).Grid = True
        .Axis(AXIS_Y).Max = 240
        .Axis(AXIS_Y).Title = "운동시간(분)"
        .Volume = 30
        
        .Border = True
        .RGBBk = vbWhite
        .ToolBar = False
        
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

Private Sub InitialChtWeek()
    With chtWeek
        .Gallery = BAR
        .Chart3D = False
        .Axis(AXIS_X).Grid = False
        .Axis(AXIS_Y).Grid = False
        
        .RGBBk = &HEFEFEF
        .Border = True
        .BorderStyle = BORDER_NONE
        .MultipleColors = False
        .Axis(AXIS_Y).Min = 0
        .Axis(AXIS_Y).Max = 830
        .PointLabels = True
        
        .LegendBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""
        .TopGap = 0
        .BottomGap = 0
        .LeftGap = 0
        .RightGap = 0
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialGridHistory()
    With grdExHistory
        .Clear
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .RowS = 2
        .ColS = 4
        .FixedCols = 0
        .FixedRows = 1
        .BackColorBkg = vbWhite
        .ScrollBars = flexScrollBarVertical
        
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
        .TextMatrix(0, 0) = "운동일"
        .TextMatrix(0, 1) = "운동종목"
        .TextMatrix(0, 2) = "시간"
        .TextMatrix(0, 3) = "소모열량"
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1100
       
        '그리드의 선 색깔..
        .GridColor = FRM_GRAY
        .GridLineWidth = 2
        
        .MergeCells = flexMergeRestrictColumns
    End With
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH06 & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH06 & IMG_PRINT_OFF)
    '운동일기 평가 출력
    
End Sub

Private Sub imgStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStart.Picture = LoadPicture(App.Path & "\Back\Counsel\on.jpg")
End Sub

Private Sub imgStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStart.Picture = LoadPicture(App.Path & "\Back\Counsel\off.jpg")
    Call ShowVal
End Sub

Private Sub imgSub_Click(Index As Integer)
    Select Case Index
        Case 0:
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB1_ON)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB3_OFF)
        
            chtExHistory.Visible = False
            chtExTime.Visible = False
            grdExHistory.Visible = True
        Case 1:
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB2_ON)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB3_OFF)
            
            grdExHistory.Visible = False
            chtExTime.Visible = False
            chtExHistory.Visible = True
        Case 2:
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH06 & IMG_SUB3_ON)
            
            grdExHistory.Visible = False
            chtExHistory.Visible = False
            chtExTime.Visible = True
    End Select
End Sub

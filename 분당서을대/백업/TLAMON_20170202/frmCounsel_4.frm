VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCounsel_4 
   BorderStyle     =   0  '����
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
      Style           =   2  '��Ӵٿ� ���
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
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   9
      Top             =   2190
      Width           =   1515
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      Left            =   10500
      Style           =   2  '��Ӵٿ� ���
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
      BackStyle       =   0  '����
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
      BackStyle       =   0  '����
      Caption         =   "���� > ���İ�Ƽ > ������ > �ҹ� > ���߱�ġ"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� > ���İ�Ƽ > ������ > �ҹ� > ���߱�ġ"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� > ���İ�Ƽ > ������ > �ҹ� > ���߱�ġ"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� > ���İ�Ƽ > ������ > �ҹ� > ���߱�ġ"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� > ���İ�Ƽ > ������ > �ҹ� > ���߱�ġ"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,200"
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,200"
      Height          =   225
      Index           =   1
      Left            =   3090
      TabIndex        =   18
      Top             =   2970
      Width           =   915
   End
   Begin VB.Label lblRecommend 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,200"
      Height          =   225
      Index           =   0
      Left            =   3090
      TabIndex        =   17
      Top             =   2490
      Width           =   915
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "4,801"
      Height          =   195
      Index           =   4
      Left            =   2190
      TabIndex        =   16
      Top             =   4350
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "5,230"
      Height          =   195
      Index           =   3
      Left            =   2190
      TabIndex        =   15
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "10.9g(24%)"
      Height          =   195
      Index           =   2
      Left            =   2190
      TabIndex        =   14
      Top             =   3450
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65:20:15"
      Height          =   195
      Index           =   1
      Left            =   2190
      TabIndex        =   13
      Top             =   2970
      Width           =   885
   End
   Begin VB.Label lblIntake 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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

Private Const IMG_PRINT_ON As String = "���-������ ��� on.jpg"
Private Const IMG_PRINT_OFF As String = "���-������ ��� off.jpg"

Dim crxApplication As New CRAXDRT.Application
Public crxReport As CRAXDRT.Report
Public crxReport2 As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxFormula As CRAXDRT.FormulaFieldDefinition
Dim strServer As String, strDBName As String, strUID As String, strPWD As String

Public Sub Form_Load()
    Dim i As Integer
'���� �ߴ� ��ġ �� �׷��� ==================================================
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\04\04.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite

    If ExistDiary = False Then
        MsgBox "�Է��� �Ļ��ϱⰡ �����ϴ�. ", vbOKOnly + vbExclamation
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
'�Է��� �ϱ� �߿� ���� ������ �����򰡸� ������
'���� ���� ���õ� �ϱⰡ �ִٸ� �װ��� ������- ���� ����ġ�� �����򰡸� �����ְ� �ִ��� ������ ��
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
'�ش�ȯ�ڿ� �Է��� �Ļ��ϱⰡ �ִ��� üũ
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
    cmbPeriod.AddItem "Ư����"
    cmbPeriod.AddItem "Ư���Ⱓ"
    cmbPeriod.AddItem "��ü"
    cmbPeriod.ListIndex = 0
    
    For i = 0 To 4
        lblIntake(i).Caption = ""
        lblRecommend(i).Caption = ""
        lblTopFood(i).Caption = ""
    Next i

'Į�θ�,�ܹ���,��Ÿ��A,��Ÿ��E,��Ÿ��C,��Ÿ��B1,��Ÿ��B2,���̾ƽ�,��Ÿ��B6,����,Į��,��,ö,�ƿ�
    With cmbTopFood
        .Clear
        .AddItem "Į�θ�"
        .AddItem "ź��ȭ��"
        .AddItem "�ܹ���"
        .AddItem "����"
        .AddItem "��Ÿ��A"
        .AddItem "��Ÿ��E"
        .AddItem "��Ÿ��C"
        .AddItem "��Ÿ��B1"
        .AddItem "��Ÿ��B2"
        .AddItem "���̾ƽ�"
        .AddItem "��Ÿ��B6"
        .AddItem "����"
        .AddItem "Į��"
        .AddItem "��Ʈ��"
        .AddItem "��"
        .AddItem "ö"
        .AddItem "�ƿ�"
        .ListIndex = 0
    End With
End Sub

Private Sub cmbPeriod_Change()
    Call cmbPeriod_Click
End Sub

Private Sub cmbPeriod_Click()
    Select Case cmbPeriod.ListIndex
        Case 0:   'Ư����
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            cmbDaily.Visible = True
            Call InitialDailyCombo
            Call imgStart_Click
       Case 1:   'Ư���Ⱓ
            cmbDaily.Visible = False
            dtpBegin.Visible = True
            dtpEnd.Visible = True
        Case 2:   '��ü�Ⱓ
            cmbDaily.Visible = False
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            Call imgStart_Click
    End Select
End Sub

Private Sub cmbDaily_Change()
    'Call cmbDaily_Click
End Sub

'<< �Ļ��ϱ� �� >> �������� ����ϱ� ���� �غ��ϴ� �Լ� /////////////////////////////////////////
Private Sub PrintData()
    Dim strConString As String
    Dim qrySelect As String, rValue As Variant
    Dim strBeginDay As String, strEndDay As String
    Dim i As Integer

On Error GoTo PrintErr
    '������� ���õ� �Ⱓ���� ����� ������ �ִ����� ���� Ȯ���� ��
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
        MsgBox "�Ⱓ���� �Էµ� �Ļ��ϱⰡ �����ϴ�.", vbInformation
        Set clsSelect = Nothing
        Exit Sub
    End If
    Set clsSelect = Nothing
    '������ ���� ����
    strServer = ServerName
'2005-01-18 ������ DB��������
    strDBName = DBinfo.DBName
    strUID = DBinfo.DBID
    strPWD = DBinfo.DBPWD
'    strDBName = "Body"
'    strUID = "sa"
'    strPWD = "1111"


    Set crxReport = crxApplication.OpenReport(App.Path & "\Report\�Ļ��ϱ���.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport
        .RecordSelectionFormula = "{CustomerInfo.CustomerNum}=" & glngCustomerNum

        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
'//////////////////////////////////////////  RDC ��ĺ���
        '+--------------------------------------------------
        '+ 1) ����� �������
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
        '    - ���뷮, ���� ���� ���Ե� �ټ����� ����
        '    - ���� / �����淮 / ��ȭ���� / ��ȭ,����ȭ / �ݷ����׷� / ��Ʈ��
        .FormulaFields(3).Text = "'" & RPT_TopFood("����") & "'"
        .FormulaFields(4).Text = "'" & RPT_TopFood("����") & "'"
        .FormulaFields(5).Text = "'" & RPT_TopFood("��ȭ����") & "'"
        .FormulaFields(6).Text = "'" & RPT_TopFood("�ݷ����׷�") & "'"
        .FormulaFields(7).Text = "'" & RPT_TopFood("��Ʈ��") & "'"
'        '    - ���� ���差(���õ� �Ⱓ�� ó��� Į�θ����� ��հ�)
        .FormulaFields(8).Text = "'" & Format(WhatTreatCalory, "#,###") & "'"
        '    - ���õ� �Ⱓ �ѷ���
        If cmbPeriod.ListIndex = 0 Then   'Ư����
            .FormulaFields(9).Text = "'" & Format(cmbDaily.Text, "YYYY.M.D") & "'"
        ElseIf cmbPeriod.ListIndex = 1 Then
            If dtpBegin.Value = dtpEnd.Value Then   'Ư���Ⱓ ���ý�
                .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YYYY.M.D") & "'"
            Else
                .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YY.M.D") & " ~ " & Format(dtpEnd.Value, "YY.M.D") & "'"
            End If
        ElseIf cmbPeriod.ListIndex = 2 Then         '��ü ���ý�
            '�ʱ� �湮�Ϻ��� ~ ?
            Set clsSelect = New clsSelect

            qrySelect = "SELECT MIN(MealDate) FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum

            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                .FormulaFields(9).Text = "'" & Format(rValue(0, 0), "YYYY.M.D") & " ~'"
            End If
        End If

        '+--------------------------------------------------
        '+ 2) �Ľ���
        '+--------------------------------------------------
        '    [1] ��� �Ϻ� �Ļ� Ƚ��
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
        '    [2] �Ļ���� / �ð�
        '11 : @��Ҿ�ħ
        '12 : @�������
        '13 : @�������
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
                    Case 1    ' ��ħ
                        .FormulaFields(11).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 2    ' ����
                        .FormulaFields(12).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 3    ' ����
                        .FormulaFields(13).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                End Select
            Next i
        End If
        '    [3] �ɸ��ð�
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
        '    [4] ���
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
        '    [5] �� �� ����� ����
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
        '    [6] �ܽ�Ƚ��
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
    '+ �ι�° ��
    '+--------------------------------------------------
    Dim strTemp As String, strBeginDay1 As String, strEndDay1 As String
    Set crxReport2 = crxApplication.OpenReport(App.Path & "\Report\�Ļ��ϱ���2.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport2
        '1 : @GabCalory
        '2 : @GabMent
        '3 : @Rice
        '4 : @Exercise
        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
        '    [1] ����� �Ⱓ ����
        If cmbPeriod.ListIndex = 1 Then
            strBeginDay1 = Left(strBeginDay, 4) & "," & Mid(strBeginDay, 5, 2) & "," & Right(strBeginDay, 2)
            strEndDay1 = Left(strEndDay, 4) & "," & Mid(strEndDay, 5, 2) & "," & Right(strEndDay, 2)
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum & " AND {DietDiary.MealDate} IN DateTime (" & strBeginDay1 & ") TO DateTime (" & strEndDay1 & ")"
        Else
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum
        End If
        .RecordSelectionFormula = strTemp

        '    [2] �ϴܿ� �������
        Dim sngAvgTreatCal As Single, sngAvgMealCal As Single
        Dim sngAvgWeight As Single
        Dim sngGabCal As Single, sngRice As Single, intExercise As Integer
        '        - �ش�Ⱓ���� ó��� Į�θ�(Treat.TreatCalory)�� ��հ�
        sngAvgTreatCal = WhatTreatCalory
        If sngAvgTreatCal <> 0 Then
        '        - �ش�Ⱓ���� ���� ����(�Ϻ�)(DietDiary)�� ��հ�
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
                    .FormulaFields(2).Text = "'����'"
                ElseIf sngGabCal < 0 Then
                    .FormulaFields(2).Text = "'�ʰ�'"
                Else
                    .FormulaFields(2).Text = "'���差'"
                End If
            '    - �� �Ѱ��� 300kcal
                sngRice = Abs(sngGabCal) / 300
                If sngRice >= 0.6 Then
                    .FormulaFields(3).Text = "'" & Format(sngRice, "#") & "'"
                ElseIf sngRice < 0.6 And sngRice >= 0.4 Then
                    .FormulaFields(3).Text = "'��'"
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
                    .FormulaFields(4).Text = "'" & intExercise & " ��'"
                Else  '�Ⱓ�� �Էµ� ü���� ���ٸ� ���� �ֱ� ü��
                    qrySelect = "SELECT TOP 1 Weight FROM BodyData LEFT JOIN Treat "
                    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
                    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
                    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

                    rValue = clsSelect.Query(qrySelect)
                    If Not IsNull(rValue(0, 0)) Then
                        sngAvgWeight = CSng(rValue(0, 0))
                        intExercise = sngGabCal / (sngAvgWeight * 0.16)
                        .FormulaFields(4).Text = "'" & intExercise & " ��'"
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

    MsgBox "����� �Ϸ�Ǿ����ϴ�.", vbOKOnly + vbInformation, "���"
    Set clsSelect = Nothing

    Exit Sub
PrintErr:
    '2004-12-23 ������ �αױ��
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
        '���� �Ⱓ���� �Էµ� ó��Į�θ��� ���ٸ� ���� �ֱ� ������ ����Ѵ�.
        qrySelect = "SELECT TOP 1 TreatCalory FROM Treat "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            WhatTreatCalory = rValue(0, 0)
        Else
            WhatTreatCalory = "0"   '�� �ܰ���� ���� �ȵ�
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
        Case "����"
            strFldNutrition = "tblFood.Energy"
        Case "����"
            strFldNutrition = "tblFood.Fat"
        Case "��ȭ����"
            strFldNutrition = "tblFood.SFA"
        Case "�ݷ����׷�"
            strFldNutrition = "tblFood.Cholesterol"
        Case "��Ʈ��"
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
                '���ĸ��� 7���̻��̸� ...���� ����ǥ��
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

'<7> ���� - 19���� ��ǰ���� ����Ƚ�� //////////////////////////////////////////////////////////
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
        .TextMatrix(0, 1) = "��� ��" & vbNewLine & "�� ��ǰ"
        .TextMatrix(0, 2) = "���� ��" & vbNewLine & "���з�"
        .TextMatrix(0, 3) = "��� ��" & vbNewLine & "�� ��ǰ"
        .TextMatrix(0, 4) = "�η� ��" & vbNewLine & "�� ��ǰ"
        .TextMatrix(0, 5) = "���Ƿ� ��" & vbNewLine & "�� ��ǰ"
        .TextMatrix(0, 6) = "ä�ҷ�"
        .TextMatrix(0, 7) = "������"
        .TextMatrix(0, 8) = "���Ƿ�"
        .TextMatrix(0, 9) = "���� ��" & vbNewLine & "�� ��ǰ"
        .TextMatrix(0, 10) = "����"
        .TextMatrix(0, 11) = "���з�"
        .TextMatrix(0, 12) = "������"
        .TextMatrix(0, 13) = "������" & vbNewLine & "�� ����ǰ"
        .TextMatrix(0, 14) = "������"
        .TextMatrix(0, 15) = "���� ��" & vbNewLine & "�ַ�"
        .TextMatrix(0, 16) = "���̷�"
        .TextMatrix(0, 17) = "��������" & vbNewLine & "��ǰ��"
        .TextMatrix(0, 18) = "�����ķ�"
        .TextMatrix(0, 19) = "��Ÿ"
    End With
End Sub

'<6> ���� - ������/�Ĺ��� ������� /////////////////////////////////////////////////////////////
Private Sub AniVegRate()
'���ӼҺ� ��,�Ĺ��� �������
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
        .TextMatrix(0, 1) = "�ܹ���"
        .TextMatrix(0, 2) = "�ܹ���"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "Į��"
        .TextMatrix(0, 6) = "Į��"
        .TextMatrix(0, 7) = "ö��"
        .TextMatrix(0, 8) = "ö��"

        For i = 1 To 8 Step 2
            .TextMatrix(1, i) = "������" & vbNewLine & "g(%)"
            .TextMatrix(1, i + 1) = "�Ĺ���" & vbNewLine & "g(%)"
        Next i
    End With
End Sub

'<5> ���� - TOP Food ////////////////////////////////////////////////////////////////
Private Sub TopFood(strNutrition As String)
    Dim qrySelect As String, rValue As Variant
    Dim intMealCal As Single, strFldNutrition As String
    Dim i As Integer

    Set clsSelect = New clsSelect
    Call InitialMealRank
    Select Case strNutrition
        Case "Į�θ�"
            strFldNutrition = "tblFood.Energy"
            grdMealRank.TextMatrix(0, 3) = "kcal"
        Case "ź��ȭ��"
            strFldNutrition = "tblFood.Carbohy"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "�ܹ���"
            strFldNutrition = "tblFood.Protein"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "����"
            strFldNutrition = "tblFood.Fat"
            grdMealRank.TextMatrix(0, 3) = "g"
        Case "��Ÿ��A"
            strFldNutrition = "tblFood.Vitamine_A"
            grdMealRank.TextMatrix(0, 3) = "R.E"
        Case "��Ÿ��C"
            strFldNutrition = "tblFood.Vitamine_C"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��Ÿ��B1"
            strFldNutrition = "tblFood.Vitamine_B1"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��Ÿ��B2"
            strFldNutrition = "tblFood.Vitamine_B2"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��Ÿ��E"
            strFldNutrition = "tblFood.Vitamine_E"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "���̾ƽ�"
            strFldNutrition = "tblFood.Niacin"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��Ÿ��B6"
            strFldNutrition = "tblFood.Vitamine_B6"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "����"
            strFldNutrition = "tblFood.Folic"
            grdMealRank.TextMatrix(0, 3) = "ug"
        Case "Į��"
            strFldNutrition = "tblFood.Ca"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��"
            strFldNutrition = "tblFood.P"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "ö"
            strFldNutrition = "tblFood.Fe"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "�ƿ�"
            strFldNutrition = "tblFood.Zn"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��Ʈ��"
            strFldNutrition = "tblFood.Na"
            grdMealRank.TextMatrix(0, 3) = "mg"
        '****** ����� �׸� �߰�
        Case "�ݷ����׷�"
            strFldNutrition = "tblFood.Cholesterol"
            grdMealRank.TextMatrix(0, 3) = "mg"
        Case "��ȭ����"
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

'<4> ����Һ� ���差 ���////////////////////////////////////////////////////////////////
Private Sub NutritionCompare()
    Dim qrySelect As String
    Dim rValue As Variant, rValue2 As Variant
    Dim intDayCount As Integer


On Error GoTo ShowErr
    Call InitialChartNutrition
    '���������� ���̾�Ʈ����, ��ü����, ����, �������� �ҷ���

    With typCustomerInfo
            Set clsSelect = New clsSelect
            qrySelect = "SELECT m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,"
            qrySelect = qrySelect & "m11,m12,m13,m14,m15,m16,m17,m18,m19,m20,m21,m22 "
            qrySelect = qrySelect & "FROM Nutrion WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND bt=6;"

            rValue = clsSelect.Query(qrySelect)
            '������ �Ⱓ������ ����Ұ��� ��� ������ ���̹Ƿ�
            '�Ⱓ�� �Էµ� �ϼ���ŭ ������ �Ϻ� ����� ���ġ�� ���Ѵ�.
            If Not IsNull(rValue) Then
'Į�θ�,�ܹ���,��Ÿ��A,��Ÿ��E,��Ÿ��C,��Ÿ��B1,��Ÿ��B2,���̾ƽ�,��Ÿ��B6,����,Į��,��,ö,�ƿ�
'1,2,13,22,20,16,17,19,18,21,7,8,9,12
                With chtNutrition
                    .OpenDataEx COD_VALUES, 1, 1
                    .Axis(AXIS_X).Label(0) = "Į�θ�"
                    .Axis(AXIS_X).Label(1) = "�ܹ���"
                    .Axis(AXIS_X).Label(2) = "VitA"
                    .Axis(AXIS_X).Label(3) = "VitE"
                    .Axis(AXIS_X).Label(4) = "VitC"
                    .Axis(AXIS_X).Label(5) = "VitB1"
                    .Axis(AXIS_X).Label(6) = "VitB2"
                    .Axis(AXIS_X).Label(7) = "���̾ƽ�"
                    .Axis(AXIS_X).Label(8) = "VitB6"
                    .Axis(AXIS_X).Label(9) = "����"
                    .Axis(AXIS_X).Label(10) = "Į��"
                    .Axis(AXIS_X).Label(11) = "��"
                    .Axis(AXIS_X).Label(12) = "ö"
                    .Axis(AXIS_X).Label(13) = "�ƿ�"
                    
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
    '2004-12-23 ������ �αױ��
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

'<3> ���Ϻ� ��/////////////////////////////////////////////////////////////////////////
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
        strMeal(0) = "��ħ"
        strMeal(1) = "����"
        strMeal(2) = "����"
        strMeal(3) = "����"
    End If

    If intTotal = 0 Then
        Exit Sub
    End If
    
    Set cfxArray = CreateObject("cfxData.Array")

    'íƮ �����ֱ�
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
    '2004-12-23 ������ �αױ��
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
' ����� ���� ���� =============///
' 0 : ����
' 1 : ź��ȭ��:�ܹ���:����
' 2 : �����淮(����)
' 3 : �ݷ����׷�
' 4 : ��Ʈ��
    Dim qrySelect As String, rValue As Variant
    Dim sngTotal As Single, sngC As Single, sngP As Single, sngF As Single
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT m1, m2, m3, m4, m23, m10 FROM Nutrion "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND bt=5;"
    
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
    '1) ���뷮
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
    
    '2) ���差
    lblRecommend(0).Caption = Format(WhatTreatCalory, "#,###")
    lblRecommend(1).Caption = "65:15:20"
    lblRecommend(2).Caption = "�ѿ�����15~20%"
    lblRecommend(3).Caption = "300mg����"
    lblRecommend(4).Caption = "2,400mg" & vbNewLine & "����"
    
    '3) Top Food
    lblTopFood(0).Caption = RPT_TopFood("����")
    lblTopFood(1).Caption = ""
    lblTopFood(2).Caption = RPT_TopFood("����")
    lblTopFood(3).Caption = RPT_TopFood("�ݷ����׷�")
    lblTopFood(4).Caption = RPT_TopFood("��Ʈ��")
End Sub

Private Sub ShowValuation()
    '���������� ���̾�Ʈ����, ��ü����, ����, �������� �ҷ���
   Call LoadCustomerInfo(glngCustomerNum)
    If typCustomerInfo.sngDietCal = 0 Then
        Exit Sub
    End If

    With typCustomerInfo
        If Calculate_Nut(.sngDietCal, .intState, .intAge, .strSex) = True Then
            '���Ϻ���
            Call MealSectionRate
            '1) ����Һ� ���差���
            Call NutritionCompare
            '2) ���Ĺ��� �������
            Call AniVegRate
            '3) 19���� ��ǰ���� ����Ƚ��
            Call NutrionGroup
            '4) ��������(ǥ)
            Call LoadMealCalory
            '5) ����Į�θ�(�׷���)
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
    '�Ļ��ϱ� �Է��� ���� ��� ��Į�θ��� ������
    qrySelect = "SELECT MealDate, SUM(MealCalory) FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "'"
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate ORDER BY MealDate ASC;"
    
    Chart.Title(CHART_TOPTIT) = "����Į�θ� ��ȭ"

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
        MsgBox "ǥ���� �Էµ����Ͱ� �����ϴ�.", vbExclamation, "��ȭ�� �׷���"
       Chart.Visible = False
    End If

    Set clsSelect = Nothing
End Sub

Private Sub LoadMealCalory()
'���� �ش�ȯ���� ����� �Ļ��ϱ���� ���õ� �Ⱓ���� �ش��ϴ� ���� �ҷ��� �����ش�
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
        MsgBox "�Ⱓ���� �Էµ� �Ļ��ϱⰡ �����ϴ�.", vbInformation

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
                .TextMatrix(i, 0) = Format(strMealDate, "M��D��")
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
            '�Ϸ�ġ�� �Ѱ� �����ֱ�
            .TextMatrix(.Rows - 1, 0) = "����"
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

        .TextMatrix(0, 0) = "�Ļ���"

        For i = 1 To 3
            .TextMatrix(0, i) = "��ħ"
        Next i
        For i = 4 To 6
            .TextMatrix(0, i) = "����"
        Next i
        For i = 7 To 9
            .TextMatrix(0, i) = "����"
        Next i
        For i = 10 To 12
            .TextMatrix(0, i) = "����"
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

        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "���ĸ�"
        .TextMatrix(0, 2) = "�κ�"
        .TextMatrix(0, 3) = "����"   '������ �׸����� �ٲ�
    End With
End Sub

'�����򰡸� ���� ���̺�
Private Function Calculate_Nut(sngDietCal As Single, intState As Integer, intAge As Integer, strSex As String) As Boolean
    Dim rValue As Variant
    Dim rValue2 As Variant
    Dim qrySelect As String

    '   s(������, �����ڵ� 1=��ħ   4=����, 0=��)
    Dim s(1 To 36, 0 To 4) As Single, temp As Single
    '   �� ���ϸ� �ش�Ⱓ���� ���� Ƚ�� 1:��ħ~4:����,0�� �Ⱓ�� �Է��� �ϱ��
    '   �ش�Ⱓ���� ����� ���� ����..
    Dim intSectionCnt(0 To 4) As Integer

    ' �� �����, ���Ϻ� count
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
    strTime(1) = "��ħ"
    strTime(2) = "����"
    strTime(3) = "����"
    strTime(4) = "����"
    strTime(5) = "1�� �հ�"
    strTime(6) = "���差���%"

    strEelem(1) = "����"
    strEelem(2) = "�ܹ���"
    strEelem(3) = "����"
    strEelem(4) = "ź��ȭ��"

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
            'k is ����
            k = rValue(29, i)
            For j = 1 To 28
                If Not IsNull(rValue(j, i)) Then
                    '��� ����Ҵ� 100g�� ���� ����
                    temp = rValue(j, i) * rValue(31, i) / 100
                    s(j, k) = s(j, k) + temp
                    s(j, 0) = s(j, 0) + temp
                Else
                    temp = 0
                End If

                '������/�Ĺ��� �������
                If j = 2 Or j = 3 Or j = 7 Or j = 9 Then
                    Select Case j
                       Case 2        '�ܹ���
                            l = 29
                       Case 3        '����
                            l = 31
                       Case 7        'Į��
                            l = 33
                       Case 9        'ö��
                            l = 35
                    End Select

                    If rValue(33, i) = "1" Then      '�Ĺ���
                        If Not IsNull(rValue(j, i)) Then
                            s(l, k) = s(l, k) + temp
                            s(l, 0) = s(l, 0) + temp
                        End If
                    ElseIf rValue(33, i) = "2" Then  '������
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
        '����� ���� : ��,�Ĺ������������ 19���� ��ǰ���� ��վȳ�
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

        '�� ������� �߷����� ���Ѵ�.(�� ���Ϻ�)
        For i = 2 To 4 '2=�ܹ���, 3=���� 4=ź��ȭ��
            ss(1) = ss(1) + s(i, 1) * fac(i) '��ħ
            ss(2) = ss(2) + s(i, 2) * fac(i) '����
            ss(3) = ss(3) + s(i, 3) * fac(i) '����
            ss(4) = ss(4) + s(i, 4) * fac(i) '����
        Next i
        For i = 1 To 4 '�� �����
            If i = 1 Then '������ �� ���� ������
                For j = 1 To 4 '�� ����

                     ssr(j) = Round(s(i, j) / s(i, 0) * 100, 2)
                Next j
            Else '����Ҵ� �ѳ��Ͽ��� �� ������� ���� ������.
                For j = 1 To 4 '�� ����
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
    '�Ⱓ�� �򰡺���
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
        MsgBox "�Ⱓ���� �Էµ� �Ļ��ϱⰡ �����ϴ�.", vbInformation
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
        Case 0:  '����Һ� ���差���
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
        Case 1:  '���Ĺ����������
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
        Case 2:  '19������ǰ��������Ƚ��
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
        Case 3:  '��������(ǥ)
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
        Case 4:  '����Į�θ�(�׷���)
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
        Case 5:  'Top Food(��ü)
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

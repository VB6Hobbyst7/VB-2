VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCounsel_8 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdBadList 
      Height          =   4245
      Left            =   810
      TabIndex        =   0
      Top             =   4020
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin FPSpread.vaSpread sprDisorder 
      Height          =   7695
      Left            =   11160
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   6315
      _Version        =   196608
      _ExtentX        =   11139
      _ExtentY        =   13573
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCounsel_8.frx":0000
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_8.frx":0215
      Top             =   50
      Width           =   13140
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   24
      Left            =   9360
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   23
      Left            =   9240
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   22
      Left            =   9120
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   21
      Left            =   9000
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   20
      Left            =   8880
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   19
      Left            =   8760
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   18
      Left            =   8640
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   17
      Left            =   8520
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   16
      Left            =   8400
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   15
      Left            =   8280
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   14
      Left            =   8160
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   13
      Left            =   8040
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   12
      Left            =   7920
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   11
      Left            =   7800
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   10
      Left            =   7680
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   9
      Left            =   7560
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   8
      Left            =   7440
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   7
      Left            =   7320
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   6
      Left            =   7170
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   5
      Left            =   7080
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   4
      Left            =   6960
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   3
      Left            =   6840
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   2
      Left            =   6720
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   1
      Left            =   6600
      Top             =   2580
      Width           =   135
   End
   Begin VB.Shape shpPoint 
      FillColor       =   &H000000C0&
      FillStyle       =   0  '�ܻ�
      Height          =   135
      Index           =   0
      Left            =   6480
      Top             =   2580
      Width           =   135
   End
   Begin VB.Label lblTotalMent 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1050
      TabIndex        =   2
      Top             =   2040
      Width           =   4515
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmCounsel_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IMG_MALE As String = "\Back\Counsel\08_M.jpg"
Private Const IMG_FEMALE As String = "\Back\Counsel\08_F.jpg"

Public Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & IMG_MALE)
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Height = FRM_HEIGHT
    Me.Width = FRM_WIDTH
    Me.BackColor = vbWhite
    
    Call InitialQuestion
    
    If WhatSex = "M" Then
        Set Me.Picture = LoadPicture(App.Path & IMG_MALE)
    Else
        Set Me.Picture = LoadPicture(App.Path & IMG_FEMALE)
    End If
    Call LoadTest
End Sub

Private Sub LoadTest()
    Dim qrySelect As String, rValue As Variant
    Dim strAnswers As String, intTotal As Integer
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT Answers, TestResult FROM Disorder "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        strAnswers = Trim(rValue(0, 0))
        intTotal = CInt(rValue(1, 0))
        
        Call ShowResult(intTotal, strAnswers)
        Call DrawPoint(intTotal)
        grdBadList.Visible = True
    Else
        MsgBox WhatName & "���� �̻�Ļ��ൿ�� �˻縦 ���� �����̽��ϴ�.", vbOKOnly + vbExclamation
        Me.BackColor = vbWhite
        Set Me.Picture = LoadPicture("")
        grdBadList.Visible = False
        For i = 0 To 24
            shpPoint(i).Visible = False
        Next i
        Exit Sub
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowResult(intScore As Integer, strAnswers As String)
    Dim strSex As String
    Dim intSelNum As Integer
    Dim i As Integer, strAnswer As String, j As Integer
    Dim qrySelect As String, rValue As Variant
    Dim strLastAs As String, strLastA As String

'��Ŭ���� ����� ���� �ӽù���
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT Sex FROM CustomerInfo WHERE CustomerNum=" & glngCustomerNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        strSex = Trim(rValue(0, 0))
    End If
    
    '���� �����ֱ�
'    lblTotal.Caption = intScore & " ��"
    lblTotal.Caption = intScore
    
    If strSex = "M" Then
        Select Case intScore
            Case Is < 15: intSelNum = 1
            Case 15 To 18: intSelNum = 2
            Case 19 To 22: intSelNum = 3
            Case Is >= 23: intSelNum = 4
        End Select
    Else
        Select Case intScore
            Case Is < 18: intSelNum = 1
            Case 18 To 21: intSelNum = 2
            Case 22 To 26: intSelNum = 3
            Case Is >= 27: intSelNum = 4
        End Select
    End If
    '������ �ش��ϴ� �� �����ֱ�
    Select Case intSelNum
        Case 1
            lblTotalMent.Caption = "�̻�Ļ��ൿ�� ������ �����ϴ�."
        Case 2
            lblTotalMent.Caption = "�̻�Ļ��ൿ�� ������ ������ �Ļ���� ���ɼ��� �����ϴ�. ��� ȯ���� ��ȭ�� ���� ������ ���ʽÿ�."
        Case 3
            lblTotalMent.Caption = "�̻�Ļ��ൿ�� ������ �ߵ�̹Ƿ� �Ļ���� ���ɼ��� �����Ƿ� ��������� �����ô� ���� �ٶ����մϴ�."
        Case 4
            lblTotalMent.Caption = "�̻�Ļ��ൿ�� ������ ���� �����̹Ƿ� �Ļ���� ���ɼ��� �����Ƿ� ��������� �����ô� ���� �ٶ����մϴ�."
    End Select
    '�������� �ൿ ����Ʈ �����ֱ�
    j = 0
    For i = 1 To 26
        strAnswer = Mid(strAnswers, i, 1)
        sprDisorder.Col = 1
        If strAnswer = "1" Or strAnswer = "2" Or strAnswer = "3" Then
            sprDisorder.Row = i
            grdBadList.RowS = grdBadList.RowS + 1
            grdBadList.TextMatrix(j, 0) = j + 1
            grdBadList.TextMatrix(j, 1) = sprDisorder.Text
            j = j + 1
        End If
    Next i
    grdBadList.RowHeight(-1) = 300
    grdBadList.BackColorBkg = FRM_GRAY

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        strLastAs = rValue(0, 0)
    Else
        strLastAs = ""
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub DrawPoint(intScore As Integer)
    Dim i As Integer
    
    For i = 0 To 24
        shpPoint(i).Visible = False
    Next i
    
    If WhatSex = "M" Then
        Select Case intScore
            Case Is < 3
                shpPoint(0).Visible = True
            Case Is < 6
                shpPoint(1).Visible = True
            Case Is < 8
                shpPoint(2).Visible = True
            Case Is < 10
                shpPoint(3).Visible = True
            Case Is < 13
                shpPoint(4).Visible = True
            Case Is < 15
                shpPoint(5).Visible = True
            Case Is = 15
                shpPoint(6).Visible = True
            Case Is = 16
                shpPoint(7).Visible = True
            Case Is = 17
                shpPoint(8).Visible = True
            Case Is = 18
                shpPoint(10).Visible = True
            Case Is = 19
                shpPoint(11).Visible = True
            Case Is = 20
                shpPoint(12).Visible = True
            Case Is = 21
                shpPoint(14).Visible = True
            Case Is = 22
                shpPoint(15).Visible = True
            Case Is = 23
                shpPoint(17).Visible = True
            Case Is < 25
                shpPoint(18).Visible = True
            Case Is < 28
                shpPoint(19).Visible = True
            Case Is < 30
                shpPoint(20).Visible = True
            Case Is < 33
                shpPoint(21).Visible = True
            Case Is < 35
                shpPoint(22).Visible = True
            Case Is < 38
                shpPoint(23).Visible = True
            Case Else
                shpPoint(24).Visible = True
        End Select
    Else
        Select Case intScore
            Case Is < 3
                shpPoint(0).Visible = True
            Case Is < 6
                shpPoint(1).Visible = True
            Case Is < 8
                shpPoint(2).Visible = True
            Case Is < 10
                shpPoint(3).Visible = True
            Case Is < 13
                shpPoint(4).Visible = True
            Case Is < 16
                shpPoint(5).Visible = True
            Case Is < 19
                shpPoint(6).Visible = True
            Case Is = 19
                shpPoint(7).Visible = True
            Case Is = 20
                shpPoint(8).Visible = True
            Case Is = 21
                shpPoint(10).Visible = True
            Case Is = 22
                shpPoint(11).Visible = True
            Case Is = 23
                shpPoint(12).Visible = True
            Case Is = 24
                shpPoint(13).Visible = True
            Case Is = 25
                shpPoint(14).Visible = True
            Case Is = 26
                shpPoint(15).Visible = True
            Case Is = 27
                shpPoint(17).Visible = True
            Case Is < 29
                shpPoint(18).Visible = True
            Case Is < 32
                shpPoint(19).Visible = True
            Case Is < 35
                shpPoint(20).Visible = True
            Case Is < 37
                shpPoint(21).Visible = True
            Case Is < 40
                shpPoint(22).Visible = True
            Case Is < 43
                shpPoint(23).Visible = True
            Case Else
                shpPoint(24).Visible = True
        End Select
    End If
End Sub

Private Function MarkUp() As Integer
'�� �׸��� �亯�� �����ű��. ������ �����Ѵ�.
    Dim i As Integer
    Dim intTotal As Integer
    intTotal = 0
    With sprDisorder
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            Select Case Trim(.Text)
                Case "1": intTotal = intTotal + 3
                Case "2": intTotal = intTotal + 2
                Case "3": intTotal = intTotal + 1
                Case "4": intTotal = intTotal + 0
                Case "5": intTotal = intTotal + 0
                Case "6": intTotal = intTotal + 0
            End Select
        Next i
    End With
    
    MarkUp = intTotal
End Function

Private Sub InitialGrid()
    With grdBadList
        .Clear
        .BorderStyle = flexBorderNone
        .Appearance = flexFlat
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .ColS = 2
        .RowS = 0
        .FixedCols = 1
        
        .BackColorBkg = vbWhite
        .BackColor = FRM_GRAY
        .BackColorFixed = FRM_SKYBLUE
        .ForeColorFixed = &HE7BE7B
        .GridColor = vbWhite
        .GridLineWidth = 2
        .GridColorFixed = FRM_SKYBLUE

        .Font.Size = 8
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(0) = 400
        .ColWidth(1) = 4500
    End With
End Sub

Private Sub InitialQuestion()
'�Ļ���ְ˻� �׸��� �����Ѵ�
    Dim i As Integer, j As Integer

    With sprDisorder
        .EditEnterAction = EditEnterActionDown
        .ScrollBars = ScrollBarsNone
        .MaxCols = 2
        .MaxRows = 26
        .RowHeight(-1) = 12.5

        .ColWidth(0) = 3
        .ColWidth(1) = 45
        .ColWidth(2) = 4
        .Row = 0
        .Col = 1: .Text = "�� �� �� ��"
        .Col = 2: .Text = "�Է�"
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            .CellType = CellTypeComboBox
            .TypeComboBoxString = "6"
            .TypeComboBoxString = "5"
            .TypeComboBoxString = "4"
            .TypeComboBoxString = "3"
            .TypeComboBoxString = "2"
            .TypeComboBoxString = "1"
        Next i

        .Col = 1
        .Row = 1: .Text = "����� ���� �ηƴ�."
        .Row = 2: .Text = "�谡 ���ĵ� �Ļ縦 ���� �ʴ´�."
        .Row = 3: .Text = "���� ���Ŀ� �����ϰ� �ִ�."
        .Row = 4: .Text = "������ �� ���� ������ �� ���� �ִ�."
        .Row = 5: .Text = "������ ���� �������� ������ �Դ´�."
        .Row = 6: .Text = "�ڽ��� �԰� �ִ� ������ ����а� ������ �˰� �Դ´�."
        .Row = 7: .Text = "���̳� ���� ���� ź��ȭ���� ���� ������ Ư�� ���Ѵ�."
        .Row = 8: .Text = "���� ���� ������ ������ �ٸ� ������� �����ϴ� �� ����."
        .Row = 9: .Text = "�԰� �� ���� ���Ѵ�."
        .Row = 10: .Text = "�԰� �� ���� ���� ��å���� ������."
        .Row = 11: .Text = "�ڽ��� �� �� ���������߰ڴٴ� ������ ���� ���� �� ����."
        .Row = 12: .Text = "��� �� �� ����� ���� ������ ������ ���� ����ϰų� �����Ѵ�."
        .Row = 13: .Text = "������ ���� �ʹ� �����ٰ� �����Ѵ�."
        .Row = 14: .Text = "���� ���� �ȴٴ� ������ ���Ĺ��� ���� ����."
        .Row = 15: .Text = "�Ļ�ð��� �ٸ� ������� �� ���."
        .Row = 16: .Text = "������ �� ������ ���Ѵ�."
        .Row = 17: .Text = "ü�������� ���� ���̾�Ʈ�� ������ �Դ´�."
        .Row = 18: .Text = "������ ���� �λ��� �����Ѵٴ� ������ ���."
        .Row = 19: .Text = "���Ŀ� ���� �ڽ��� �����ɷ��� �����Ѵ�."
        .Row = 20: .Text = "�ٸ� ������� ������ ������ �Ե��� �����ϴ� �� ����."
        .Row = 21: .Text = "���Ŀ� ���� ���� �ð��� ������ �����Ѵ�."
        .Row = 22: .Text = "�� ������ �԰� ���� ������ ��ġ �ʴ�."
        .Row = 23: .Text = "ü���� ���̱� ���� ��̳� �ٸ� ���� �ϰ� �ִ�."
        .Row = 24: .Text = "���� ��� �ִ� ������ �ִ�."
        .Row = 25: .Text = "���ο� �⸧�� ���� �Դ� ���� ����."
        .Row = 26: .Text = "�Ļ� �� ���ϰ� ���� �浿�� ������."
        .Row = -1: .Lock = True
    End With
    lblTotal.Caption = ""
    lblTotalMent.Caption = ""

    Call InitialGrid
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

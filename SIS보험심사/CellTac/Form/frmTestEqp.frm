VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmTestEqp 
   Caption         =   " 장비 VS 검사코드 설정"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11970
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList imlList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":0000
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":059A
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTstListEqp 
      Height          =   5850
      Left            =   45
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   585
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   10319
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "검사 코드(마스터)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "검사 명 (마스터)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtTestNm 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2670
   End
   Begin VB.TextBox txtTestCD 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   5490
      MaxLength       =   10
      TabIndex        =   21
      Text            =   "1234567890"
      Top             =   1020
      Width           =   1020
   End
   Begin VB.TextBox txtVIndex 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   11010
      MaxLength       =   5
      TabIndex        =   12
      Top             =   660
      Width           =   675
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmTestEqp.frx":0B34
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "검사실 검사항목과 장비 검사항목을 연결 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   360
         Left            =   11520
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox txtTstcdEqp 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   675
      Width           =   1425
   End
   Begin VB.TextBox txtTstnmEqp 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1020
      Width           =   1650
   End
   Begin VB.TextBox lblTstcdEqp 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "1234567890"
      Top             =   660
      Width           =   1005
   End
   Begin VB.TextBox lblTstnmEqp 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6495
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   660
      Width           =   3075
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   5
      Top             =   6450
      Width           =   11940
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   1515
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   2895
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   4260
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
   End
   Begin MSComctlLib.ListView lvwTestListLab 
      Height          =   5085
      Left            =   4080
      TabIndex        =   15
      Top             =   1350
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   8969
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "검사 코드(마스터)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "검사 명 (마스터)"
         Object.Width           =   2540
      EndProperty
   End
   Begin HSCotrol.CButton cmdEqpItm_Add 
      Height          =   300
      Left            =   3015
      TabIndex        =   3
      Top             =   675
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":1DB6
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdEqpItm_Del 
      Height          =   300
      Left            =   3015
      TabIndex        =   4
      Top             =   1005
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":1F10
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.UserPanel pnlTestitem 
      Height          =   5385
      Left            =   7545
      TabIndex        =   10
      Top             =   1425
      Visible         =   0   'False
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   9499
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwTestitem 
         Height          =   4995
         Left            =   105
         TabIndex        =   11
         Top             =   270
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8811
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin HSCotrol.CButton cmdAdd 
      Height          =   300
      Left            =   9975
      TabIndex        =   24
      Top             =   1005
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":24AA
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdDel 
      Height          =   300
      Left            =   10935
      TabIndex        =   25
      Top             =   1005
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":2604
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdSerch 
      Height          =   300
      Left            =   6525
      TabIndex        =   26
      Top             =   1005
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   529
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":2B9E
      MaskColor       =   0
      PicCapAlign     =   1
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin VB.Frame Frame5 
      Height          =   6015
      Left            =   4020
      TabIndex        =   23
      Top             =   450
      Width           =   30
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "검사코드(LIS) :"
      Height          =   180
      Left            =   4170
      TabIndex        =   27
      Top             =   1065
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "View Index :"
      Height          =   180
      Left            =   9915
      TabIndex        =   19
      Top             =   705
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "장비 검사명 :"
      Height          =   180
      Left            =   60
      TabIndex        =   18
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "장비 검사 코드 :"
      Height          =   180
      Left            =   60
      TabIndex        =   17
      Top             =   735
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "검사코드(장비):"
      Height          =   180
      Left            =   4170
      TabIndex        =   16
      Top             =   705
      Width           =   1290
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
Private Const AUTO_VEFY     As String = "YES"
Private Const AUTO_VEFN     As String = "NO"

Private Const TLB_TEMP      As String = "TEMPTEABLE"
Private Const TLB_RESULT    As String = "INTERFACE003"

Private mAdoRs              As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdPrint_Click
        Case 1
            Call cmdSave_Click
        Case 2
            Call cmdClear_Click
        Case 3
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub cmdPrint_Click()
    Call PrintFrom(lvwTestListLab.ListItems)
End Sub

Private Sub cmdAdd_Click()
    Dim itemX As ListItem
    Dim itemS As ListItem
    Dim itemZ As ListSubItem
    If Trim(lblTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    If Trim(lblTstnmEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    If Trim(txtTestCD) = "" Then
        Call ShowMessage("장비검사코드와 연결할 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    Set itemS = lvwTestListLab.FindItem(Trim(lblTstcdEqp), lvwText, , lvwWhole)
    
    If Not itemS Is Nothing Then
        If vbYes = MsgBox(Trim(lblTstcdEqp) & " 장비검사 코드는 이미 있습니다. 바꾸시겠습니까?", vbExclamation + vbYesNo) Then
            Call lvwTestListLab.ListItems.Remove(itemS.Index)
            Set itemX = lvwTestListLab.ListItems.Add(, , Trim(lblTstcdEqp), , "TST_M")
            With itemX
                .SubItems(1) = Trim(lblTstnmEqp)
                .SubItems(2) = Trim(txtTestCD)
                .SubItems(3) = Trim(txtTestNm)
                .SubItems(4) = Val(txtVIndex)
                .SubItems(5) = "N"
            End With
        End If
    Else
        Set itemX = lvwTestListLab.ListItems.Add(, , Trim(lblTstcdEqp), , "TST_M")
        With itemX
            .SubItems(1) = Trim(lblTstnmEqp)
            .SubItems(2) = Trim(txtTestCD)
            .SubItems(3) = Trim(txtTestNm)
            .SubItems(4) = Val(txtVIndex)
            .SubItems(5) = "N"
        End With
    End If
    lblTstcdEqp = ""
    lblTstnmEqp = ""
    txtTestCD = ""
    txtTestNm = ""
    txtVIndex = ""
    
    Set itemX = Nothing
    Set itemS = Nothing
    
    lvwTstListEqp.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call Form_Clear
End Sub

Private Sub cmdDel_Click()
    Dim itemX   As ListItem
    Dim itemXs  As ListItems
    Dim i       As Long
    
    Set itemX = lvwTestListLab.SelectedItem
    
    If itemX Is Nothing Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        Exit Sub
    Else
        Set itemXs = lvwTestListLab.ListItems
        For i = itemXs.Count To 1 Step -1
           If itemXs(i).Selected = True Then
              lvwTestListLab.ListItems.Remove i
           End If
        Next
    End If
    Set itemX = Nothing
    
   lvwTstListEqp.SetFocus
End Sub

Private Sub cmdEqpItm_Add_Click()
    Dim objEqpItem  As clsCommon
    Dim strTemp     As String
        
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 입력 하시오.   ")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("장비 검사명이 없습니다. 검사명을 입력 하시오.   ")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    Set objEqpItem = New clsCommon
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        If .Let_EqpTestItem(INS_CODE, Trim(txtTstcdEqp), Trim(txtTstnmEqp)) Then
            txtTstcdEqp = ""
            txtTstnmEqp = ""
            txtTstcdEqp.SetFocus
        Else
            Call ShowMessage("오류가있어 저장 하지 못했습니다.")
        End If
    End With
    
    Set objEqpItem = Nothing
    Call GetEqp_Item(INS_CODE)
    txtTstcdEqp.SetFocus
End Sub

Private Sub cmdEqpItm_Del_Click()
    Dim itemX As ListItem
    Dim objEqpItem As clsCommon
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If itemX Is Nothing Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        Exit Sub
    Else
        Set objEqpItem = New clsCommon
        With objEqpItem
            .SetAdoCn AdoCn_Jet
            If Not .Del_EqpTestItem(INS_CODE, Trim(itemX.Text)) Then
                Call ShowMessage("오류가있어 삭제 하지 못했습니다.")
            End If
        End With
    End If
    Set itemX = Nothing
    Set objEqpItem = Nothing
    Call GetEqp_Item(INS_CODE)
End Sub

Private Sub cmdSave_Click()
    Dim objEqpMatch     As clsCommon
    Dim SaveList        As Scripting.Dictionary
    Dim itemX           As ListItem
    
    Set objEqpMatch = New clsCommon
        objEqpMatch.SetAdoCn AdoCn_Jet
    
    If lvwTestListLab.ListItems.Count < 1 Then
        If Not objEqpMatch.Set_EqpMatchItem(INS_CODE) Then
            Call ShowMessage(" 오류가 있어 저장하지 못했습니다.")
        End If
    Else
        If Not objEqpMatch.Set_EqpMatchItem(INS_CODE) Then
            Call ShowMessage(" 오류가 있어 저장하지 못했습니다.")
        End If
        For Each itemX In lvwTestListLab.ListItems
            Set SaveList = New Scripting.Dictionary
            With SaveList
                .Add "EQP_CD", INS_CODE
                .Add "TESTCD_EQP", Trim(itemX.Text)
                .Add "TESTNM_EQP", Trim(itemX.SubItems(1))
                .Add "TESTCD", Trim(itemX.SubItems(2))
                .Add "TESTNM", Trim(itemX.SubItems(3))
                .Add "OUT_SEQ", Trim(itemX.SubItems(4))
                .Add "AUTOVERIFY", IIf(Trim(itemX.SubItems(5)) = AUTO_VEFY, 1, 0)
            End With
            If Not objEqpMatch.Let_EqpMatchItem(SaveList) Then Exit For
            Set SaveList = Nothing
        Next itemX
    End If
    
    Set SaveList = Nothing
    Set objEqpMatch = Nothing
End Sub

Private Sub cmdSerch_Click()
    Dim objTestItem As clsCommon
    
    Set objTestItem = New clsCommon
    
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_TestItem(INS_CODE)
    End With
    
    Set objTestItem = Nothing
    
    Call PopUp_List.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        If Not mAdoRs.EOF Then
            Call DataLoadLvw(PopUp_List, vbCr, vbTab, mAdoRs.GetString)
            Call PopUp_List.ListItems.Remove(PopUp_List.ListItems.Count)
            
            With pnlTestitem
                .Visible = True
                .ZOrder
            End With
            PopUp_List.SetFocus
        End If
    Else
        Call ShowMessage("등록된 검사항목이 없습니다.")
    End If
    
    Set mAdoRs = Nothing
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear_Click
    Call Set_ListView
    Call GetEqp_Item(INS_CODE)
    
    With pnlTestitem
        .Moveble = True
    End With
    
    Set PopUp_List = lvwTestitem
    
    With PopUp_List
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Add 1, , "검사코드", (PopUp_List.Width - 310) * 0.4
            .Add 2, , "검사항목", (PopUp_List.Width - 310) * 0.6
            .Add 3, , "타입", 0
            .Add 4, , "Unit", 0
            .Add 5, , "PrtSeq", 0
        End With
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub Image1_DblClick()
    If lvwTstListEqp.Top > txtTstnmEqp.Top Then
        Call lvwTstListEqp.Move(Label1.left, CaptionBar1.Height, lvwTstListEqp.Width, ScaleHeight - (CaptionBar1.Height + fraCmdBar.Height + 30))
        txtTstcdEqp.Enabled = False
        txtTstnmEqp.Enabled = False
        cmdEqpItm_Add.Enabled = False
        cmdEqpItm_Del.Enabled = False
        Call lvwTstListEqp.ZOrder
    Else
        Call lvwTstListEqp.Move(Label1.left, lvwTestListLab.Top, lvwTstListEqp.Width, lvwTestListLab.Height)
        txtTstcdEqp.Enabled = True
        txtTstnmEqp.Enabled = True
        cmdEqpItm_Add.Enabled = True
        cmdEqpItm_Del.Enabled = True
    End If
End Sub

Private Sub lvwTestListLab_Click()
    Dim itemX As ListItem

    Set itemX = lvwTestListLab.SelectedItem
    If Not itemX Is Nothing Then
        With itemX
            lblTstcdEqp = .Text             '장비검사 코드
            lblTstnmEqp = .SubItems(1)      '장비검사 이름
            txtTestCD = .SubItems(2)        '임상검사 코드
            txtTestNm = .SubItems(3)        '임상검사 이름
            txtVIndex = .SubItems(4)        '출력 순서
        End With
    End If
End Sub

Private Sub lvwTestListLab_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetListView_Sort(lvwTestListLab, ColumnHeader)
End Sub

Private Sub lvwTstListEqp_Click()
    Dim itemX As ListItem
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If itemX Is Nothing Then
        Exit Sub
    Else
        txtTstcdEqp = Trim(itemX)
        txtTstnmEqp = Trim(itemX.SubItems(1))
    End If
    Set itemX = Nothing
End Sub

Private Sub lvwTstListEqp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetListView_Sort(lvwTstListEqp, ColumnHeader)
End Sub

Private Sub lvwTstListEqp_DblClick()
    Dim itemX As ListItem
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If itemX Is Nothing Then
        Exit Sub
    Else
        lblTstcdEqp = Trim(itemX.Text)
        lblTstnmEqp = Trim(itemX.SubItems(1))
        txtTestCD.SetFocus
    End If
    
    Set itemX = Nothing
End Sub

Private Sub lvwTstListEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lvwTstListEqp_DblClick
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub pnlTestitem_CloseMe()
    pnlTestitem.Visible = False
End Sub

Private Sub GetEqp_Item(ByVal strEqp_Cd As String)
    Dim objEqpItem  As clsCommon
    Dim strTemp     As String
    Dim itemX       As ListItem
    
    Set objEqpItem = New clsCommon
    
    lvwTstListEqp.ListItems.Clear
    lvwTestListLab.ListItems.Clear
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpTestItem(Trim(strEqp_Cd))
        If Not mAdoRs Is Nothing Then
            Do Until mAdoRs.EOF
                Set itemX = lvwTstListEqp.ListItems.Add(, , Trim(mAdoRs("TESTCD_EQP") & ""), , "TST_E")
                    itemX.SubItems(1) = Trim(mAdoRs("TESTNM_EQP") & "")
                    'ItemX.EnsureVisible
                mAdoRs.MoveNext
            Loop
        End If
        Set mAdoRs = Nothing
        Set mAdoRs = .Get_TestItemList(Trim(strEqp_Cd))
        If Not mAdoRs Is Nothing Then
            Do Until mAdoRs.EOF
                Set itemX = lvwTestListLab.ListItems.Add(, , Trim(mAdoRs("TESTCD_EQP") & ""), , "TST_M")
                With itemX
                    .SubItems(1) = Trim(mAdoRs("TESTNM_EQP") & "")
                    .SubItems(2) = Trim(mAdoRs("TESTCD") & "")
                    .SubItems(3) = Trim(mAdoRs("TESTNM") & "")
                    .SubItems(4) = Trim(mAdoRs("OUT_SEQ") & "")
                    .SubItems(5) = Trim(mAdoRs("AUTOVERIFY") & "")
                End With
                mAdoRs.MoveNext
            Loop
        End If
    End With
    
    Set mAdoRs = Nothing
    Set objEqpItem = Nothing
End Sub

Private Sub Set_ListView()
    Dim lvwWidth    As Long
    
    With lvwTstListEqp
        .View = lvwReport
        Set .SmallIcons = imlList
        Set .ColumnHeaderIcons = imlList
        
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        lvwWidth = .Width - 310
        With .ColumnHeaders
            .Clear
            Call .Add(1, "Code", "검사코드(순서)", lvwWidth * 0.4)
            Call .Add(2, "Name", "검사명", lvwWidth * 0.6)
        End With
    End With
    
    With lvwTestListLab
        .View = lvwReport
        Set .SmallIcons = imlList
        Set .ColumnHeaderIcons = imlList
        
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = True
        lvwWidth = .Width - 310
        With .ColumnHeaders
            .Clear
            Call .Add(1, "CodeEqp", "장비코드", lvwWidth * 0.15)
            Call .Add(2, "NameEqp", "장비검사명", lvwWidth * 0.2)
            Call .Add(3, "Code", "검사코드", lvwWidth * 0.17)
            Call .Add(4, "Name", "검사명", lvwWidth * 0.35)
            Call .Add(5, "Prtseq", "출력순서", lvwWidth * 0.125, lvwColumnRight)
            Call .Add(6, "Vetify", "자동확인", lvwWidth * 0, lvwColumnCenter)
        End With
    End With
    
End Sub

Private Sub Form_Clear()
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    lblTstcdEqp = ""
    lblTstnmEqp = ""
    txtTestCD = ""
    txtTestNm = ""
End Sub

Private Sub PopUp_List_DblClick()
    Dim itemX As ListItem
    Set itemX = PopUp_List.SelectedItem
    
    If Not itemX Is Nothing Then
        txtTestCD = itemX.Text
        txtTestNm = itemX.SubItems(1)
        Call pnlTestitem_CloseMe
        txtVIndex.SetFocus
    End If
End Sub

Private Sub PopUp_List_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PopUp_List_DblClick
        KeyAscii = 0
    End If
End Sub

Private Sub txtTestCD_Change()
    txtTestNm = ""
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    txtTestCD.Locked = False
    If KeyAscii = vbKeyReturn Then
        Call cmdSerch_Click
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtTstnmEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdEqpItm_Add_Click
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtVIndex_GotFocus()
    Call TextBoxs_GotFocus(txtVIndex)
End Sub

Private Sub txtVIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtVIndex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtVIndex.IMEMode = 8
End Sub

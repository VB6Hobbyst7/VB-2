VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmBBS800 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "공통코드 인덱스"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS800.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Save"
      Height          =   435
      Left            =   6135
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   7785
      Width           =   1530
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   435
      Left            =   7755
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7785
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F4F0F2&
      Caption         =   "E&xit"
      Height          =   435
      Left            =   9225
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   7785
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "각 필드에 대한 참조사항"
      Height          =   6990
      Left            =   5910
      TabIndex        =   1
      Top             =   300
      Width           =   4890
      Begin VB.TextBox txtKeyName 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1170
         TabIndex        =   10
         Top             =   825
         Width           =   3510
      End
      Begin VB.TextBox txtKeyCode 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1170
         TabIndex        =   8
         Top             =   405
         Width           =   2070
      End
      Begin FPSpread.vaSpread tblFields 
         Height          =   2265
         Left            =   135
         TabIndex        =   5
         Top             =   1395
         Width           =   4590
         _Version        =   196608
         _ExtentX        =   8096
         _ExtentY        =   3995
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   8
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmBBS800.frx":076A
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00F1F5F4&
         Height          =   2655
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmBBS800.frx":0C3D
         Top             =   4110
         Width           =   4560
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Key Name"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Key Code"
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   3885
         Width           =   1305
      End
   End
   Begin FPSpread.vaSpread ssTableKey 
      Height          =   7785
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   5715
      _Version        =   196608
      _ExtentX        =   10081
      _ExtentY        =   13732
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   5
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmBBS800.frx":0C43
   End
End
Attribute VB_Name = "frmBBS800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fRKey As String             ' Root Key
Private fTName As String            ' Access Table
Private fSKey As String             ' Current Row Key
Private fRowPos As Integer          ' Current Row

Private NewFg As Boolean
Private Sub cmdClear_Click()
    
    txtKeyCode.Text = ""
    Call ClearRtn
    Call LoadIndexList
    
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim objSSql As clsBBSMSTStatement
    Dim sSql As String
    Dim i As Integer
    Dim strFields As String
    
    strFields = ""
    With tblFields
        .Col = 2
        For i = 1 To .DataRowCnt
            .Row = i
            strFields = strFields & .value & ";"
        Next
    End With
    If txtKeyCode = "" Then
        MsgBox "KeyCode를 입력하십시오.", vbInformation + vbOKOnly, "입력오류"
        txtKeyCode.SetFocus
        Exit Sub
    End If
    
    Set objSSql = New clsBBSMSTStatement
    objSSql.setDbConn DbConn
    If objSSql.Save_frmCTControl(NewFg, txtKeyCode.Text, txtKeyName.Text, Replace(strFields, "'", "''"), txtRemark.Text) = True Then
        txtKeyCode.Text = ""
        Call ClearRtn
        Call LoadIndexList
    Else
        MsgBox "저장오류입니다.확인후 저장하세요."
    End If
    
    Set objSSql = Nothing
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    fRowPos = -1: fSKey = ""
    LoadIndexList
    LoadTableFields

End Sub

Private Sub LoadIndexList()
    Dim objSSql As clsBBSMSTStatement
    Dim Rs As New DrRecordSet
    Dim sSql As String
        
    
    Set objSSql = New clsBBSMSTStatement
    objSSql.setDbConn DbConn
    
    Set Rs = objSSql.Get_frmCtControlLoad
    
    With ssTableKey
        .MaxRows = 0
        .MaxRows = Rs.RecordCount
        .Row = 0
        While (Not Rs.EOF)
            .Row = .Row + 1
            .Col = 1: .value = Rs.Fields("CdVal1").value
            .Col = 2: .value = Rs.Fields("Field1").value
            Rs.MoveNext
        Wend
    End With
    
    Rs.RsClose
    Set Rs = Nothing
    
    Set objSSql = Nothing
End Sub


Private Sub LoadTableFields()
    Dim objSSql As clsBBSMSTStatement
    Dim i As Integer, sSql As String
    
    Dim dsTableKey As New DrRecordSet
    Dim strFields As String
    
    Set objSSql = New clsBBSMSTStatement
    objSSql.setDbConn DbConn
    
    Set dsTableKey = objSSql.Get_frmCtControlTable
    
    tblFields.MaxRows = 0
        
    If dsTableKey.EOF Then dsTableKey.RsClose: Exit Sub
    strFields = IIf(IsNull(dsTableKey.Fields("Text1").value) = True, "", dsTableKey.Fields("Text1").value)
    
    While (Trim(strFields) <> "")
        tblFields.MaxRows = tblFields.MaxRows + 1
        tblFields.Row = tblFields.MaxRows
        tblFields.Col = 1
        tblFields.Text = medShift(strFields, ";")
    Wend
    
    dsTableKey.RsClose
    Set dsTableKey = Nothing
    Set objSSql = Nothing
    medMain.stsBar.Panels(2).Text = ""
End Sub

Private Sub ClearRtn()

    txtKeyName.Text = ""
    With tblFields
        .Row = 1: .Row2 = .MaxRows
        .Col = 2: .Col2 = 2
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    txtRemark.Text = ""
    
End Sub


Private Sub lblTableName_Change()
    LoadIndexList
    LoadTableFields
End Sub

Private Sub ssTableKey_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objSSql As clsBBSMSTStatement
    Dim i As Integer, sSql As String
    Dim dsTableKey As DrRecordSet
    Dim strFields As String
    Dim aryFields() As String
    
    If Row < 1 Then Exit Sub
    Set objSSql = New clsBBSMSTStatement
    objSSql.setDbConn DbConn
    
    ssTableKey.Row = Row: ssTableKey.Col = 1
    
    Set dsTableKey = objSSql.Get_frmCtControlTableKey(ssTableKey.value)

    If dsTableKey.EOF Then dsTableKey.RsClose: Exit Sub
    
    txtKeyCode.Text = IIf(IsNull(dsTableKey.Fields("CdVal1").value) = True, "", dsTableKey.Fields("CdVal1").value)
    txtKeyName.Text = IIf(IsNull(dsTableKey.Fields("Field1").value) = True, "", dsTableKey.Fields("Field1").value)

    aryFields() = Split(dsTableKey.Fields("text1").value, ";")
    
    With tblFields
        For i = 1 To .MaxRows
            If i > (UBound(aryFields) + 1) Then Exit For
            
            .Row = i
            .Col = 2
            .value = aryFields(i - 1)
        Next
    End With
    
    txtRemark.Text = IIf(IsNull(dsTableKey.Fields("Text2").value) = True, "", dsTableKey.Fields("Text2").value)
    dsTableKey.RsClose
    Set dsTableKey = Nothing
    Set objSSql = Nothing
End Sub

Private Sub tblFields_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then tblFields.MaxRows = tblFields.MaxRows + 1
End Sub

Private Sub txtKeyCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKeyCode_LostFocus()
    Dim objSSql As clsBBSMSTStatement
    Dim i As Integer, sSql As String
    Dim dsTableKey As Object
    Dim strFields As String

    If Screen.ActiveForm.name <> Me.name Then Exit Sub
    If ActiveControl.name = cmdClear.name Then Exit Sub
    If ActiveControl.name = cmdOK.name Then Exit Sub
    If Len(txtKeyCode) > 5 Then
        MsgBox "KeyCode의 최대허용길이는 다섯자리입니다", vbInformation + vbOKOnly, "입력오류"
        txtKeyCode = "": txtKeyCode.SetFocus
    End If
    Set objSSql = New clsBBSMSTStatement
    objSSql.setDbConn DbConn
    
    Set dsTableKey = objSSql.Get_frmCtControlTableKeyCode(txtKeyCode.Text)
    
    If dsTableKey.EOF Then
        NewFg = True
        dsTableKey.RsClose
        If tblFields.MaxRows = 0 Then tblFields.MaxRows = 1
        Call ClearRtn
        Exit Sub
    End If
    
    NewFg = False
    
    txtKeyCode.Text = IIf(IsNull(dsTableKey.Fields("CdVal1").value) = True, "", dsTableKey.Fields("CdVal1").value)
    txtKeyName.Text = IIf(IsNull(dsTableKey.Fields("Field1").value) = True, "", dsTableKey.Fields("Field1").value)
        
    With tblFields
    
        .Row = 0
        strFields = IIf(IsNull(dsTableKey.Fields("Text1").value) = True, "", (dsTableKey.Fields("Text1").value))
        While (Trim(strFields) <> "")
            .Row = .Row + 1
            .Col = 2
            .value = medShift(strFields, ";")
        Wend
        
        txtRemark.Text = IIf(IsNull(dsTableKey.Fields("Text2").value) = True, "", dsTableKey.Fields("Text2").value)
        
        dsTableKey.RsClose
        Set dsTableKey = Nothing
        Set objSSql = Nothing
    End With
    
End Sub







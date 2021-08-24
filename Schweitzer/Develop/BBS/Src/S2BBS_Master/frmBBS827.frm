VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS827 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "1회 발생코드(ABO & Rh)"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
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
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   4560
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   5880
      Width           =   1260
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   420
      Left            =   3240
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   5880
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   5880
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   5880
      Width           =   1260
   End
   Begin FPSpread.vaSpread tblABORH 
      Height          =   4110
      Left            =   2340
      TabIndex        =   0
      Tag             =   "10114"
      Top             =   1440
      Width           =   5620
      _Version        =   196608
      _ExtentX        =   9913
      _ExtentY        =   7250
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   15
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS827.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   8
   End
End
Attribute VB_Name = "frmBBS827"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call Query
    tblABORH.SetFocus
End Sub

Private Sub cmdSave_Click()
    If Save = True Then Call Query
End Sub

Private Sub Form_Load()
    Call Query
End Sub


Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblABORH
                .Col = .ActiveCol
                .Row = .ActiveRow
                .Action = ActionDeleteRow
            End With
    End Select
End Sub

'Private Sub mnuDelete_Click()
'    With tblABORH
'        .Col = .ActiveCol
'        .Row = .ActiveRow
'        .Action = ActionDeleteRow
'    End With
'End Sub

Private Sub tblABORH_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    With tblABORH
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
        
'        .Col = 1: .Col2 = .MaxCols
'        .Row = Row: .Row2 = Row
'        .BlockMode = True
'        .BackColor = &H8000000F
'        .BlockMode = False
    End With
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hWnd
    End With
    Set objPop = Nothing

'    Set mnuPopup = frmControl.mnuPopup
'    Set mnuDelete = frmControl.mnuSub
'    mnuDelete.Caption = "Delete"
'
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
End Sub










Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    Dim code As String
    Dim name As String
    
    
    medClearTable tblABORH
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_ABO_RH)
    Set objcom003 = Nothing
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            code = .Fields("cdval1").Value & ""
            name = .Fields("field1").Value & ""
            
            With tblABORH
                .Row = i
                If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
                .Col = 1: .Value = code
                .Col = 2: .Value = name
            End With
            
            .MoveNext
        Next i
    End With
    Set DrRS = Nothing
    
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim cdval1() As String
    Dim idx As Long
    Dim i As Long
    Dim code As String
    Dim name As String

On Error GoTo Save_error

    DBConn.BeginTrans


    Set objcom003 = New clsCom003
    
    '먼저 데이터베이스에서 자룔을 읽는다.
    Set DrRS = objcom003.OpenRecordSet(BC2_ABO_RH)
    If DrRS Is Nothing Then GoTo Save_error

    With DrRS
        idx = .RecordCount
        For i = 1 To idx
            ReDim Preserve cdval1(i - 1)
            cdval1(i - 1) = .Fields("cdval1").Value & ""
            
            .MoveNext
        Next i
    End With

    '읽은 자료를 모두 지운다.
    If idx > 0 Then
        For i = 0 To idx - 1
            objcom003.CDINDEX = BC2_ABO_RH
            objcom003.cdval1 = cdval1(i)
            If objcom003.Delete() = False Then GoTo Save_error
        Next i
    End If

    '입력된 자료를 저장한다.
    With tblABORH
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1: code = .Value
            .Col = 2: name = .Value
            
            If code = "" Then Exit For
            
            objcom003.CDINDEX = BC2_ABO_RH
            objcom003.cdval1 = code
            objcom003.field1 = name
            If objcom003.Save() = False Then GoTo Save_error
        Next i
    End With


    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:
    DBConn.RollbackTrans
    Save = False
    MsgBox Err.Description, vbExclamation
End Function

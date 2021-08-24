VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEQ공용_장비검사코드관리_조회 
   Caption         =   "장비검사코드관리"
   ClientHeight    =   8055
   ClientLeft      =   2445
   ClientTop       =   3675
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_장비검사코드관리_조회.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10875
   Begin VB.CommandButton CmdDelete 
      Caption         =   "삭제(&D)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8940
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "신규(&I)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7020
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "수정(&U)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7980
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9900
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "조회(&V)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6060
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   6
      Top             =   600
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread sprVIEW 
      Height          =   7275
      Left            =   60
      TabIndex        =   5
      Top             =   720
      Width           =   10755
      _Version        =   393216
      _ExtentX        =   18971
      _ExtentY        =   12832
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   19
      MaxRows         =   10
      SpreadDesigner  =   "frmEQ공용_장비검사코드관리_조회.frx":263A
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비검사코드조회"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   2880
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "frmEQ공용_장비검사코드관리_조회"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height의 초기값
Dim lngMeWidth      As Long '/Me.Width의 초기값

Private Type ConWhere   ' 사용자 정의 형식을 만듭니다.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Sub SUB_MM_CANCEL()
    barStatus.Max = 100
    barStatus.Value = 100
    
    Call SUB_MM_KEY_CLEAR
End Sub

Public Function FUNC_MM_DELETE() As Boolean
    FUNC_MM_DELETE = False
    
    Dim intActCol    As Integer
    Dim intActRow    As Integer
    
    '/1.삭제 조건 Check
    If sprVIEW.ActiveRow = 0 Then MsgBox "삭제할 내용을 선택하십시오", vbInformation, "확인": Exit Function
    
    '/2.삭제 질의
    If MsgBox("장비검사코드 : " & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & vbCrLf & _
              "장비검사명   : " & GET_CELL(sprVIEW, 2, sprVIEW.ActiveRow) & vbCrLf & vbCrLf & _
              "위 자료를 삭제하겠습니까?", vbQuestion + vbOKCancel, "삭제질의") = vbCancel Then Exit Function
    
    '/3.Process
    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    If sprVIEW.IsBlockSelected Then
        intActCol = sprVIEW.SelBlockCol
        intActRow = sprVIEW.SelBlockRow
    Else
        intActCol = sprVIEW.ActiveCol
        intActRow = sprVIEW.ActiveRow
    End If
    If sprVIEW.IsBlockSelected Then
        For intX = sprVIEW.SelBlockRow To sprVIEW.SelBlockRow2
            gstrQuy = "DELETE FROM EQ_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, intX) & "' "
            If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
        Next intX
    Else
        gstrQuy = "DELETE FROM EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & "' "
        If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    End If
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_DELETE = True
    
    MsgBox "삭제되었습니다!", vbInformation, "확인"
    
    '/4.화면처리
    Call FUNC_MM_VIEW
    sprVIEW.Col = intActCol
    sprVIEW.Row = intActRow
    sprVIEW.Action = ActionActiveCell
End Function

Private Sub SUB_MM_INITIAL()
    '/Form Resize를 위한 컨트롤 초기값 읽기
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Resize를 위한 초기값 설정
    lngMeHeight = 8535
    lngMeWidth = 10995
    
    '/화면 가운데 위치
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    '''Me.Show
    
    Call SUB_MM_CANCEL
End Sub

Public Sub SUB_MM_INPUT()
    gstrInputUpdate = "1" '/1.Input, 2.Update
    gstrInputUpdateYN = False

    frmEQ공용_장비검사코드관리_입력.Show vbModal

    If gstrInputUpdateYN = True Then
        Call FUNC_MM_VIEW
    End If
End Sub

Private Sub SUB_MM_KEY_CLEAR()
    If sprVIEW.MaxRows > 0 Then sprVIEW.MaxRows = 0
End Sub

Public Sub SUB_MM_UPDATE()
    Dim intActCol    As Integer
    Dim intActRow    As Integer
    
    If sprVIEW.ActiveRow = 0 Then MsgBox "수정할 대상을 선택하십시오!", vbInformation, "확인": Exit Sub
    
    gstrInputUpdate = "2" '/1.Input, 2.Update
    gstrInputUpdateYN = False
    gstrArgTemp1 = GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow)
    
    frmEQ공용_장비검사코드관리_입력.Show vbModal
    
    If gstrInputUpdateYN = True Then
        intActCol = sprVIEW.ActiveCol
        intActRow = sprVIEW.ActiveRow

        Call FUNC_MM_VIEW
    
        sprVIEW.Col = intActCol
        sprVIEW.Row = intActRow
        sprVIEW.Action = ActionActiveCell
    End If
End Sub

Public Function FUNC_MM_VIEW() As Boolean
    FUNC_MM_VIEW = False
    
On Error GoTo RTN_ERR

    Call SUB_MM_KEY_CLEAR
    
    If ConnDB_LOC = False Then End
    
    With sprVIEW
        gstrQuy = "SELECT A.*, B.ORDERCNT "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST A LEFT OUTER JOIN "
        gstrQuy = gstrQuy & vbCrLf & "       (SELECT EQCD, COUNT(EXCD) AS ORDERCNT "
        gstrQuy = gstrQuy & vbCrLf & "          FROM EX_MST "
        gstrQuy = gstrQuy & vbCrLf & "         WHERE (EQORDREADYN = 'Y' OR EQRESSENDYN = 'Y') "
        gstrQuy = gstrQuy & vbCrLf & "         GROUP BY EQCD) B ON A.EQCD = B.EQCD "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY A.EQSEQ "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            .MaxRows = ARC_LOC
            barStatus.Max = ARC_LOC
            intX = 0
            
            Do Until ADR_LOC.EOF
                intX = intX + 1: .Row = intX: barStatus.Value = intX
                
                .Col = 1:  .Text = UCase(Trim(ADR_LOC!EQCD & ""))   '/장비검사코드
                .Col = 2:  .Text = Trim(ADR_LOC!EQNM & "")          '/장비검사명
                .Col = 3:  .Text = Trim(ADR_LOC!EQORDYN & "")       '/장비검사코드 장비전송여부(Y.전송, N.미전송)
                .Col = 4                                            '/처방연동여부
                If Val(ADR_LOC!ORDERCNT & "") > 0 Then
                    .Text = "Y"
                Else
                    .Text = "N"
                End If
                .Col = 5:  .Text = Trim(ADR_LOC!EQUNIT & "")    '/검사결과단위
                
                .Col = 6:  .Text = Trim(ADR_LOC!EQRMLVAL & "")  '/정상참고치(남 Low)
                Select Case Trim(ADR_LOC!EQRMLREF & "")
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                    Case "3": .Text = .Text & " >=" '/3.이하
                    Case "4": .Text = .Text & " >"  '/4.미만
                End Select
                .Col = 7:  .Text = Trim(ADR_LOC!EQRMHVAL & "")  '/정상참고치(남 High)
                Select Case Trim(ADR_LOC!EQRMHREF & "")
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                    Case "3": .Text = .Text & " >=" '/3.이하
                    Case "4": .Text = .Text & " >"  '/4.미만
                End Select
                .Col = 8:  .Text = Trim(ADR_LOC!EQRFLVAL & "")  '/정상참고치(여 Low)
                Select Case Trim(ADR_LOC!EQRFLREF & "")
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                    Case "3": .Text = .Text & " >=" '/3.이하
                    Case "4": .Text = .Text & " >"  '/4.미만
                End Select
                .Col = 9: .Text = Trim(ADR_LOC!EQRFHVAL & "")  '/정상참고치(여 High)
                Select Case Trim(ADR_LOC!EQRFHREF & "")
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                    Case "3": .Text = .Text & " >=" '/3.이하
                    Case "4": .Text = .Text & " >"  '/4.미만
                End Select
                
                If IsNumeric(Trim(ADR_LOC!EQRSTRANGE & "")) Then    '/소수점표시제한
                    .Col = 10
                    Select Case Val(ADR_LOC!EQRSTRANGE & "")
                        Case 0: .Text = "전체표시"
                        Case Is > 0: .Text = Val(ADR_LOC!EQRSTRANGE & "") & " 자리"
                    End Select
                End If
                
                .Col = 11: .Text = Trim(ADR_LOC!EQLIMITVALUE1 & "") '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
                Select Case Trim(ADR_LOC!EQLIMITFLAG1 & "")
                    Case "0": .Text = "미적용"      '/0.미적용
                    Case "1": .Text = .Text & " >=" '/1.이하
                    Case "2": .Text = .Text & " >"  '/2.미만
                End Select
                
                .Col = 12: .Text = Trim(ADR_LOC!EQLIMITVALUE2 & "") '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
                Select Case Trim(ADR_LOC!EQLIMITFLAG2 & "")
                    Case "0": .Text = "미적용"      '/0.미적용
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                End Select
                
                .Col = 13   '/CUTOFF 구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
                Select Case Trim(ADR_LOC!EQCUTOFFGB & "")
                    Case "0": .Text = "적용안함"
                    Case "1": .Text = "상한 Positive"
                    Case "2": .Text = "하한 Positive"
                    Case "3": .Text = "장비결과값적용"
                End Select
                
                .Col = 14   '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
                Select Case Trim(ADR_LOC!EQCUTOFFNM & "")
                    Case "1": .Text = "Negative/Positive"
                    Case "2": .Text = "Neg/Pos"
                    Case "3": .Text = "Nonreactive/Reactive"
                    Case "4": .Text = "NEGATIVE/POSITIVE"
                    Case "5": .Text = "NEG/POS"
                End Select
                
                .Col = 15: .Text = Trim(ADR_LOC!EQCUTLVAL & "") '/CUTOFF 하한값
                Select Case Trim(ADR_LOC!EQCUTLREF & "")        '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
                    Case "1": .Text = .Text & " >=" '/1.이하
                    Case "2": .Text = .Text & " >"  '/2.미만
                End Select
                
                .Col = 16: .Text = Trim(ADR_LOC!EQCUTHVAL & "") '/CUTOFF 상한값
                Select Case Trim(ADR_LOC!EQCUTHREF & "")        '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
                    Case "1": .Text = .Text & " <=" '/1.이상
                    Case "2": .Text = .Text & " <"  '/2.초과
                End Select
                
                .Col = 17  '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
                Select Case Trim(ADR_LOC!EQCUTMNM & "")
                    Case "1": .Text = "Grayzone"
                    Case "2": .Text = "Weakly positive"
                    Case "3": .Text = "Low Titer"
                End Select
                
                .Col = 18   '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
                Select Case Trim(ADR_LOC!EQCUTRTYPE & "")
                    Case "1": .Text = "Negative/Positive"
                    Case "2": .Text = "Negative/Positive(수치)"
                    Case "3": .Text = "Negative/Grayzone(수치)/Positive(수치)"
                    Case "4": .Text = "Negative(수치)/Grayzone(수치)/Positive(수치)"
                End Select
                
                .Col = 19: .Text = Trim(ADR_LOC!EQSEQ & "") '/정렬순서
                
                If .MaxTextRowHeight(intX) > 13.3 Then .RowHeight(intX) = .MaxTextRowHeight(intX)
                
                ADR_LOC.MoveNext
            Loop
            ADR_LOC.Close: Set ADR_LOC = Nothing
        Else
            MsgBox "자료가 없습니다.", vbInformation, "확인"
        End If
    End With

    Call CloseDB_LOC

    FUNC_MM_VIEW = True
    
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    MsgBox Err.Description, vbCritical, "조회실패"
    Call CloseDB_LOC
End Function

Private Sub CmdDelete_Click()
    Call FUNC_MM_DELETE
End Sub

Private Sub CmdInput_Click()
    Call SUB_MM_INPUT
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub CmdUpdate_Click()
    Call SUB_MM_UPDATE
End Sub

Private Sub cmdView_Click()
    Call FUNC_MM_VIEW
    If sprVIEW.MaxRows > 0 Then sprVIEW.SetFocus
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    Call FUNC_MM_VIEW
End Sub

Private Sub Form_Resize()
On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : 높이가 늘어나는 개체 3개, 디자인상 해당 개체 위에 늘어난 개체가 2개
    For intX = 0 To UBound(CW)
        Select Case CW(intX).Nm
            Case cmdView.Name:      cmdView.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case CmdInput.Name:     CmdInput.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case CmdUpdate.Name:    CmdUpdate.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case CmdDelete.Name:    CmdDelete.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case cmdView.Name:      cmdView.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case cmdQuit.Name:      cmdQuit.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case barStatus.Name: barStatus.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case sprVIEW.Name:   sprVIEW.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height + (Me.Height - lngMeHeight)
        End Select
    Next intX
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Set frmEQ공용_장비검사코드관리_조회 = Nothing
End Sub

Private Sub sprVIEW_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim intActCol    As Integer
    Dim intActRow    As Integer
    
    If Row > 0 Then
        gstrInputUpdate = "2" '/1.Input, 2.Update
        gstrInputUpdateYN = False
        gstrArgTemp1 = GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow)
        
        frmEQ공용_장비검사코드관리_입력.Show vbModal
        
        If gstrInputUpdateYN = True Then
            intActCol = sprVIEW.ActiveCol
            intActRow = sprVIEW.ActiveRow
        
            Call FUNC_MM_VIEW
            
            sprVIEW.Col = intActCol
            sprVIEW.Row = intActRow
            sprVIEW.Action = ActionActiveCell
        End If
    End If
End Sub

Private Sub sprVIEW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call sprVIEW_DblClick(sprVIEW.ActiveCol, sprVIEW.ActiveRow)
End Sub

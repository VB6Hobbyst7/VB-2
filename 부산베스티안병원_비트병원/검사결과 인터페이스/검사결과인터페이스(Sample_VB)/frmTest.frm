VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmTest 
   Caption         =   "인터페이스 테스트"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "소모통계interface TEST"
      Height          =   2655
      Left            =   10080
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   525
      Left            =   9240
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton cmdTest2 
      Caption         =   "테스트(Update)"
      Height          =   555
      Left            =   4530
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "테스트(날짜 + 차트번호)"
      Height          =   525
      Left            =   2190
      TabIndex        =   2
      Top             =   0
      Width           =   2325
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   7425
      Left            =   30
      TabIndex        =   1
      Top             =   750
      Width           =   9795
      _Version        =   196608
      _ExtentX        =   17277
      _ExtentY        =   13097
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
      SpreadDesigner  =   "frmTest.frx":0000
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "테스트(날짜)"
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''데이터베이스 컬럼 명칭 리스트
''LBQACPNUM, LBQTRNDEP, LBQOCMNUM, LBQCHTNUM, LBQPATNAM,
''LBQPATTYP, LBQINSCOD, LBQDEPCOD, LBQDTRCOD, LBQWRDCOD,
''LBQROMCOD, LBQACPSTT, LBQEMGYON, LBQCASCOD, LBQTRTUID,
''LBQODRDTE, LBQACPDTM, LBQACPUID, OSPODRCOD, OSPODRDTM,
''OSPODRQTY, OSPODRDAY, OSPODRTMS, OSPSPLCMT, RSBSPMCOD,
''RSBITFYON, RSBACKSTT, RSBCONLVL, RSBPRNYON, RSBPRNDTM,
''RSBPRNUID, RSBTSTDTM, RSBTSTUID, RSBREDDTM, RSBREDUID,
''RSBRLTTYP, RSBPARYON, RSBTRYNUM, RSBRLTVAL, RESODRSEQ,
''RESSEQ, RESSUBSEQ, RESRSBACP, RESLABCOD, RESVOLFLG,
''RESRLTSEQ, RESRLTVAL, RESRLTCMT, RESREPTYP, RESUPDDTM,
''RESUPDUID, LABSHTNAM, LABRLTTYP, LABRLTOPT, LABDEFRLT,
''LABUNTCOD, LABMAXLIN, LABMAXLEN, LABJBSSEQ, LABRLTSEQ,
''LABVIWYON, LABPRTYON, LABDLTTYP, LABADPDTE, LABEXPDTE,
''LABGRPYON, LABSEEYON, LABSPMPOS, LRFPCSCOD, LRFVOLFLG,
''LRFHIGVAL, LRFLOWVAL, LRFSTDVAL, LRFBASVAL, LRFABLPRD,
''PbsResNum, OcmComStt, RsbBarCod


''검사결과 인터페이스 개체선언
Private LabResultObject As BITLabResultInterface.BITLabResultInterface

''날짜를 이용한 데이터 가져오기
Private Sub cmdTest_Click()
    
    Dim iNumber As Integer
    
    If LabResultObject.GetLabStandByListByWorkDate("20050603", "CHE") Then
        
        With Me.spdTest
            
            .MaxRows = LabResultObject.GetRowCount
            
            For iNumber = 0 To LabResultObject.GetRowCount - 1
                
                .Row = iNumber + 1
                
                .Col = 1
                .Value = LabResultObject.GetDataValue(iNumber, "LBQACPNUM")
                
                
                .Col = 2
                .Value = LabResultObject.GetDataValue(iNumber, "LABSHTNAM")
                    
                .Col = 3
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRCOD")
                
                .Col = 4
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRDAY")
                
                .Col = 5
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRTMS")
                
                 .Col = 6
                .Value = LabResultObject.GetDataValue(iNumber, "PBSRESNUM")
                
                 .Col = 7
                .Value = LabResultObject.GetDataValue(iNumber, "RSBACKSTT")
                
                
                
            Next
        
        End With
        
    End If
    
End Sub

''차트번호및 날짜를 이용한 데이터 가져오기
Private Sub cmdTest1_Click()
    
     Dim iNumber As Integer
    
    
    If LabResultObject.GetLabStandByListByChtNum("20050428", "20209") Then
        
        With Me.spdTest
            
            .MaxRows = LabResultObject.GetRowCount
            
           
            For iNumber = 0 To LabResultObject.GetRowCount - 1
                
                .Row = iNumber + 1
                
                .Col = 1
                .Value = LabResultObject.GetDataValue(iNumber, "LBQACPNUM")
                
                .Col = 2
                .Value = LabResultObject.GetDataValue(iNumber, "LABSHTNAM")
                
                .Col = 3
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRQTY")
                
                .Col = 4
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRDAY")
                
                .Col = 5
                .Value = LabResultObject.GetDataValue(iNumber, "OSPODRTMS")
                
                .Col = 6
                .Value = LabResultObject.GetDataValue(iNumber, "PbsResNum")
                
                    
            Next
        
        End With
        
    End If
    
End Sub

'''Update 테스트
Private Sub cmdTest2_Click()
    
    If LabResultObject.UpdateLabResult("215", _
                                       "1040507349", _
                                       "6", _
                                       "1", _
                                       "0", _
                                       "B1050", _
                                       "TEST", _
                                       "I", _
                                       "LAB", _
                                       "200504131530") Then
        
        MsgBox "Update 성공"
        
   End If
   
                                    
End Sub

''아이디/패스워드 체크
Private Sub Command1_Click()
    
    Dim sUserID As String
    Dim sPassWord As String
    
    sUserID = Me.Text1.Text
    sPassWord = Me.Text2.Text
    
    
    If LabResultObject.GetValidUserYon(Trim(sUserID), Trim(sPassWord), "20050413") Then
    
        MsgBox ("적절한 사용자입니다.")
    
    Else
        
        MsgBox ("적절하지 않은 사용자입니다.")
    
    End If
    
End Sub

Private Sub Command2_Click()

    If LabResultObject.BITReadConsumptionSummary("20051101", "20051102") Then
        MsgBox (LabResultObject.GetConsumptionDataValue(0, "OdrCod"))
        MsgBox (CStr(LabResultObject.GetConsumptionRowCount))
    End If

End Sub

Private Sub Form_Load()
    
    
    Set LabResultObject = New BITLabResultInterface.BITLabResultInterface
    
    '''초기화시에 아래의 함수를 반드시 호출하여 주어야 한다.
    Call LabResultObject.InitializeServer
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '''프로그램 종료시에 아래의 함수를 반드시 호출하여 주어야 한다.
    Call LabResultObject.FinalizeServer
    Set LabResultObject = Nothing
    
End Sub

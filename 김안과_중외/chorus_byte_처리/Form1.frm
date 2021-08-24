VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   17310
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   12255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   13320
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gTxbuffer(100) As Byte     '전역Buffer
Dim gInputCount As Integer     '총 받은갯수
Dim gNeedCount As Integer      '장비에서 보내겠다고 알려준 데이터 갯수
Dim gCMDCode As Byte           '장비 Comment Code

Private Sub Form_Load()
  MSComm1.PortOpen = True
  
  '전역Buffer 초기화
  InitGlobal
  
End Sub

Private Sub MSComm1_OnComm()

    '바이너리 모드로 처리!
    MSComm1.InputMode = comInputModeBinary
    
    Dim InCnt As Integer
    Dim i As Long
    Dim Str As String
    Dim Buffer() As Byte
    
    InCnt = MSComm1.InBufferCount
    Buffer = MSComm1.Input
    
    '테스트용
    Str = Byte2Str(Buffer(), InCnt)
    Debug.Print "CHORUS-> "; Str
    '
    
    For i = 0 To InCnt - 1
        gInputCount = gInputCount + 1                 '총 받은갯수
        
        gTxbuffer(gInputCount - 1) = Buffer(i)        '전역Buffer에 ADD
        
        Select Case gInputCount
            Case 2
                gNeedCount = gTxbuffer(gInputCount - 1) '장비에서 보내겠다고 알려준 데이터 갯수
            Case 3
                gCMDCode = gTxbuffer(gInputCount - 1)   '장비 Comment Code
            Case Else
                If gInputCount = gNeedCount + 3 Then  'gNeedCount + STX + CS
                    Select Case gCMDCode
                        Case 5                        'ENQ - 데이터 START
                            SendACK
                                                
                        Case 210                      'D2 - 오더요청
                            SendChorusOrder
                            
                        Case 211                      'D3 - 데이터 END
                            SendACK
                            
                        Case 215                      'D7 - 결과데이터, 한건씩 처리하려면 여기서!!
                            SendACK
                            Call DataDefine(gTxbuffer())
                        Case 216                      'D8 - 결과데이터 END, 결과를 모두 받은후에 한번에 처리하려면 여기서 해야함!!
                            SendACK
                    End Select
                    
                    Call InitGlobal                   '데이터 다 나와서 처리했으니 초기화
                End If
        End Select
    Next

    'Debug.Print "Input->" & Str
    'Debug.Print Mid(Str, 1, 1) & "->" & Asc(Mid(Str, 1, 1))
End Sub

'데이터 파싱 및 결과처리
Private Sub DataDefine(AllTxBuffer() As Byte)
    Dim s As String
    Dim i As Integer
    Dim ExamIf As String
    Dim barcode As String
    Dim ResFlag As String
    Dim ResultVal As String
    Dim Unit As String
    Dim ResultTxt As String
    
    s = Byte2Str(AllTxBuffer(), gInputCount)
    
    barcode = Trim(Mid(s, 4, 18))
    ExamIf = Trim(Mid(s, 23, 7))
    ResFlag = Trim(Mid(s, 30, 1))
    ResultVal = Trim(Mid(s, 31, 12))
    Unit = Trim(Mid(s, 43, 10))
    
    Select Case ResFlag
        Case "P"
            ResultTxt = "Positive(" & ResultVal & ")"
        Case "N"
            ResultTxt = "Positive(" & ResultVal & ")"
        Case Else
            ResultTxt = ResultVal
    End Select
    
    Text1.Text = Text1.Text & vbCrLf & "원본문자열->" & s
    Text1.Text = Text1.Text & vbCrLf & "바코드:" & barcode & ", 검사명:" & ExamIf & ", 결과:" & ResultTxt & ", 단위:" & Unit
    
End Sub

'전역변수 초기화
Private Sub InitGlobal()
    InitGlobalBuffer
    gInputCount = 0
    gNeedCount = 0
    gCMDCode = 0
End Sub

'전역Buffer 초기화
Private Sub InitGlobalBuffer()
    Dim i As Integer
    
    For i = LBound(gTxbuffer) To UBound(gTxbuffer)
        gTxbuffer(i) = 0
    Next
End Sub

Private Function Null2Space(Buffer As Byte)


End Function
'Byte 데이터를 문자열로변환.. 테스트용...
Private Function Byte2Str(ByteData() As Byte, Count As Integer) As String
    Dim s As String
    Dim i As Integer
    
    For i = 0 To Count - 1
        If ByteData(i) = 0 Then     'Null 값은 space로 변환, Text 사용하기 위해!
            s = s & " "
        Else
            s = s & Chr(ByteData(i))
        End If
    Next
    
    Byte2Str = s
    
End Function


'코러스용 응답 데이터 전송
Private Sub SendACK()
  Dim Chorus_ACK(4) As Byte
    Chorus_ACK(0) = 2   'STX
    Chorus_ACK(1) = 1   'SOH
    Chorus_ACK(2) = 4   'EOT
    Chorus_ACK(3) = 5   'ENQ: CheckSum
    
    Debug.Print "HOST-> SendACK"
    
    Call CommSendBuffer(Chorus_ACK())
End Sub

'컴포트 바이너리 Output
Private Sub CommSendBuffer(OutBuffer() As Byte)
    MSComm1.Output = OutBuffer
End Sub

'Chorus 오더 전송.. 테스트용...
Private Sub SendChorusOrder()
  Dim bSender(26) As Byte
    bSender(0) = &H2
    bSender(1) = &H17
    bSender(2) = &H4
    bSender(3) = &H53
    bSender(4) = &H61
    bSender(5) = &H6D
    bSender(6) = &H70
    bSender(7) = &H6C
    bSender(8) = &H65
    bSender(9) = &H20
    bSender(10) = &H31
    bSender(11) = &H0
    bSender(12) = &H0
    bSender(13) = &H0
    bSender(14) = &H0
    bSender(15) = &H0
    bSender(16) = &H0
    bSender(17) = &H0
    bSender(18) = &H0
    bSender(19) = &H0
    bSender(20) = &H0
    bSender(21) = &H0
    bSender(22) = &H0
    bSender(23) = &H1
    bSender(24) = &H1
    bSender(25) = &H24
                        
    Debug.Print "HOST-> SendOrder"
            
    CommSendBuffer bSender
End Sub


'BCC체크섬.. 테스트용으로만 사용..
Public Function BccCheckSum(cData As String) As Byte   ' checksum 계산
Dim iChk As Integer
Dim ix2 As Integer

  iChk = 0
  For ix2 = 1 To Len(cData)
      If ix2 = 1 Then
          If Len(cData) = 1 Then
              iChk = iChk + Asc(Mid$(cData, 1, 1)) Xor Asc(Mid$(cData, 2, 1))
          Else
              iChk = iChk + Asc(Mid$(cData, 1, 1)) Xor Asc(Mid$(cData, 2, 1))
          End If
      ElseIf _
          ix2 = 2 Then  '1번째 * 2번째 처리
      Else
          iChk = iChk Xor Asc(Mid$(cData, ix2, 1))
      End If
  Next ix2
  
  BccCheckSum = Chr(iChk)
  
End Function



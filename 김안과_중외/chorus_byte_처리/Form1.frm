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
   StartUpPosition =   3  'Windows �⺻��
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

Dim gTxbuffer(100) As Byte     '����Buffer
Dim gInputCount As Integer     '�� ��������
Dim gNeedCount As Integer      '��񿡼� �����ڴٰ� �˷��� ������ ����
Dim gCMDCode As Byte           '��� Comment Code

Private Sub Form_Load()
  MSComm1.PortOpen = True
  
  '����Buffer �ʱ�ȭ
  InitGlobal
  
End Sub

Private Sub MSComm1_OnComm()

    '���̳ʸ� ���� ó��!
    MSComm1.InputMode = comInputModeBinary
    
    Dim InCnt As Integer
    Dim i As Long
    Dim Str As String
    Dim Buffer() As Byte
    
    InCnt = MSComm1.InBufferCount
    Buffer = MSComm1.Input
    
    '�׽�Ʈ��
    Str = Byte2Str(Buffer(), InCnt)
    Debug.Print "CHORUS-> "; Str
    '
    
    For i = 0 To InCnt - 1
        gInputCount = gInputCount + 1                 '�� ��������
        
        gTxbuffer(gInputCount - 1) = Buffer(i)        '����Buffer�� ADD
        
        Select Case gInputCount
            Case 2
                gNeedCount = gTxbuffer(gInputCount - 1) '��񿡼� �����ڴٰ� �˷��� ������ ����
            Case 3
                gCMDCode = gTxbuffer(gInputCount - 1)   '��� Comment Code
            Case Else
                If gInputCount = gNeedCount + 3 Then  'gNeedCount + STX + CS
                    Select Case gCMDCode
                        Case 5                        'ENQ - ������ START
                            SendACK
                                                
                        Case 210                      'D2 - ������û
                            SendChorusOrder
                            
                        Case 211                      'D3 - ������ END
                            SendACK
                            
                        Case 215                      'D7 - ���������, �ѰǾ� ó���Ϸ��� ���⼭!!
                            SendACK
                            Call DataDefine(gTxbuffer())
                        Case 216                      'D8 - ��������� END, ����� ��� �����Ŀ� �ѹ��� ó���Ϸ��� ���⼭ �ؾ���!!
                            SendACK
                    End Select
                    
                    Call InitGlobal                   '������ �� ���ͼ� ó�������� �ʱ�ȭ
                End If
        End Select
    Next

    'Debug.Print "Input->" & Str
    'Debug.Print Mid(Str, 1, 1) & "->" & Asc(Mid(Str, 1, 1))
End Sub

'������ �Ľ� �� ���ó��
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
    
    Text1.Text = Text1.Text & vbCrLf & "�������ڿ�->" & s
    Text1.Text = Text1.Text & vbCrLf & "���ڵ�:" & barcode & ", �˻��:" & ExamIf & ", ���:" & ResultTxt & ", ����:" & Unit
    
End Sub

'�������� �ʱ�ȭ
Private Sub InitGlobal()
    InitGlobalBuffer
    gInputCount = 0
    gNeedCount = 0
    gCMDCode = 0
End Sub

'����Buffer �ʱ�ȭ
Private Sub InitGlobalBuffer()
    Dim i As Integer
    
    For i = LBound(gTxbuffer) To UBound(gTxbuffer)
        gTxbuffer(i) = 0
    Next
End Sub

Private Function Null2Space(Buffer As Byte)


End Function
'Byte �����͸� ���ڿ��κ�ȯ.. �׽�Ʈ��...
Private Function Byte2Str(ByteData() As Byte, Count As Integer) As String
    Dim s As String
    Dim i As Integer
    
    For i = 0 To Count - 1
        If ByteData(i) = 0 Then     'Null ���� space�� ��ȯ, Text ����ϱ� ����!
            s = s & " "
        Else
            s = s & Chr(ByteData(i))
        End If
    Next
    
    Byte2Str = s
    
End Function


'�ڷ����� ���� ������ ����
Private Sub SendACK()
  Dim Chorus_ACK(4) As Byte
    Chorus_ACK(0) = 2   'STX
    Chorus_ACK(1) = 1   'SOH
    Chorus_ACK(2) = 4   'EOT
    Chorus_ACK(3) = 5   'ENQ: CheckSum
    
    Debug.Print "HOST-> SendACK"
    
    Call CommSendBuffer(Chorus_ACK())
End Sub

'����Ʈ ���̳ʸ� Output
Private Sub CommSendBuffer(OutBuffer() As Byte)
    MSComm1.Output = OutBuffer
End Sub

'Chorus ���� ����.. �׽�Ʈ��...
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


'BCCüũ��.. �׽�Ʈ�����θ� ���..
Public Function BccCheckSum(cData As String) As Byte   ' checksum ���
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
          ix2 = 2 Then  '1��° * 2��° ó��
      Else
          iChk = iChk Xor Asc(Mid$(cData, ix2, 1))
      End If
  Next ix2
  
  BccCheckSum = Chr(iChk)
  
End Function



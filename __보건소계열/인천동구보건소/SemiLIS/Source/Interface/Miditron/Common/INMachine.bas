Attribute VB_Name = "ModuleMachine"
Option Explicit

Public MDIactivekey                 As Integer  'MDIform�� �̹� load �Ǿ� �ִ� ���¸� ��Ÿ���� Ű
Public OrderKey                     As Integer  '��������� �˻�� Order�� ������ ȭ��� ��� ����
Public CallLabKey                   As Integer  'Lab Manager Program�� ȣ���� ���ΰ��� ���θ� ��Ÿ���� Ű
Public SlipDigit                    As Integer  'Slip�� �� �ڸ��� �� ���ΰ� ����
Public DigitShape                   As String   'Slip�� �ڸ��� ���� ������ ���
Public FieldAddIdenTBFlag           As Integer  'Field�� IdenTB�� �߰��� ������ �ϸ� �󸶳� �� ������ ����
Public FieldAddResultTBFlag         As Integer  'Field�� ResultTB�� �߰��� ������ �ϸ� �󸶳� �� ������ ����
Public IdTBNField()
Public IdTBFieldDig()
Public IdTBFieldName()
Public RTBNField()
Public RTBFieldDig()
Public RTBFieldName()

Sub FieldADD(IdTBno As Integer, RTBno As Integer)
    
    FieldAddIdenTBFlag = IdTBno
    FieldAddResultTBFlag = RTBno
    
    If FieldAddIdenTBFlag <> 0 Then
        ReDim IdTBNField(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldDig(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldName(1 To FieldAddIdenTBFlag)
    End If
    
    If FieldAddResultTBFlag <> 0 Then
        ReDim RTBNField(1 To FieldAddResultTBFlag)
        ReDim RTBFieldDig(1 To FieldAddResultTBFlag)
        ReDim RTBFieldName(1 To FieldAddResultTBFlag)
    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBNField(1) = "timeNF"
        ''
        ''

    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldDig(1) = 6
        
        ''
        ''

    End If

    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldName(1) = "Testtime"
        
        ''
        ''

    End If
    
        
    If FieldAddResultTBFlag <> 0 Then
        RTBNField(1) = ""
        RTBNField(2) = ""
        ''
        ''

    End If
    
    If FieldAddResultTBFlag <> 0 Then
        RTBFieldDig(1) = ""
        RTBFieldDig(2) = ""
        ''
        ''

    End If

    If FieldAddResultTBFlag <> 0 Then
        RTBFieldName(1) = ""
        RTBFieldName(2) = ""
        ''
        ''

    End If

    
End Sub

Public Sub MachineConfig()
    
    Dim RetVal As Long
    Dim sBuf As String
    Dim CallLab As String
    
    Set interfacfrm = INTface41 'INTface41 ���� ��ü�ϴ� �̸�
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "MachineName", "Machine", sBuf, 255, App.Path & "\initial.ini")
    machstr = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "Title", "�˻���", sBuf, 255, App.Path & "\initial.ini")
    Title = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "LabManagerCall", "No", sBuf, 255, App.Path & "\initial.ini")
    CallLab = Left(sBuf, RetVal)
    
    ''retval = GetPrivateProfileInt("Path", "MaxItemNo", 50, App.Path & "\initial.ini")
    
    
    machinit = Left(machstr, 3) & "_"   ''"MIN_" '�ӽ� �̴ϼ��� ���ڷ�, ��� �� �����Ͽ�
                      '�ִ� 4�ڷ� ��
    
    fileInit = Left(machstr, 3)     '"MIN" '���ϸ� �̴ϼȷ�, ���� ���� machinit����
                     '��� �ٸ� ���� ��
    
        
    INTmain00.Caption = Title & " " & " " & "�������̽� �ʱ� ȭ��"
      
    FileName = App.Path & "\" '& machstr & "\"
    ''filename = App.Path & "\interfac\" & machstr & "\"
    commstr = "clinic\setcomm.mdb"
    codestr = "clinic\setcode.mdb"
        
    Call delcheck(FileName & "comm\", machstr)

'MDIactivekey�� MDIform�� �̹� Load �Ǿ� �ִ� ��(= true)�� ��Ÿ���� ����
'������ �ٸ� app.���� �ٽ� ȣ���Ͽ� Load �Ǵ� ���� ������ �� ����.       '
    
    ImgClickkey = False
    MDIactivekey = True

'######################## CONFIGURATION #######################################################################################

'Slip�� �� �ڸ��� �� ���ΰ��� ����, �ڸ��� ���� ������ ��� ����t
    SlipDigit = 4
    Call SlipDigitShape

'Order ���ο� ���� ���� - Order used(True), Not used Order(False)
    OrderKey = False
    
'Lab Manager Program�� ȣ���� ���̰��� ���� ���� - ȣ���(True), ��ȣ���(False)
    If CallLab = "Yes" Then
        CallLabKey = True
    Else
        CallLabKey = False
    End If
    
'Column�� �߰��� ���ΰ��� ���� �� �� ���� Column�� �߰��� ���ΰ� ����(0:�߰�����, 1: 1Col�߰�, 2: 2Col�߰�, ...., n: nCol�߰�)
    ''Call FieldADD(1, 0)
    
End Sub
Sub Send_Order()
    
End Sub


Sub SlipDigitShape()
    Select Case SlipDigit
    
        Case 4
            DigitShape = "0000"
        Case 5
            DigitShape = "00000"
        Case 6
            DigitShape = "000000"
        Case 7
            DigitShape = "0000000"
        Case 8
            DigitShape = "00000000"
        Case 9
            DigitShape = "000000000"
        Case 10
            DigitShape = "0000000000"
        Case 11
            DigitShape = "00000000000"
        Case 12
            DigitShape = "000000000000"
        Case 13
            DigitShape = "0000000000000"
        Case 14
            DigitShape = "00000000000000"
    
    End Select
    
    
End Sub



VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows �⺻��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim StrMsg      As String       ' �ؽ�Ʈ
    Dim IntLen      As Integer      ' �ؽ�Ʈ ����
    Dim SngLF       As Single       ' ���� ǥ�ø� ���� X��ǥ
    Dim SngTP       As Single       ' ���� ǥ�ø� ���� Y��ǥ
    Dim IntRow      As Integer      ' �����μ⿡�� �� ��ȣ
    Dim IntCol      As Integer      ' �����μ⿡�� ĭ ��ȣ
    Dim IntIdx      As Integer      ' Loop �ӽ� ����
    Dim StrChar     As String       ' ����

    
    Me.AutoRedraw = True
    Me.Font = "@����"
    Me.Font.Bold = True
    Me.ForeColor = RGB(0, 0, 128)
    
    StrMsg = ""
    StrMsg = StrMsg & "�Ź�ó�� ���ξ���..." & vbCrLf
    StrMsg = StrMsg & "VB���� �������� �ʴ� ����Դϴ�." & vbCrLf
    StrMsg = StrMsg & "�׷��ٰ� �����Ͻ÷ƴϱ�??" & vbCrLf
    StrMsg = StrMsg & "�ȵǸ� �ǰ��ؾ���~!!!" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "������ �����̴� �뵵�� �°� �����ؼ� ����ϼ���." & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "���ĺ� �ҹ��� : abcdefghijklmnopqrstuvwxyz" & vbCrLf
    StrMsg = StrMsg & "���ĺ� �빮�� : ABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCrLf
    StrMsg = StrMsg & "���� : 1234567890" & vbCrLf
    StrMsg = StrMsg & "�ѱ� : �����ٶ󸶹ٻ������īŸ����" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "2006.07.04. ����� 182cm@korea.com" & vbCrLf


    Me.Cls
    
    IntLen = Len(StrMsg)

    IntRow = 0
    IntCol = 0
    
    For IntIdx = 1 To IntLen
        StrChar = Mid(StrMsg, IntIdx, 1)
        
        Select Case StrChar
            Case vbLf
            
            Case vbCr
                ' �ٹٲ�
                IntCol = IntCol + 1
                IntRow = 0
            
            Case Else
                ' ���� ǥ��
                
                ' ���� �� ����
                SngTP = IntRow * (Me.TextWidth("��") + 0)

'                ' �����ٿ��� ���� ���� ����
'                SngLF = IntCol * (Me.TextWidth("��") + 60)      ' ���� ���� ����
                
                ' �����ٿ��� ���� ��� ����
                SngLF = (IntCol + 1) * (Me.TextWidth("��") + 60) - (Me.TextWidth(StrChar) / 3)
                
'                ' �����ٿ��� ���� ������ ����
'                SngLF = (IntCol + 1) * (Me.TextWidth("��") + 60) - Me.TextWidth(StrChar)
                
                Me.CurrentX = SngLF + 300
                Me.CurrentY = SngTP + 300
                Me.Print StrChar

                IntRow = IntRow + 1
        End Select
        
    Next
    
End Sub


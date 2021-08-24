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
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim StrMsg      As String       ' 텍스트
    Dim IntLen      As Integer      ' 텍스트 길이
    Dim SngLF       As Single       ' 글자 표시를 위한 X좌표
    Dim SngTP       As Single       ' 글자 표시를 위한 Y좌표
    Dim IntRow      As Integer      ' 세로인쇄에서 줄 번호
    Dim IntCol      As Integer      ' 세로인쇄에서 칸 번호
    Dim IntIdx      As Integer      ' Loop 임시 변수
    Dim StrChar     As String       ' 글자

    
    Me.AutoRedraw = True
    Me.Font = "@굴림"
    Me.Font.Bold = True
    Me.ForeColor = RGB(0, 0, 128)
    
    StrMsg = ""
    StrMsg = StrMsg & "신문처럼 세로쓰기..." & vbCrLf
    StrMsg = StrMsg & "VB에서 제공되지 않는 기능입니다." & vbCrLf
    StrMsg = StrMsg & "그렇다고 포기하시렵니까??" & vbCrLf
    StrMsg = StrMsg & "안되면 되게해야져~!!!" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "간단한 예제이니 용도에 맞게 응용해서 사용하세요." & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "알파벳 소문자 : abcdefghijklmnopqrstuvwxyz" & vbCrLf
    StrMsg = StrMsg & "알파벳 대문자 : ABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCrLf
    StrMsg = StrMsg & "숫자 : 1234567890" & vbCrLf
    StrMsg = StrMsg & "한글 : 가나다라마바사아자차카타파하" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "" & vbCrLf
    StrMsg = StrMsg & "2006.07.04. 용재님 182cm@korea.com" & vbCrLf


    Me.Cls
    
    IntLen = Len(StrMsg)

    IntRow = 0
    IntCol = 0
    
    For IntIdx = 1 To IntLen
        StrChar = Mid(StrMsg, IntIdx, 1)
        
        Select Case StrChar
            Case vbLf
            
            Case vbCr
                ' 줄바꿈
                IntCol = IntCol + 1
                IntRow = 0
            
            Case Else
                ' 글자 표시
                
                ' 세로 줄 간격
                SngTP = IntRow * (Me.TextWidth("가") + 0)

'                ' 세로줄에서 글자 왼쪽 맞춤
'                SngLF = IntCol * (Me.TextWidth("가") + 60)      ' 가로 간격 조절
                
                ' 세로줄에서 글자 가운데 맞춤
                SngLF = (IntCol + 1) * (Me.TextWidth("가") + 60) - (Me.TextWidth(StrChar) / 3)
                
'                ' 세로줄에서 글자 오른쪽 맞춤
'                SngLF = (IntCol + 1) * (Me.TextWidth("가") + 60) - Me.TextWidth(StrChar)
                
                Me.CurrentX = SngLF + 300
                Me.CurrentY = SngTP + 300
                Me.Print StrChar

                IntRow = IntRow + 1
        End Select
        
    Next
    
End Sub


VERSION 5.00
Begin VB.Form frmUpdate 
   Caption         =   "자동업데이트"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
'    Dim valzTmp             As Variant
'    Dim files1              As String
'    Dim iCnt                As Integer
'    Dim tmpRemark           As String
'    Dim varRemark           As Variant
'    Dim tmpPath             As String
'    Dim varVersion          As Variant
'    Dim tmpVersion2         As String
'    Dim wshshell            As Object
    
    Dim ClientVer   As String
    Dim ServerVer   As String
    Dim chkHospital         As String
    Dim sIP         As String
    
    ClientVer = ""
    ServerVer = ""
    chkHospital = ""
    
    sIP = "\\172.16.1.75\ACF\"
    
    '현재 클라이언트의 버젼 정보를 가지고 온다.
    If Not Dir(App.Path & "\version.ini") = "" Then
        Open App.Path & "\version.ini" For Input As #1
        Input #1, chkHospital
        Close #1
    End If
    ClientVer = FileVersion(chkHospital)
    chkHospital = ""
    
'MsgBox ClientVer

    '현재 배포서버의 버젼 정보를 가지고 온다.
    If Not Dir(sIP & "\version.ini") = "" Then
        Open sIP & "\version.ini" For Input As #2
        Input #2, chkHospital
        Close #2
    End If
    ServerVer = FileVersion(chkHospital)

'MsgBox ServerVer
               
    If ClientVer = ServerVer Then
       Shell App.Path & "\KDBAR.exe", vbNormalFocus '프로그램 실행
       End '업데이트창 종료
    Else
        'FileCopy 원본파일, 타겟파일
        FileCopy sIP & "\KDBAR.exe", App.Path & "\KDBAR.exe"
        

        'INI 업데이트
        Call WritePrivateProfileString("Update", "KDBAR.EXE", ServerVer, App.Path & "\version.ini")
        
        
        Shell App.Path & "\KDBAR.exe", vbNormalFocus '프로그램 실행
        End '업데이트창 종료
        
    End If
    
    
End Sub

Function FileVersion(ByVal varPara As String) As String
    Dim sFileList           As Variant
    Dim iCnt                As Integer
    
    sFileList = Split(varPara, vbLf)
    For iCnt = 0 To UBound(sFileList)
        If "KDBAR.EXE" = UCase(Split(sFileList(iCnt), "=")(0)) Then
            FileVersion = Split(sFileList(iCnt), "=")(1)
            Exit For
        End If
    
    Next

End Function

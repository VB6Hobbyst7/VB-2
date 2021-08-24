VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmErrMessage 
   Caption         =   "Error Message Box "
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmErrMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form_Error_Message_Box"
   ScaleHeight     =   3840
   ScaleWidth      =   6840
   StartUpPosition =   2  '�飛� 陛遴等
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   180
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text_View 
      BeginProperty Font 
         Name            =   "掉葡羹"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '熱霜
      TabIndex        =   0
      Top             =   60
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   9255
   End
   Begin BHButton.BHImageButton cmdOK 
      Height          =   375
      Left            =   4230
      TabIndex        =   2
      Top             =   3375
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Ok"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "掉葡"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdSaveTo 
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   3375
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "SaveTo"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "掉葡"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
End
Attribute VB_Name = "frmErrMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdSaveTo_Click()
    Dim strFileName As String
    Dim strFilePath As String
    Dim objFile     As FileSystemObject
    Dim logFile     As TextStream
    
    With dlgSave
        .CancelError = False
        .DialogTitle = "Error Message Save."
        .InitDir = App.Path
        .Filter = "Text (*.txt)|*.txt| Log (*.log)|*.log|"
        .ShowSave
        strFileName = .FileName
        strFilePath = .InitDir
    End With
    
    Screen.MousePointer = 11
    If Len(strFileName) <> 0 Then
        Set objFile = New FileSystemObject
        Set logFile = objFile.OpenTextFile(strFileName, ForAppending, True)
        Call logFile.WriteLine(vbCrLf & "[ LIMAS ERROR LOG ]")
        Call logFile.WriteLine("收收收收收收收收收收收收收收收收收收收收" & _
                               "收收收收收收收收收收收收收收收收收收收收")
        Call logFile.WriteLine(Text_View)
        Call logFile.WriteLine("收收收收收收收收收收收收收收收收收收收收" & _
                               "收收收收收收收收收收收收收收收收收收收收")
        logFile.Close
        Set objFile = Nothing
        Set logFile = Nothing
    End If
    Screen.MousePointer = 0
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call Unload(Me)
        Case vbKeyReturn
            Call cmdOk_Click
        Case Else
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ErrorLog_Save
End Sub

Private Sub Form_Resize()
    If Height < 1000 Then Exit Sub
    If Width < 1000 Then Exit Sub
    Call Text_View.Move(60, 60, ScaleWidth - 120, Height - 1000) '955)
    Call Frame1.Move(0, Text_View.Height + 105, Width)
    Call cmdOK.Move(Width - 1560, Text_View.Height + 195)
    Call cmdSaveTo.Move(Width - 3000, Text_View.Height + 195)
End Sub

Private Sub ErrorLog_Save()
    
    Dim objFile     As FileSystemObject
    Dim logFile     As TextStream
    Dim FileName    As String

On Error Resume Next
    Screen.MousePointer = 11
    Set objFile = New FileSystemObject
    
    FileName = Format(Now, "YYYYMMDD") & ".LOG"
    
    With objFile
        If Not .FolderExists(DirPath & "ErrorLog\") Then
            .CreateFolder DirPath & "ErrorLog\"
        End If
        Set logFile = .OpenTextFile(DirPath & "ErrorLog\" & FileName, ForAppending, True)
    End With
    
    Call logFile.WriteLine(Text_View)
    Call logFile.WriteLine("收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收收")
    logFile.Close
    
    Set objFile = Nothing
    Set logFile = Nothing
    Screen.MousePointer = 0
End Sub


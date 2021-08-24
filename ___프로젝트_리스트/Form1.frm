VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows 기본값
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   8850
      TabIndex        =   4
      Top             =   840
      Width           =   5085
   End
   Begin VB.DirListBox Dir2 
      Height          =   3240
      Left            =   3870
      TabIndex        =   3
      Top             =   90
      Width           =   3585
   End
   Begin VB.TextBox Text1 
      Height          =   5175
      Left            =   7680
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   6030
      TabIndex        =   1
      Top             =   3570
      Width           =   1245
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strFolder As String
    
    Text1.Text = ""
    
    For i = 0 To Dir1.ListCount
        'strFolder = strFolder & mGetP(Dir1.List(i), 3, "\") & vbNewLine
        'Dir2.Path = Dir1.List(i)
        'For j = 0 To Dir2.ListCount
            strFolder = strFolder & mGetP(Dir1.List(i), 9, "\") & vbNewLine
        'Next
    Next
    
    Text1.Text = strFolder
    
    
    Call MkDir(App.Path & "\Log")
    
    
    
'    Text1.Text = ""
'
'    For i = 0 To File1.ListCount
'        Text1.Text = Text1.Text & File1.List(i) & vbCrLf
'
'    Next
    
End Sub

Private Sub Dir1_Click()
    Dim i As Integer
    
    File1.Path = Dir1.Path
    
    Text1.Text = ""
    
    For i = 0 To File1.ListCount - 1
        'Text1.Text = Text1.Text & File1.List(i) & vbCrLf
        
        If Dir("F:\내사진\_Family\" & Mid(File1.List(i), 1, 8), vbDirectory) <> Mid(File1.List(i), 1, 8) Then
            Call MkDir("F:\내사진\_Family\" & Mid(File1.List(i), 1, 8))
        End If
        
        
        'CALL F:\문서\개인자료\My Pictures\2019-01-31까지_핸드폰백업
        
    Next
    
End Sub

Private Sub Form_Load()

    Dir1.Path = "F:\내사진\_Family"
    
End Sub

Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


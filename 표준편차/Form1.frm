VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtDev 
      Height          =   495
      Left            =   5610
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdDev 
      Caption         =   "표준편차"
      Height          =   495
      Left            =   4140
      TabIndex        =   3
      Top             =   1440
      Width           =   1365
   End
   Begin VB.TextBox txtAvr 
      Height          =   495
      Left            =   5610
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton cmdAvr 
      Caption         =   "평균"
      Height          =   495
      Left            =   4140
      TabIndex        =   1
      Top             =   660
      Width           =   1365
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   8745
      Left            =   270
      TabIndex        =   0
      Top             =   150
      Width           =   2955
      _Version        =   196608
      _ExtentX        =   5212
      _ExtentY        =   15425
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
      MaxCols         =   5
      MaxRows         =   100
      SpreadDesigner  =   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'평균값 함수
Function Mean(Arr() As Double) As Double
    Dim Sum As Double
    Dim i As Integer
    Dim k As Integer
    
    k = UBound(Arr)
    
    Sum = 0
    For i = 1 To k
        Sum = Sum + Arr(i)
    Next i
    
    Mean = Sum / k

End Function

 

'표준편차 함수
Function StdDev(Arr() As Double, ByVal avg As Double) As Double
    Dim i As Integer
    Dim SumSq As Double
    Dim k As Integer
    
    k = UBound(Arr)

    For i = 1 To k
         SumSq = SumSq + (Arr(i) - avg) ^ 2
    Next i
    
    StdDev = Sqr(SumSq / (k))


End Function

 

Private Sub cmdAvr_Click()
    Dim i As Integer
    Dim Arr() As Double
    Dim Average As Double
    Dim Std_Dev As Double
    
    
    For i = 1 To vaSpread1.DataRowCnt
        vaSpread1.Row = i
        vaSpread1.Col = 1
        ReDim Preserve Arr(i)
        Arr(i) = vaSpread1.Text
    Next
    

    Average = Mean(Arr()) '평균
    txtAvr.Text = Average

End Sub

Private Sub cmdDev_Click()
    Dim i As Integer
    Dim Arr() As Double
    Dim Std_Dev As Double
    
    
    For i = 1 To vaSpread1.DataRowCnt
        vaSpread1.Row = i
        vaSpread1.Col = 1
        ReDim Preserve Arr(i)
        Arr(i) = vaSpread1.Text
    Next

    Average = Mean(Arr()) '평균
    Std_Dev = StdDev(Arr(), Average) '표준편차
    txtDev.Text = Std_Dev
    
End Sub

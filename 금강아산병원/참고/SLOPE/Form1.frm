VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2655
      Left            =   4410
      TabIndex        =   5
      Top             =   840
      Width           =   5985
      _Version        =   393216
      _ExtentX        =   10557
      _ExtentY        =   4683
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
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   945
      Left            =   1980
      TabIndex        =   4
      Top             =   3390
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4350
      TabIndex        =   3
      Top             =   2970
      Width           =   1905
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1890
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   645
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   540
      TabIndex        =   0
      Top             =   330
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type point
    x As Long
    y As Long
End Type

Private mData(3, 3) As Double


Private Sub Command1_Click()

    Dim points(240) As point
    Me.Refresh
    Randomize
    
    'Draw axis
    Me.Line (-ScaleWidth, 0)-(ScaleWidth, 0), vbRed
    Me.Line (0, -ScaleHeight)-(0, ScaleHeight), vbRed
    For i = 0 To UBound(points) - 1
    
        If Rnd > 0.5 Then
            x = -1
        Else
            x = 1
        End If
        
        If Rnd > 0.5 Then
            y = -1
        Else
            y = 1
        End If
        
        points(i).x = CInt(ScaleWidth * Rnd) * x
        points(i).y = CInt(ScaleHeight * Rnd) * y
        Me.Circle (points(i).x, points(i).y), 2, vbBlack
        Sx = Sx + points(i).x
        Sy = Sy + points(i).y
        Sxy = Sxy + (points(i).x * points(i).y)
        Sx2 = Sx2 + points(i).x * points(i).x
    Next
    
    N = UBound(points)
    
    M = (N * Sxy - Sx * Sy) / (N * Sx2 - Sx * Sx)
    B = (Sy - M * Sx) / N
    
    Me.Line (-1000, M * -1000 + B)-(1000, M * 1000 + B)
   
End Sub

Private Sub Command2_Click()
'''''f.y.i. the algorithm that I gave you calculated the same slope that
'''''you provided in your sample data set.
'''''
'''''Here is a code snippet to do the calculation :
'''''
'''''
'''''// values needed prior to code snippet
'''''//
'''''// int n = number of (x,y) points
'''''// double x[ i ] for i = 0 to n-1 ... the x points
'''''// double y[ i ] for i = 0 to n-1 ... the y points
'''''//
''''
''''Dim i As Integer
''''
''''Dim x(3) As Double
''''Dim y(3) As Double
''''
''''Dim s0 As Double
''''Dim s1 As Double
''''Dim s2 As Double
''''Dim t0 As Double
''''Dim t1 As Double
''''
''''Dim Slope       As Double
''''Dim intercept   As Double
''''
''''x(0) = 0.602059991
''''x(1) = 0
''''x(2) = -0.602059991
''''x(3) = 0
''''
''''y(0) = -0.045998832
''''y(1) = -0.663540266
''''y(2) = -1.292429824
''''y(3) = -2.096910013
''''
''''
'''''double s0 = 0;
'''''double s1 = 0;
'''''double s2 = 0;
'''''double t0 = 0;
'''''double t1 = 0;
''''
'''''for (int i=0; i&lt;n; i++)
'''''{
'''''s0++;
'''''s1 = s1 + x[ i ];
'''''s2 = s2 + x[ i ]*x[ i ];
'''''t0 = t0 + y[ i ];
'''''t1 = t1 + x[ i ]*y[ i ];
'''''}
''''
''''For i = 0 To UBound(x)
''''    s0 = s0 + 1
''''    s1 = s1 + x(i)
''''    s2 = s2 + x(i) * x(i)
''''
''''    t0 = t0 + y(i)
''''    t1 = t1 + x(i) * y(i)
''''Next
''''
''''Slope = (s0 * s2 - s1 * s1) / (s0 * t1 - s1 * t0) '; // slope
''''
''''intercept = (s0 * s2 - s1 * s1) / (s2 * t0 - s1 * t1) '; // y-intercept
''''
'''''Debug.Print Application.WorksheetFunction.Slope(A(), B())


'f.y.i. the algorithm that I gave you calculated the same slope that
'you provided in your sample data set.
'
'Here is a code snippet to do the calculation :
'
'
'// values needed prior to code snippet
'//
'// int n = number of (x,y) points
'// double x[ i ] for i = 0 to n-1 ... the x points
'// double y[ i ] for i = 0 to n-1 ... the y points
'//

Dim i As Integer

Dim x(3) As Double
Dim y(3) As Double

Dim s0 As Double
Dim s1 As Double
Dim s2 As Double
Dim t0 As Double
Dim t1 As Double

Dim Slope       As Double
Dim intercept   As Double

x(0) = 0.602059991
x(1) = 0
x(2) = -0.602059991
x(3) = 0

y(0) = -0.045998832
y(1) = -0.663540266
y(2) = -1.292429824
y(3) = -2.096910013


y(0) = -0.32503785899555
y(1) = -1.89047544216721
y(2) = -0.332679438382517
y(3) = -1.83885107676191



x(0) = 1.386294361
x(1) = 0
x(2) = -1.386294361
x(3) = 0



'double s0 = 0;
'double s1 = 0;
'double s2 = 0;
'double t0 = 0;
'double t1 = 0;

'for (int i=0; i&lt;n; i++)
'{
's0++;
's1 = s1 + x[ i ];
's2 = s2 + x[ i ]*x[ i ];
't0 = t0 + y[ i ];
't1 = t1 + x[ i ]*y[ i ];
'}

For i = 0 To UBound(x)
    s0 = s0 + 1
    s1 = s1 + x(i)
    s2 = s2 + x(i) * x(i)
        
    t0 = t0 + y(i)
    t1 = t1 + x(i) * y(i)
Next

Slope = (s0 * t1 - s1 * t0) / (s0 * s2 - s1 * s1) '; // slope

intercept = (s2 * t0 - s1 * t1) / (s0 * s2 - s1 * s1)  '; // y-intercept

'Debug.Print Application.WorksheetFunction.Slope(A(), B())


End Sub


Private Sub Command4_Click()
    Dim points(3) As point
    
    Dim dblSlope    As Double
    
    points(0).x = 0.602059991
    points(1).x = 0
    points(2).x = -0.602059991
    points(3).x = 0

    points(0).y = -0.045998832
    points(1).y = -0.663540266
    points(2).y = -1.292429824
    points(3).y = -2.096910013

    dblSlope = Slope(points)
    
End Sub

Private Function Slope(pointArray() As point) As Double
    Dim ix As Integer
    Dim avgX As Double '= 0
    Dim avgY As Double '= 0
    Dim pointCount As Integer
    'Dim points(3) As point
    Dim pointsX(3) As Double
    Dim pointsY(3) As Double
    
    pointsX(0) = 0.602059991
    pointsX(1) = 0
    pointsX(2) = -0.602059991
    pointsX(3) = 0

    pointsY(0) = -0.045998832
    pointsY(1) = -0.663540266
    pointsY(2) = -1.292429824
    pointsY(3) = -2.096910013
    
    pointCount = UBound(pointArray) + 1
    ' first get the average X and average Y:
'    For ix = 0 To pointCount - 1
'        avgX = avgX + pointArray(ix).x
'        avgY = avgY + pointArray(ix).y
'    Next
    
    For ix = 0 To pointCount - 1
        avgX = avgX + pointsX(ix)
        avgY = avgY + pointsY(ix)
    Next
    
    avgX = avgX / pointCount
    avgY = avgY / pointCount
    ' now get the top and bottom sums:
    Dim topSum As Double '= 0.0
    Dim bottomSum As Double '= 0.0
    
    
    For ix = 0 To pointCount - 1
        Dim xdiff As Double
        xdiff = pointsX(ix) - avgX
        topSum = topSum + (xdiff * (pointsY(ix) - avgY))
        bottomSum = bottomSum + (xdiff * xdiff)
    Next
    Slope = topSum / bottomSum

End Function


Private Sub SaveExcel()

'On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow    As Integer
Dim iCol    As Integer
Dim i       As Integer
Dim dblVal   As Double
Dim dblVal1   As Double
Dim dblVal2   As Double

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
  '  Set xlBook = xlapp.Workbooks.Add
    
  '  Set xlSheet = xlBook.Worksheets(1)
        
    i = 0
 
Dim A(), B(), x As Long
A() = Array(1.386294361, 0, -1.386294361)
B() = Array(0.086636307, -1.052683357, -1.701005106)
 
 Call vaSpread1.SetText(1, 1, "1.386294361")
 Call vaSpread1.SetText(1, 2, "0")
 Call vaSpread1.SetText(1, 3, "-1.386294361")
' Call vaSpread1.SetText(1, 4, "-2.096910013")
 
 

 
 Call vaSpread1.SetText(2, 1, "0.086636307")
 Call vaSpread1.SetText(2, 2, "-1.052683357")
 Call vaSpread1.SetText(2, 3, "-1.701005106")
' Call vaSpread1.SetText(2, 4, "0")
 




' Dim x
' Dim y
 
 x = -0.045998832
 
 Call vaSpread1.SetText(3, 1, xlapp.WorksheetFunction.LinEst(B(), A()))
 
'X = 4
'ReDim B(X)
 
 'For i = 1 To X
 '  B(i) = i
 'Next i
 
 
'X(0) = 0.602059991
'X(1) = 0
'X(2) = -0.602059991
'X(3) = 0
'
'y(0) = -0.045998832
'y(1) = -0.663540266
'y(2) = -1.292429824
'y(3) = -2.096910013

 dblVal1 = xlapp.WorksheetFunction.Slope(B(), A())
 dblVal2 = xlapp.WorksheetFunction.Slope(A(), B())
 
 dblVal1 = xlapp.WorksheetFunction.LinEst(A(), B())
 dblVal2 = xlapp.WorksheetFunction.LinEst(B(), B())
    
    dblVal = dblVal2 - dblVal1
    
    dblVal = xlapp.WorksheetFunction.intercept(A(), B())
'    xlBook.SaveAs (Filename)
'    xlapp.Quit


End Sub



Private Sub Command5_Click()

    Call SaveExcel
    
End Sub

'Private Sub Command3_Click()
'
'mData(0, 0) = (1.4, 2.0)
'
'= {{1.4, 2.0}, {1.8, 3.9}, {1.9, 4.1}, {2.3, 4.5}, {2.5, 4.9}}
'
'
''Sorry about the 'stutters' but here are a couple of Slope functions (in VB.Net) that work from paired x,y, data. If you are looking at a trend, substitute 1,2,3... for the X points.
'
'
'    'Each assumes a 2 dimensional array with X values
'    ' in the first column and Y values in the second column
'
'    Public Function Slope2(ByVal arr(,) As Double) As Double
'        Dim SXiYi As Double
'        Dim SXi As Double
'        Dim SYi As Double
'        Dim SXi2 As Double
'        Dim i As Integer
'        Dim N As Integer
'        Dim Sxy As Double
'        Dim Sxx As Double
'        For i = arr.GetLowerBound(0) To arr.GetUpperBound(0)
'            SXiYi += arr(i, 0) * arr(i, 1)
'            SXi += arr(i, 0)
'            SYi += arr(i, 1)
'            SXi2 += arr(i, 0) ^ 2
'        Next
'        N = arr.GetUpperBound(0) - arr.GetLowerBound(0) + 1
'        Sxy = SXiYi - ((SXi * SYi) / N)
'        Sxx = SXi2 - ((SXi ^ 2) / N)
'        Return SXY / SXX
'    End Function
'    Public Function Slope1(ByVal arr(,) As Double) As Double
'        Dim XMean As Double
'        Dim YMean As Double
'        Dim Sx As Double
'        Dim Sy As Double
'        Dim Sxy As Double
'        Dim N As Double
'        Dim Sxx As Double
'        Dim i As Integer
'        For i = arr.GetLowerBound(0) To arr.GetUpperBound(0)
'            Sx += arr(i, 0)
'            Sy += arr(i, 1)
'        Next
'        N = arr.GetUpperBound(0) - arr.GetLowerBound(0) + 1
'        XMean = Sx / N
'        YMean = Sy / N
'        For i = arr.GetLowerBound(0) To arr.GetUpperBound(0)
'            Sxy += (arr(i, 0) - XMean) * (arr(i, 1) - YMean)
'            Sxx += (arr(i, 0) - XMean) ^ 2
'        Next
'        Return Sxy / Sxx
'    End Function
'End Sub

Private Sub Form_Load()
    ScaleMode = 0
    ScaleLeft = -1000
    ScaleWidth = 2000
    ScaleTop = 1000
    ScaleHeight = -2000
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print x & "," & y
End Sub

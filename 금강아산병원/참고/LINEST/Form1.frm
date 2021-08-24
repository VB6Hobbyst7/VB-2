VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1155
      Left            =   840
      TabIndex        =   0
      Top             =   1020
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINT
    x As Double
    y As Double
End Type


Private Function Trend(Data() As POINT, ByVal Degree As Long) As POINT()

    'degree 1 = straight line y=a+bx
    'degree n = polynomials!!
    
    Dim a() As Double
    Dim Ai() As Double
    Dim B() As Double
    Dim P() As Double
    Dim SigmaA() As Double
    Dim SigmaP() As Double
    Dim PointCount As Long
    Dim MaxTerm As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long
    Dim Ret() As POINT
    
    Degree = Degree + 1
    
    MaxTerm = (2 * (Degree - 1))
    PointCount = UBound(Data) + 1
    
    ReDim SigmaA(MaxTerm - 1)
    ReDim SigmaP(MaxTerm - 1)
    
    ' Get the coefficients lists for matrices A, and P
    For m = 0 To (MaxTerm - 1)
        For n = 0 To (PointCount - 1)
            SigmaA(m) = SigmaA(m) + (Data(n).x ^ (m + 1))
            SigmaP(m) = SigmaP(m) + ((Data(n).x ^ m) * Data(n).y)
        Next
    Next
    
    ' Create Matrix A, and fill in the coefficients
    ReDim a(Degree - 1, Degree - 1)
    For i = 0 To (Degree - 1)
        For j = 0 To (Degree - 1)
            If i = 0 And j = 0 Then
                a(i, j) = PointCount
            Else
               a(i, j) = SigmaA((i + j) - 1)
            End If
        Next
    Next
    
    ' Create Matrix P, and fill in the coefficients
    ReDim P(Degree - 1, 0)
    For i = 0 To (Degree - 1)
        P(i, 0) = SigmaP(i)
    Next
    
    ' We have A, and P of AB=P, so we can solve B because B=AiP
    Ai = MxInverse(a)
    B = MxMultiplyCV(Ai, P)
    
    ' Now we solve the equations and generate the list of points
    PointCount = PointCount - 1
    ReDim Ret(PointCount)
    
    ' Work out non exponential first term
    For i = 0 To PointCount
        Ret(i).x = Data(i).x
        Ret(i).y = B(0, 0)
    Next
    
    ' Work out other exponential terms including exp 1
    For i = 0 To PointCount
        For j = 1 To Degree - 1
            Ret(i).y = Ret(i).y + (B(j, 0) * Ret(i).x ^ j)
        Next
    Next
    
    Trend = Ret
    
End Function

Public Function MxMultiplyCV(Matrix1() As Double, ColumnVector() As Double) As Double()

    Dim i As Long
    Dim j As Long
    Dim Rows As Long
    Dim Cols As Long
    Dim Ret() As Double
    
    Rows = UBound(Matrix1, 1)
    Cols = UBound(Matrix1, 2)
    
    ReDim Ret(UBound(ColumnVector, 1), 0) 'returns a column vector
    
    For i = 0 To Rows
        For j = 0 To Cols
            Ret(i, 0) = Ret(i, 0) + (Matrix1(i, j) * ColumnVector(j, 0))
        Next
    Next
    
    MxMultiplyCV = Ret
    
End Function

Public Function MxInverse(Matrix() As Double) As Double()
    
    Dim i As Long
    Dim j As Long
    Dim Rows As Long
    Dim Cols As Long
    Dim Tmp() As Double
    Dim Ret() As Double
    Dim Degree As Long
    
    Tmp = Matrix
    
    Rows = UBound(Tmp, 1)
    Cols = UBound(Tmp, 2)
    Degree = Cols + 1
    
    'Augment Identity matrix onto matrix M to get [M|I]
    ReDim Preserve Tmp(Rows, (Degree * 2) - 1)
    For i = Degree To (Degree * 2) - 1
        Tmp(i Mod Degree, i) = 1
    Next
    
    ' Now find the inverse using Gauss-Jordan Elimination which should get us [I|A-1]
    MxGaussJordan Tmp
    
    ' Copy the inverse (A-1) part to array to return
    ReDim Ret(Rows, Cols)
    For i = 0 To Rows
        For j = Degree To (Degree * 2) - 1
            Ret(i, j - Degree) = Tmp(i, j)
        Next
    Next
    
    MxInverse = Ret
    
End Function

Public Sub MxGaussJordan(Matrix() As Double)
    
    Dim Rows As Long
    Dim Cols As Long
    Dim P As Long
    Dim i As Long
    Dim j As Long
    Dim m As Double
    Dim d As Double
    Dim Pivot As Double
    
    Rows = UBound(Matrix, 1)
    Cols = UBound(Matrix, 2)

    ' Reduce so we get the leading diagonal
    For P = 0 To Rows
        Pivot = Matrix(P, P)
        For i = 0 To Rows
            If Not P = i Then
                m = Matrix(i, P) / Pivot
                For j = 0 To Cols
                    Matrix(i, j) = Matrix(i, j) + (Matrix(P, j) * -m)
                Next
            End If
        Next
    Next
    
    'Divide through to get the identity matrix
    'Note: the identity matrix may have very small values (close to zero)
    'because of the way floating points are stored.
    For i = 0 To Rows
        d = Matrix(i, i)
        For j = 0 To Cols
            Matrix(i, j) = Matrix(i, j) / d
        Next
    Next
    
End Sub

Private Sub Command1_Click()
    Dim points(2) As POINT
    
    
'0.086636307 1.386294361
    
'-1.052683357    0
    
'-1.701005106    -1.386294361
    
    
    
    points(0).x = 1.386294361
    points(1).x = 0
    points(2).x = -1.386294361
'    points(3).x = 0

    points(0).y = 0.086636307
    points(1).y = -1.052683357
    points(2).y = -1.701005106
'    points(3).y = -2.096910013

    Call Trend(points, 3)
End Sub

Private Sub Form_Load()

End Sub

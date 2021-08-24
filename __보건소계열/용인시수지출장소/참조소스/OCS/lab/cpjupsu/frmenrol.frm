VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmEnrol 
   Caption         =   "접수item 화면"
   ClientHeight    =   8055
   ClientLeft      =   4980
   ClientTop       =   750
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   6870
   Begin FPSpreadADO.fpSpread ssEnrol 
      Height          =   6045
      Left            =   135
      TabIndex        =   0
      Top             =   1440
      Width           =   6540
      _Version        =   196608
      _ExtentX        =   11536
      _ExtentY        =   10663
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   2
      SpreadDesigner  =   "frmEnrol.frx":0000
      Appearance      =   1
      ScrollBarTrack  =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmEnrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Order_Data
    sJeobsuDt       As String
    nSLipno1        As Integer
    nJeobsuT1       As Integer
    nJeobsuT2       As Integer
    sPtno           As String
    sSex            As String
    nAgeYY          As Integer
    nAgeMM          As Integer
    sIndate         As String
    sRoomCode       As String
    sDeptCode       As String
    sGBio           As String
    sBi             As String
    sGbER           As String
    sGeomchCD       As String
    sGeomsaGu       As String
    sOrderDt        As String
    nQuantity       As Integer
    sCmDoctor       As String
    sDrCode         As String
    sJeobsuYn       As String
    sGbinfo         As String
    sItemCD         As String
    sCollDate       As String
    sCollHH         As String
    sColMM          As String
    sJeobsu_Lab     As String
End Type
Dim sEorder     As Order_Data

Private Sub Form_Load()
    Dim sPtno       As String
    Dim sJeobsuDt   As String
    Dim sDept       As String
    Dim sName       As String
    Dim sRowID      As String
    Dim sEitem      As String
    
    
    ssEnrol.RowHeight(-1) = 11
    Me.Top = 1
    Me.Left = (Screen.Width - Me.Width)
    
    
    GoSub Variable_Vinding
    GoSub Order_Data_Expand
    
    For i = 1 To ssEnrol.DataRowCnt
        ssEnrol.Row = i
        ssEnrol.Col = 3: sEitem = ssEnrol.Text
        ssEnrol.Col = 2: ssEnrol.Text = GET_SLipname(Left(sEitem, 2))
    Next
    Exit Sub
    
    
Variable_Vinding:
    frmMain.ssOrder.Row = nRow(0)
    frmMain.ssOrder.Col = 3:  sRowID = frmMain.ssOrder.Text
    frmMain.ssOrder.Col = 4:  sJeobsuDt = frmMain.ssOrder.Text
    frmMain.ssOrder.Col = 5:  sPtno = frmMain.ssOrder.Text
    frmMain.ssOrder.Col = 6:  sName = frmMain.ssOrder.Text
    frmMain.ssOrder.Col = 15: sDept = frmMain.ssOrder.Text
    Return
    
Order_Data_Expand:
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_Order"
    strSQL = strSQL & " WHERE  JeobsuDt  =  TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSQL = strSQL & " AND    Ptno      =  '" & sPtno & "'"
    strSQL = strSQL & " AND    DeptCode  =  '" & sDept & "'"
    
    ssEnrol.MaxRows = 0
    If False = adoSetOpen(strSQL, adoSet) Then Return
    ssEnrol.MaxRows = 100
    
    Do Until adoSet.EOF
        ssEnrol.Row = ssEnrol.DataRowCnt + 1
        
        ssEnrol.Col = 3: ssEnrol.Text = adoSet.Fields("iTemCD").Value & ""
        
        If IsRoutineCode(adoSet.Fields("iTemCD").Value & "") Then
            ssEnrol.Col = 4: ssEnrol.Text = Get_RoutineName(adoSet.Fields("iTemcd").Value & "")
            ssEnrol.Col = 1: ssEnrol.Text = "*"
            GoSub Get_RoutinCode_Data
        Else
            ssEnrol.Col = 4: ssEnrol.Text = Get_itemName(adoSet.Fields("iTemcd").Value & "")
        End If
        
        ssEnrol.Col = 5: ssEnrol.Text = adoSet.Fields("GeomchCd").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
Get_RoutinCode_Data:
    Dim sRcode      As String
    Dim adoRt       As adodb.Recordset
    
    sRcode = adoSet.Fields("itemCd").Value & ""
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, b.itemnm"
    strSQL = strSQL & " FROM   TWEXAM_ROUTINE a,"
    strSQL = strSQL & "        TWEXAM_Itemml  b "
    strSQL = strSQL & " WHERE  a.ROUTINCD = '" & sRcode & "'"
    strSQL = strSQL & " AND    a.CODEKY   =  b.codeky(+)"
    
    If False = adoSetOpen(strSQL, adoRt) Then Return
    
    Do Until adoRt.EOF
        ssEnrol.Row = ssEnrol.DataRowCnt + 1
        
        ssEnrol.Col = 3: ssEnrol.Text = adoRt.Fields("Codeky").Value & ""
        ssEnrol.Col = 4: ssEnrol.Text = "  " & Trim(adoRt.Fields("itemnm").Value & "")
        adoRt.MoveNext
    Loop
    Call adoSetClose(adoRt)
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

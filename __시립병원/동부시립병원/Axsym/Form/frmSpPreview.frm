VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmSpPreview 
   Caption         =   "미리보기 창"
   ClientHeight    =   8790
   ClientLeft      =   195
   ClientTop       =   495
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8790
   ScaleMode       =   0  '사용자
   ScaleWidth      =   10110
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread vaSpread1 
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   9795
      _Version        =   393216
      _ExtentX        =   17277
      _ExtentY        =   1085
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
      SpreadDesigner  =   "frmSpPreview.frx":0000
   End
   Begin FPSpreadADO.fpSpreadPreview vaSpreadPreview1 
      Height          =   7305
      Left            =   90
      TabIndex        =   0
      Top             =   690
      Width           =   9765
      _Version        =   393216
      _ExtentX        =   17224
      _ExtentY        =   12885
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
End
Attribute VB_Name = "frmSpPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
    '** 출력폼 선택(폼 이름)==============================================
    Select Case Gbl_FormName
        Case "frmResult"
            Me.vaSpreadPreview1.hWndSpread = frmResult.spdRstDetail.hwnd
            
            '-- Title Setting(제목, 날짜 등)
            Call UpDate_Title_Consum001
        
        Case "frmStatistics"
            Me.vaSpreadPreview1.hWndSpread = frmStatistics.spdResult1.hwnd
            
            '-- Title Setting(제목, 날짜 등)
            Call UpDate_Title_Consum000
            
'        Case "frmTestEqp"
'            Me.vaSpreadPreview1.hWndSpread = frmTestEqp.spdTestListDt.hwnd
            
            '-- Title Setting(제목, 날짜 등)
            Call UpDate_Title_Consum002
        
        Case Else
    
    End Select
    '=====================================================================
    
    'Update page count listing
    Call UpDatePageCount
    
End Sub

Private Sub Form_Load()
   
    SetupToolbar
    
    'Disable Previous button
    DisableButton 4, "LEFT"
        
    'Get the zoom display
    'GetZoom zoomindex
    
    Select Case Gbl_FormName
        Case "frmResult" '-- 검사결과
            If frmResult.spdRstDetail.PrintPageCount = 1 Then
                'Disable Next button if only one page
                DisableButton 2, "LEFT"
            End If
            
            frmResult.spdRstDetail.PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape
        
        Case "frmStatistics" '-- 통계
            If frmStatistics.spdResult1.PrintPageCount = 1 Then
                'Disable Next button if only one page
                DisableButton 2, "LEFT"
            End If
            
            frmStatistics.spdResult1.PrintOrientation = PrintOrientationLandscape
            
'        Case "frmTestEqp" '-- 검사항목
'            If frmTestEqp.spdTestListDt.PrintPageCount = 1 Then
'                'Disable Next button if only one page
'                DisableButton 2, "LEFT"
'            End If
'
'            frmTestEqp.spdTestListDt.PrintOrientation = PrintOrientationLandscape ' PrintOrientationPortrait
            
        Case Else
    
    End Select
    '=====================================================================
End Sub

Private Sub SetupToolbar()
Dim I As Integer
    
    With vaSpread1
        'Specify whether Edit Mode is to remain on when switching between cells
        .EditModePermanent = True
    
        .Col = -1
        .Row = -1
        .Lock = True
        
        'Set the number of rows in the spreadsheet
        .MaxRows = 1
     
        'Set the height of a selected row
        .RowHeight(1) = 15
       
        'Set the number of columns in the spreadsheet
        .MaxCols = 17
     
        'Set the column widths
        For I = 1 To .MaxCols Step 2
            .ColWidth(I) = 0.3
        Next I
       
        'Resize wide column
        .ColWidth(14) = 15
        
        'Show or hide the column headers
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        
        'Turn off scroll bars
        .ScrollBars = ScrollBarsNone
        
        'Turn off border
        .BorderStyle = BorderStyleNone
          
        'Select row(s)
        .Row = 1
        .Col = -1
    
        'Determine the color of background, foreground and border color
        .ForeColor = RGB(0, 0, 0)
        .BackColor = RGB(192, 192, 192)
        .FontName = "MS Sans Serif"
        .FontSize = 8
        .Fontbold = False
        
        'Select a single cell
        .Col = 2
        .Row = 1
    
        'Define cells as type BUTTON
        .CellType = CellTypeButton
        .Lock = False
        .TypeButtonText = "Next"
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\RIGHT.BMP")
        .TypeButtonAlign = TypeButtonAlignLeft
        
        'Select a single cell
        .Col = 4
        .Row = 1
    
        'Define cells as type BUTTON
        .CellType = CellTypeButton
        .Lock = False
        .TypeButtonText = "Previous"
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\LEFT.BMP")
        .TypeButtonAlign = TypeButtonAlignRight
        
        'Select a single cell
        .Col = 6
        .Row = 1
    
        'Define cells as type BUTTON
'        .CellType = CellTypeButton
'        .Lock = False
'        .TypeButtonText = "Zoom"
'        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\ZOOM.BMP")
'        .TypeButtonAlign = TypeButtonAlignRight
        
        'Select a single cell
        .Col = 8
        .Row = 1
    
        'Define cells as type BUTTON
        .CellType = CellTypeButton
        .Lock = False
        .TypeButtonText = "Print"
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\PRINT.BMP")
        .TypeButtonAlign = TypeButtonAlignRight
        
        'Select a single cell
        .Col = 10
        .Row = 1
    
        'Define cells as type BUTTON
'        .CellType = CellTypeButton
'        .Lock = False
'        .TypeButtonText = "Setup"
'        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\SETUP.BMP")
'        .TypeButtonAlign = TypeButtonAlignRight
        
        
        'Select a single cell
        .Col = 16
        .Row = 1
    
        'Define cells as type BUTTON
        .CellType = CellTypeButton
        .Lock = False
        .TypeButtonText = "Close"
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\CLOSE.BMP")
        .TypeButtonAlign = TypeButtonAlignRight
        .TextTip = TextTipFloating
        Dim bRet As Boolean
        bRet = .SetTextTipAppearance("MS Sans Serif", 8, 0, 0, &HC0FFFF, &H0)
        .CursorType = CursorTypeLockedCell
        .CursorStyle = CursorStyleArrow
        .NoBeep = True
    End With
    
End Sub

Private Sub DisableButton(Col As Long, bitmapdirection As String)
    With vaSpread1
        'Disable specified button
        .ReDraw = False
        
        .Row = 1
        .Col = Col
        
        .Lock = True
        .TypeButtonTextColor = RGB(128, 128, 128)
        .Protect = True
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\" & bitmapdirection & "DIS.BMP")
        
        .ReDraw = True
    End With
End Sub

Private Sub EnableButton(Col As Long, bitmapdirection As String)
    With vaSpread1
        'Enable specified button
        .ReDraw = False
        
        .Row = 1
        .Col = Col
        
        .Lock = False
        .TypeButtonTextColor = RGB(0, 0, 0)
        .Protect = False
        Set .TypeButtonPicture = LoadPicture(App.Path & "\Image\" & bitmapdirection & ".BMP")
        
        .ReDraw = True
    End With
End Sub

Private Sub Form_Resize()
    vaSpread1.Move 0, 0, ScaleWidth, vaSpread1.Height
    vaSpreadPreview1.Move 0, vaSpread1.Height, ScaleWidth, ScaleHeight - vaSpread1.Height
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    vaSpread1.Col = Col
    vaSpread1.Row = Row
    
    If vaSpread1.CellType = CellTypeButton Then
        Select Case Col
            Case 2  'Next
                '** 출력폼 선택(폼 이름)==============================================
                Select Case Gbl_FormName
                    Case "frmResult"
                        With frmResult
                            If vaSpreadPreview1.PageCurrent < frmResult.spdRstDetail.PrintPageCount Then
                                vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent + vaSpreadPreview1.PagesPerScreen
                                EnableButton Col, "RIGHT"
                                'Enable Previous button
                                EnableButton 4, "LEFT"
                               'Update page count listing
                                UpDatePageCount
                            End If
                            
                             'If at last page, disable button
                            If vaSpreadPreview1.PageCurrent > frmResult.spdRstDetail.PrintPageCount - vaSpreadPreview1.PagesPerScreen Then
                                DisableButton Col, "RIGHT"
                            End If
                        End With
                        
                    Case "frmStatistics"
                        With frmStatistics
                            If vaSpreadPreview1.PageCurrent < frmStatistics.spdResult1.PrintPageCount Then
                                vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent + vaSpreadPreview1.PagesPerScreen
                                EnableButton Col, "RIGHT"
                                'Enable Previous button
                                EnableButton 4, "LEFT"
                               'Update page count listing
                                UpDatePageCount
                            End If
                            
                             'If at last page, disable button
                            If vaSpreadPreview1.PageCurrent > frmStatistics.spdResult1.PrintPageCount - vaSpreadPreview1.PagesPerScreen Then
                                DisableButton Col, "RIGHT"
                            End If
                        End With
                    
'                    Case "frmTestEqp"
'                        With frmTestEqp
'                            If vaSpreadPreview1.PageCurrent < frmTestEqp.spdTestListDt.PrintPageCount Then
'                                vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent + vaSpreadPreview1.PagesPerScreen
'                                EnableButton Col, "RIGHT"
'                                'Enable Previous button
'                                EnableButton 4, "LEFT"
'                               'Update page count listing
'                                UpDatePageCount
'                            End If
'
'                             'If at last page, disable button
'                            If vaSpreadPreview1.PageCurrent > frmTestEqp.spdTestListDt.PrintPageCount - vaSpreadPreview1.PagesPerScreen Then
'                                DisableButton Col, "RIGHT"
'                            End If
'                        End With
                    
                    Case Else

                End Select
                '=====================================================================
             
            Case 4  'Previous
                If vaSpreadPreview1.PageCurrent > 1 Then
                    vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent - vaSpreadPreview1.PagesPerScreen
                    EnableButton Col, "LEFT"
                    EnableButton 2, "RIGHT"
                    'Update page count listing
                    UpDatePageCount
                End If
                
                'If at first page, disable button
                If vaSpreadPreview1.PageCurrent = 1 Then
                    DisableButton Col, "LEFT"
                End If
                
            Case 6  'Zoom
                vaSpreadPreview1.ZoomState = 3
                
            Case 8  'Print
                PrintDlg.Show
                                 
            Case 10 'Setup
                'PageSetup.Show 1
             
            Case 16 'Close
                Unload Me
        End Select
    End If
End Sub

Private Sub UpDate_Title_Consum000()
    Dim strTitle    As String
    Dim strBDate    As String
    Dim strPage     As String
    Dim strPDate    As String
    Dim strArea     As String
    
    With SpPrint
        strTitle = .strTitle
        strBDate = .strBaseDate
        strPage = .strPageCount
        strArea = .strAreaName
        strPDate = .strPrintDate
    End With
    
    With frmStatistics.spdResult1
        .PrintHeader = strTitle & strBDate & strPage & strArea & strPDate
    End With

End Sub

Private Sub UpDate_Title_Consum001()
    Dim strTitle    As String
    Dim strBDate    As String
    Dim strPage     As String
    Dim strPDate    As String
    Dim strArea     As String
    Dim strSpcNo    As String
    Dim strName     As String
    'Dim strSex      As String
    'Dim strAge      As String
    Dim strGbn      As String
    Dim strChart    As String

        
    With SpPrint
        strTitle = .strTitle
        strBDate = .strBaseDate
        strPage = .strPageCount
        strArea = .strAreaName
        strPDate = .strPrintDate
        strSpcNo = .strSpcNo
        strName = .strName
        'strSex = .strSex
        'strAge = .strAge
        strGbn = .strGbn
        strChart = .strChart
    End With
    
    With frmResult.spdRstDetail
        .PrintHeader = strTitle & strBDate & strPage & strArea & strPDate & strSpcNo & strName & strGbn & strChart
    End With

End Sub

Private Sub UpDate_Title_Consum002()
    Dim strTitle    As String
    Dim strBDate    As String
    Dim strPage     As String
    Dim strPDate    As String
    Dim strArea     As String
    
    With SpPrint
        strTitle = .strTitle
        strBDate = .strBaseDate
        strPage = .strPageCount
        strArea = .strAreaName
        strPDate = .strPrintDate
    End With
    
'    With frmTestEqp.spdTestListDt
'        .PrintHeader = strTitle & strBDate & strPage & strArea & strPDate
'    End With

End Sub

Private Sub UpDatePageCount()
    '** 출력폼 선택(폼 이름)==============================================
    Select Case Gbl_FormName
        Case "frmResult"
            With vaSpread1
                'Page Count
                .Row = 1
                .Col = 14
                .Text = "Page " & vaSpreadPreview1.PageCurrent & " of " & frmResult.spdRstDetail.PrintPageCount
            End With
            
        Case "frmStatistics"
            With vaSpread1
                'Page Count
                .Row = 1
                .Col = 14
                .Text = "Page " & vaSpreadPreview1.PageCurrent & " of " & frmStatistics.spdResult1.PrintPageCount
            End With
            
'        Case "frmTestEqp"
'            With vaSpread1
'                'Page Count
'                .Row = 1
'                .Col = 14
'                .Text = "Page " & vaSpreadPreview1.PageCurrent & " of " & frmTestEqp.spdTestListDt.PrintPageCount
'            End With
            
        Case Else
            
    End Select
    '=====================================================================
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    With vaSpread1
        .Col = Col
        .Row = Row
        If .CellType = CellTypeButton And Not .Lock Then
            ShowTip = True
            TipText = .TypeButtonText
        ElseIf .CellType = CellTypeEdit And .Text <> "" Then
            ShowTip = True
            TipText = .Text
        End If
    End With
End Sub


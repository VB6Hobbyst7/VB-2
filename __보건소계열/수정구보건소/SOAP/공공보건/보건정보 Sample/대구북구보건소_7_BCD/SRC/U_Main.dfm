object F_Main: TF_Main
  Left = 219
  Top = 160
  Width = 947
  Height = 786
  ActiveControl = gdIf
  Caption = '  SANSOFT ABL Interface Program'
  Color = clBtnFace
  Font.Charset = HANGEUL_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #44404#47548#52404
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Visible = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 939
    Height = 616
    Align = alClient
    BevelOuter = bvNone
    Caption = 'Panel2'
    TabOrder = 0
    object gdIf: TAdvStringGrid
      Left = 0
      Top = 0
      Width = 939
      Height = 479
      Cursor = crDefault
      Align = alClient
      Color = clWhite
      ColCount = 5
      Ctl3D = False
      DefaultColWidth = 50
      DefaultRowHeight = 21
      DefaultDrawing = False
      FixedColor = 16763594
      FixedCols = 0
      RowCount = 2
      FixedRows = 1
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      GridLineWidth = 1
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing]
      ParentCtl3D = False
      ParentFont = False
      ParentShowHint = False
      ScrollBars = ssBoth
      ShowHint = True
      TabOrder = 0
      GridLineColor = clSilver
      ActiveCellShow = False
      ActiveCellFont.Charset = ANSI_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -12
      ActiveCellFont.Name = #44404#47548#52404
      ActiveCellFont.Style = [fsBold]
      ActiveCellColor = clGray
      Bands.PrimaryColor = 15663103
      Bands.PrimaryLength = 1
      Bands.SecondaryColor = clWhite
      Bands.SecondaryLength = 1
      Bands.Print = True
      AutoNumAlign = False
      AutoSize = False
      VAlignment = vtaCenter
      EnhTextSize = False
      EnhRowColMove = True
      SizeWithForm = False
      Multilinecells = False
      OnGetCellColor = gdIfGetCellColor
      OnGetAlignment = gdIfGetAlignment
      OnClickCell = gdIfClickCell
      OnDblClickCell = gdIfDblClickCell
      OnCanEditCell = gdIfCanEditCell
      DragDropSettings.OleAcceptFiles = True
      DragDropSettings.OleAcceptText = True
      SortSettings.AutoColumnMerge = False
      SortSettings.Column = 0
      SortSettings.Show = False
      SortSettings.IndexShow = False
      SortSettings.IndexColor = clYellow
      SortSettings.Full = True
      SortSettings.SingleColumn = False
      SortSettings.IgnoreBlanks = False
      SortSettings.BlankPos = blFirst
      SortSettings.AutoFormat = True
      SortSettings.Direction = sdAscending
      SortSettings.InitSortDirection = sdAscending
      SortSettings.FixedCols = False
      SortSettings.NormalCellsOnly = False
      SortSettings.Row = 0
      SortSettings.UndoSort = False
      FloatingFooter.Color = clBtnFace
      FloatingFooter.Column = 0
      FloatingFooter.FooterStyle = fsFixedLastRow
      FloatingFooter.Visible = False
      ControlLook.Color = clBlack
      ControlLook.CheckSize = 15
      ControlLook.RadioSize = 10
      ControlLook.ControlStyle = csWinXP
      ControlLook.DropDownAlwaysVisible = False
      ControlLook.ProgressMarginX = 2
      ControlLook.ProgressMarginY = 2
      EnableBlink = False
      EnableHTML = True
      EnableWheel = True
      Flat = True
      Look = glXP
      HintColor = clInfoBk
      SelectionColor = 15387318
      SelectionTextColor = clBlack
      SelectionRectangle = False
      SelectionResizer = False
      SelectionRTFKeep = False
      HintShowCells = False
      HintShowLargeText = False
      HintShowSizing = False
      PrintSettings.FooterSize = 0
      PrintSettings.HeaderSize = 0
      PrintSettings.Time = ppNone
      PrintSettings.Date = ppNone
      PrintSettings.DateFormat = 'dd/mm/yyyy'
      PrintSettings.PageNr = ppNone
      PrintSettings.Title = ppNone
      PrintSettings.Font.Charset = DEFAULT_CHARSET
      PrintSettings.Font.Color = clWindowText
      PrintSettings.Font.Height = -11
      PrintSettings.Font.Name = 'MS Sans Serif'
      PrintSettings.Font.Style = []
      PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
      PrintSettings.FixedFont.Color = clWindowText
      PrintSettings.FixedFont.Height = -11
      PrintSettings.FixedFont.Name = 'MS Sans Serif'
      PrintSettings.FixedFont.Style = []
      PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
      PrintSettings.HeaderFont.Color = clWindowText
      PrintSettings.HeaderFont.Height = -11
      PrintSettings.HeaderFont.Name = 'MS Sans Serif'
      PrintSettings.HeaderFont.Style = []
      PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
      PrintSettings.FooterFont.Color = clWindowText
      PrintSettings.FooterFont.Height = -11
      PrintSettings.FooterFont.Name = 'MS Sans Serif'
      PrintSettings.FooterFont.Style = []
      PrintSettings.Borders = pbSingle
      PrintSettings.BorderStyle = psSolid
      PrintSettings.Centered = True
      PrintSettings.RepeatFixedRows = False
      PrintSettings.RepeatFixedCols = False
      PrintSettings.LeftSize = 0
      PrintSettings.RightSize = 0
      PrintSettings.ColumnSpacing = 0
      PrintSettings.RowSpacing = 0
      PrintSettings.TitleSpacing = 0
      PrintSettings.Orientation = poPortrait
      PrintSettings.PageNumberOffset = 0
      PrintSettings.MaxPagesOffset = 0
      PrintSettings.FixedWidth = 0
      PrintSettings.FixedHeight = 0
      PrintSettings.UseFixedHeight = False
      PrintSettings.UseFixedWidth = False
      PrintSettings.FitToPage = fpNever
      PrintSettings.PageNumSep = '/'
      PrintSettings.NoAutoSize = False
      PrintSettings.NoAutoSizeRow = False
      PrintSettings.PrintGraphics = False
      PrintSettings.UseDisplayFont = True
      HTMLSettings.Width = 100
      HTMLSettings.XHTML = False
      Navigation.AlwaysEdit = True
      Navigation.AdvanceOnEnter = True
      Navigation.AdvanceDirection = adTopBottom
      Navigation.AdvanceAuto = True
      Navigation.InsertPosition = pInsertBefore
      Navigation.HomeEndKey = heFirstLastColumn
      Navigation.TabToNextAtEnd = False
      Navigation.TabAdvanceDirection = adLeftRight
      ColumnSize.Location = clRegistry
      CellNode.Color = clSilver
      CellNode.ExpandOne = False
      CellNode.NodeColor = clBlack
      CellNode.NodeIndent = 12
      CellNode.ShowTree = True
      CellNode.TreeColor = clSilver
      MaxEditLength = 0
      Grouping.HeaderColor = clNone
      Grouping.HeaderColorTo = clNone
      Grouping.HeaderTextColor = clNone
      Grouping.MergeHeader = False
      Grouping.MergeSummary = False
      Grouping.Summary = False
      Grouping.SummaryColor = clNone
      Grouping.SummaryColorTo = clNone
      Grouping.SummaryTextColor = clNone
      IntelliPan = ipVertical
      URLColor = clBlue
      URLShow = False
      URLFull = False
      URLEdit = False
      ScrollType = ssNormal
      ScrollColor = clNone
      ScrollWidth = 16
      ScrollSynch = False
      ScrollProportional = False
      ScrollHints = shNone
      OemConvert = False
      FixedFooters = 0
      FixedRightCols = 0
      FixedColWidth = 22
      FixedRowHeight = 25
      FixedFont.Charset = ANSI_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -12
      FixedFont.Name = 'Arial'
      FixedFont.Style = [fsBold]
      FixedAsButtons = False
      FloatFormat = '%.2f'
      IntegralHeight = False
      WordWrap = True
      ColumnHeaders.Strings = (
        ''
        #48148#53076#46300
        #54872#51088#48264#54840
        #54872#51088#49457#47749
        #49345#53468)
      Lookup = False
      LookupCaseSensitive = False
      LookupHistory = False
      BackGround.Top = 0
      BackGround.Left = 0
      BackGround.Display = bdTile
      BackGround.Cells = bcNormal
      Filter = <>
      ColWidths = (
        22
        109
        90
        103
        80)
      RowHeights = (
        25
        21)
      object pnPort: TPanel
        Left = 298
        Top = 226
        Width = 399
        Height = 105
        BevelOuter = bvNone
        Caption = 'ComPort Error! '
        Color = clBlack
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -32
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 2
        Visible = False
      end
      object pnBcd: TPanel
        Left = 74
        Top = 76
        Width = 195
        Height = 123
        TabOrder = 3
        Visible = False
        OnMouseMove = pnBcdMouseMove
        object lbRow: TLabel
          Left = 172
          Top = 92
          Width = 30
          Height = 12
          Caption = 'lbRow'
          Visible = False
        end
        object Panel6: TPanel
          Left = 1
          Top = 1
          Width = 193
          Height = 26
          Align = alTop
          BevelInner = bvLowered
          Caption = #48148#53076#46300' '#48320#44221
          Color = clMoneyGreen
          TabOrder = 0
          OnMouseMove = Panel6MouseMove
        end
        object edOld: TEdit
          Left = 66
          Top = 30
          Width = 117
          Height = 18
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          TabOrder = 1
        end
        object edNew: TEdit
          Left = 66
          Top = 58
          Width = 117
          Height = 18
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          TabOrder = 2
        end
        object btnBcdChange: TButton
          Left = 24
          Top = 86
          Width = 71
          Height = 27
          Caption = #51200#51109
          TabOrder = 3
          OnClick = btnBcdChangeClick
        end
        object btnBcdClose: TButton
          Left = 98
          Top = 86
          Width = 67
          Height = 27
          Caption = #45803#44592
          TabOrder = 4
          OnClick = btnBcdCloseClick
        end
        object Panel7: TPanel
          Left = 5
          Top = 31
          Width = 60
          Height = 21
          BevelInner = bvLowered
          Caption = #48320#44221#51204
          Color = clMoneyGreen
          TabOrder = 5
        end
        object Panel8: TPanel
          Left = 5
          Top = 60
          Width = 60
          Height = 21
          BevelInner = bvLowered
          Caption = #48320#44221#54980
          Color = clMoneyGreen
          TabOrder = 6
        end
      end
    end
    object Panel1: TPanel
      Left = 0
      Top = 479
      Width = 939
      Height = 137
      Align = alBottom
      TabOrder = 1
      object Panel3: TPanel
        Left = 1
        Top = 1
        Width = 582
        Height = 135
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object GroupBox1: TGroupBox
          Left = 0
          Top = 10
          Width = 582
          Height = 125
          Align = alBottom
          Caption = #44208#44284#47196#44536
          TabOrder = 0
          object mmView: TMemo
            Left = 2
            Top = 14
            Width = 578
            Height = 109
            Hint = #45908#53364#53364#47533' -> '#47196#44536#49325#51228
            Align = alClient
            Color = 13172735
            Font.Charset = HANGEUL_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            ParentFont = False
            ParentShowHint = False
            ScrollBars = ssBoth
            ShowHint = True
            TabOrder = 0
            OnDblClick = mmViewDblClick
          end
        end
      end
      object Panel5: TPanel
        Left = 583
        Top = 1
        Width = 355
        Height = 135
        Align = alRight
        BevelOuter = bvNone
        TabOrder = 1
        object gdResult: TAdvStringGrid
          Left = 0
          Top = 0
          Width = 355
          Height = 135
          Cursor = crDefault
          Align = alClient
          ColCount = 7
          Ctl3D = False
          DefaultRowHeight = 21
          DefaultDrawing = False
          FixedColor = 16763620
          FixedCols = 0
          RowCount = 6
          FixedRows = 1
          Font.Charset = HANGEUL_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = #44404#47548#52404
          Font.Style = [fsBold]
          GridLineWidth = 1
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          GridLineColor = clSilver
          ActiveCellShow = False
          ActiveCellFont.Charset = DEFAULT_CHARSET
          ActiveCellFont.Color = clWindowText
          ActiveCellFont.Height = -11
          ActiveCellFont.Name = 'Tahoma'
          ActiveCellFont.Style = [fsBold]
          ActiveCellColor = clGray
          Bands.PrimaryColor = clInfoBk
          Bands.PrimaryLength = 1
          Bands.SecondaryColor = clWindow
          Bands.SecondaryLength = 1
          Bands.Print = False
          AutoNumAlign = False
          AutoSize = False
          VAlignment = vtaTop
          EnhTextSize = False
          EnhRowColMove = True
          SizeWithForm = False
          Multilinecells = False
          OnGetCellColor = gdResultGetCellColor
          DragDropSettings.OleAcceptFiles = True
          DragDropSettings.OleAcceptText = True
          SortSettings.AutoColumnMerge = False
          SortSettings.Column = 0
          SortSettings.Show = False
          SortSettings.IndexShow = False
          SortSettings.IndexColor = clYellow
          SortSettings.Full = True
          SortSettings.SingleColumn = False
          SortSettings.IgnoreBlanks = False
          SortSettings.BlankPos = blFirst
          SortSettings.AutoFormat = True
          SortSettings.Direction = sdAscending
          SortSettings.InitSortDirection = sdAscending
          SortSettings.FixedCols = False
          SortSettings.NormalCellsOnly = False
          SortSettings.Row = 0
          SortSettings.UndoSort = False
          FloatingFooter.Color = clBtnFace
          FloatingFooter.Column = 0
          FloatingFooter.FooterStyle = fsFixedLastRow
          FloatingFooter.Visible = False
          ControlLook.Color = clBlack
          ControlLook.CheckSize = 15
          ControlLook.RadioSize = 10
          ControlLook.ControlStyle = csWinXP
          ControlLook.DropDownAlwaysVisible = False
          ControlLook.ProgressMarginX = 2
          ControlLook.ProgressMarginY = 2
          EnableBlink = False
          EnableHTML = True
          EnableWheel = True
          Flat = True
          Look = glXP
          HintColor = clInfoBk
          SelectionColor = 15387318
          SelectionTextColor = clBlack
          SelectionRectangle = False
          SelectionResizer = False
          SelectionRTFKeep = False
          HintShowCells = False
          HintShowLargeText = False
          HintShowSizing = False
          PrintSettings.FooterSize = 0
          PrintSettings.HeaderSize = 0
          PrintSettings.Time = ppNone
          PrintSettings.Date = ppNone
          PrintSettings.DateFormat = 'dd/mm/yyyy'
          PrintSettings.PageNr = ppNone
          PrintSettings.Title = ppNone
          PrintSettings.Font.Charset = DEFAULT_CHARSET
          PrintSettings.Font.Color = clWindowText
          PrintSettings.Font.Height = -11
          PrintSettings.Font.Name = 'MS Sans Serif'
          PrintSettings.Font.Style = []
          PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
          PrintSettings.FixedFont.Color = clWindowText
          PrintSettings.FixedFont.Height = -11
          PrintSettings.FixedFont.Name = 'MS Sans Serif'
          PrintSettings.FixedFont.Style = []
          PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
          PrintSettings.HeaderFont.Color = clWindowText
          PrintSettings.HeaderFont.Height = -11
          PrintSettings.HeaderFont.Name = 'MS Sans Serif'
          PrintSettings.HeaderFont.Style = []
          PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
          PrintSettings.FooterFont.Color = clWindowText
          PrintSettings.FooterFont.Height = -11
          PrintSettings.FooterFont.Name = 'MS Sans Serif'
          PrintSettings.FooterFont.Style = []
          PrintSettings.Borders = pbSingle
          PrintSettings.BorderStyle = psSolid
          PrintSettings.Centered = True
          PrintSettings.RepeatFixedRows = False
          PrintSettings.RepeatFixedCols = False
          PrintSettings.LeftSize = 0
          PrintSettings.RightSize = 0
          PrintSettings.ColumnSpacing = 0
          PrintSettings.RowSpacing = 0
          PrintSettings.TitleSpacing = 0
          PrintSettings.Orientation = poPortrait
          PrintSettings.PageNumberOffset = 0
          PrintSettings.MaxPagesOffset = 0
          PrintSettings.FixedWidth = 0
          PrintSettings.FixedHeight = 0
          PrintSettings.UseFixedHeight = False
          PrintSettings.UseFixedWidth = False
          PrintSettings.FitToPage = fpNever
          PrintSettings.PageNumSep = '/'
          PrintSettings.NoAutoSize = False
          PrintSettings.NoAutoSizeRow = False
          PrintSettings.PrintGraphics = False
          PrintSettings.UseDisplayFont = True
          HTMLSettings.Width = 100
          HTMLSettings.XHTML = False
          Navigation.AdvanceDirection = adLeftRight
          Navigation.InsertPosition = pInsertBefore
          Navigation.HomeEndKey = heFirstLastColumn
          Navigation.TabToNextAtEnd = False
          Navigation.TabAdvanceDirection = adLeftRight
          ColumnSize.Location = clRegistry
          CellNode.Color = clSilver
          CellNode.ExpandOne = False
          CellNode.NodeColor = clBlack
          CellNode.NodeIndent = 12
          CellNode.ShowTree = True
          CellNode.TreeColor = clSilver
          MaxEditLength = 0
          Grouping.HeaderColor = clNone
          Grouping.HeaderColorTo = clNone
          Grouping.HeaderTextColor = clNone
          Grouping.MergeHeader = False
          Grouping.MergeSummary = False
          Grouping.Summary = False
          Grouping.SummaryColor = clNone
          Grouping.SummaryColorTo = clNone
          Grouping.SummaryTextColor = clNone
          IntelliPan = ipVertical
          URLColor = clBlue
          URLShow = False
          URLFull = False
          URLEdit = False
          ScrollType = ssNormal
          ScrollColor = clNone
          ScrollWidth = 16
          ScrollSynch = False
          ScrollProportional = False
          ScrollHints = shNone
          OemConvert = False
          FixedFooters = 0
          FixedRightCols = 0
          FixedColWidth = 6
          FixedRowHeight = 21
          FixedFont.Charset = DEFAULT_CHARSET
          FixedFont.Color = clWindowText
          FixedFont.Height = -11
          FixedFont.Name = 'Tahoma'
          FixedFont.Style = [fsBold]
          FixedAsButtons = False
          FloatFormat = '%.2f'
          IntegralHeight = False
          WordWrap = True
          ColumnHeaders.Strings = (
            ''
            #44160#49324#47749
            #44208#44284
            #44160#49324#47749
            #44208#44284
            #44160#49324#47749
            #44208#44284)
          Lookup = False
          LookupCaseSensitive = False
          LookupHistory = False
          BackGround.Top = 0
          BackGround.Left = 0
          BackGround.Display = bdTile
          BackGround.Cells = bcNormal
          Filter = <>
          ColWidths = (
            6
            60
            50
            60
            50
            60
            50)
        end
      end
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 713
    Width = 939
    Height = 19
    Panels = <
      item
        Width = 500
      end>
  end
  object pnLog: TPanel
    Left = 0
    Top = 616
    Width = 939
    Height = 97
    Align = alBottom
    TabOrder = 2
    Visible = False
    object mmTemp: TMemo
      Left = 1
      Top = 54
      Width = 937
      Height = 42
      Align = alBottom
      ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
      TabOrder = 0
      WordWrap = False
    end
    object mmLog: TMemo
      Left = 1
      Top = 1
      Width = 937
      Height = 53
      Align = alClient
      ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
      TabOrder = 1
      WordWrap = False
    end
  end
  object btnTest: TButton
    Left = 95
    Top = 2
    Width = 33
    Height = 21
    Caption = 'TEST'
    TabOrder = 3
    Visible = False
    OnClick = btnTestClick
  end
  object pnSvr: TPanel
    Left = 302
    Top = 118
    Width = 399
    Height = 105
    BevelOuter = bvNone
    Caption = #49436#48260#51217#49549' '#50724#47448
    Color = clBlack
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -32
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 4
    Visible = False
  end
  object MainMenu1: TMainMenu
    AutoHotkeys = maManual
    Left = 348
    Top = 48
    object N4: TMenuItem
      Caption = #51089#50629
      object N6: TMenuItem
        Caption = #49440#53469#51204#49569
        OnClick = N6Click
      end
      object CLEAR1: TMenuItem
        Caption = #54868#47732'CLEAR'
        OnClick = CLEAR1Click
      end
    end
    object N1: TMenuItem
      Caption = #54872#44221#49444#51221
      object N1_1: TMenuItem
        Caption = #44160#49324#53076#46300' '#51077#47141
        OnClick = N1_1Click
      end
      object N1_4: TMenuItem
        Caption = #54252#53944' '#49444#51221
        OnClick = N1_4Click
      end
    end
    object L1: TMenuItem
      Caption = #44592#53440
      object DEBUG1: TMenuItem
        AutoCheck = True
        Caption = 'DEBUG'
        OnClick = DEBUG1Click
      end
      object Rcv1: TMenuItem
        Caption = 'Rcv'#53580#49828#53944
        OnClick = Rcv1Click
      end
      object N7: TMenuItem
        Caption = #49436#48260#53580#49828#53944
        OnClick = N7Click
      end
      object N2: TMenuItem
        Caption = '-'
      end
      object N3: TMenuItem
        Caption = #51333#47308
        OnClick = N3Click
      end
    end
  end
  object ComPort1: TComPort
    BaudRate = br9600
    Port = 'COM1'
    Parity.Bits = prNone
    StopBits = sbOneStopBit
    DataBits = dbEight
    Events = [evRxChar, evTxEmpty, evRxFlag, evRing, evBreak, evCTS, evDSR, evError, evRLSD, evRx80Full]
    FlowControl.OutCTSFlow = False
    FlowControl.OutDSRFlow = False
    FlowControl.ControlDTR = dtrEnable
    FlowControl.ControlRTS = rtsEnable
    FlowControl.XonXoffOut = False
    FlowControl.XonXoffIn = False
    OnRxChar = ComPort1RxChar
    Left = 376
    Top = 48
  end
end

object F_Find: TF_Find
  Left = 160
  Top = 185
  Width = 1096
  Height = 637
  Caption = #44208#44284' '#44288#47532
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel2: TPanel
    Left = 0
    Top = 581
    Width = 1088
    Height = 29
    Align = alBottom
    TabOrder = 0
    object Label2: TLabel
      Left = 40
      Top = 6
      Width = 89
      Height = 18
      AutoSize = False
      Caption = '[ '#47700#49464#51648' ]'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #44404#47548#52404
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lbMsg: TLabel
      Left = 132
      Top = 6
      Width = 68
      Height = 16
      Caption = #51221#49345#52376#47532
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #44404#47548#52404
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 113
    Width = 1088
    Height = 468
    Align = alClient
    Caption = 'Panel3'
    TabOrder = 1
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 624
      Height = 466
      Align = alLeft
      Caption = 'Panel4'
      TabOrder = 0
      object gdMaster: TAdvStringGrid
        Left = 1
        Top = 1
        Width = 622
        Height = 464
        Cursor = crDefault
        Align = alClient
        ColCount = 11
        DefaultRowHeight = 21
        DefaultDrawing = False
        FixedCols = 0
        RowCount = 2
        FixedRows = 1
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        GridLineWidth = 1
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing]
        ParentFont = False
        ParentShowHint = False
        ScrollBars = ssBoth
        ShowHint = True
        TabOrder = 0
        OnClick = gdMasterClick
        OnDblClick = gdMasterDblClick
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
        OnGetCellColor = gdMasterGetCellColor
        OnGetAlignment = gdMasterGetAlignment
        OnGridHint = gdMasterGridHint
        OnClickCell = gdMasterClickCell
        OnCanEditCell = gdMasterCanEditCell
        DragDropSettings.OleAcceptFiles = True
        DragDropSettings.OleAcceptText = True
        SortSettings.AutoColumnMerge = False
        SortSettings.Column = 0
        SortSettings.Show = True
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
        Flat = False
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
        ColumnSize.Stretch = True
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
        FixedColWidth = 19
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
          #44160#49324#51068#51088
          'No'
          #44160#52404#48264#54840
          #54872#51088#48264#54840
          'Flag'
          'SID'
          'ERR'
          'LotNo'
          'Location'
          #44160#49324#51088
          #51109#48708#53076#46300)
        Lookup = False
        LookupCaseSensitive = False
        LookupHistory = False
        BackGround.Top = 0
        BackGround.Left = 0
        BackGround.Display = bdTile
        BackGround.Cells = bcNormal
        Filter = <>
        ColWidths = (
          19
          135
          34
          90
          71
          29
          28
          30
          57
          56
          68)
        object pnData: TPanel
          Left = 232
          Top = 56
          Width = 313
          Height = 257
          Color = 16765650
          TabOrder = 2
          Visible = False
          OnMouseMove = pnDataMouseMove
          object Panel13: TPanel
            Left = 16
            Top = 15
            Width = 65
            Height = 17
            Caption = #44160#49324#51068#51088
            Color = 14155775
            TabOrder = 0
          end
          object Panel14: TPanel
            Left = 16
            Top = 41
            Width = 65
            Height = 17
            Caption = #44160#49324'SEQ'
            Color = 14155775
            TabOrder = 1
          end
          object Panel15: TPanel
            Left = 16
            Top = 66
            Width = 65
            Height = 17
            Caption = #44160#52404#48264#54840
            Color = 14155775
            TabOrder = 2
          end
          object Panel16: TPanel
            Left = 16
            Top = 90
            Width = 65
            Height = 17
            Caption = #46321#47197#48264#54840
            Color = 14155775
            TabOrder = 3
          end
          object Panel17: TPanel
            Left = 16
            Top = 115
            Width = 65
            Height = 17
            Caption = 'Location'
            Color = 14155775
            TabOrder = 4
          end
          object Panel18: TPanel
            Left = 16
            Top = 138
            Width = 65
            Height = 17
            Caption = #44160#49324'Flag'
            Color = 14155775
            TabOrder = 5
          end
          object Panel19: TPanel
            Left = 16
            Top = 162
            Width = 65
            Height = 17
            Caption = 'UserID'
            Color = 14155775
            TabOrder = 6
          end
          object edSendSpcid: TEdit
            Left = 84
            Top = 63
            Width = 121
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            MaxLength = 11
            ParentFont = False
            TabOrder = 7
            OnKeyDown = edSendSpcidKeyDown
            OnKeyPress = edSendSpcidKeyPress
          end
          object edSendPatId: TEdit
            Left = 84
            Top = 87
            Width = 97
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            MaxLength = 8
            ParentFont = False
            TabOrder = 8
            OnKeyDown = edSendPatIdKeyDown
            OnKeyPress = edSendPatIdKeyPress
          end
          object edLocaion: TEdit
            Left = 84
            Top = 111
            Width = 81
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            ParentFont = False
            TabOrder = 9
            OnKeyDown = edLocaionKeyDown
            OnKeyPress = edLocaionKeyPress
          end
          object edFlag: TEdit
            Left = 84
            Top = 135
            Width = 33
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            ParentFont = False
            TabOrder = 10
            OnChange = edFlagChange
            OnKeyDown = edFlagKeyDown
            OnKeyPress = edFlagKeyPress
          end
          object edUId: TEdit
            Left = 84
            Top = 159
            Width = 81
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            ParentFont = False
            TabOrder = 11
            OnKeyDown = edUIdKeyDown
            OnKeyPress = edUIdKeyPress
          end
          object pnDatetime: TPanel
            Left = 84
            Top = 14
            Width = 129
            Height = 22
            BevelInner = bvLowered
            Color = clWhite
            TabOrder = 16
          end
          object pnSeq: TPanel
            Left = 84
            Top = 39
            Width = 36
            Height = 20
            BevelInner = bvLowered
            Color = clWhite
            TabOrder = 17
          end
          object ProgressBar1: TProgressBar
            Left = 16
            Top = 211
            Width = 281
            Height = 16
            TabOrder = 18
          end
          object pnCheck: TPanel
            Left = 16
            Top = 230
            Width = 65
            Height = 20
            Caption = #45936#51060#53552#52404#53356
            TabOrder = 19
          end
          object pnOrdCreate: TPanel
            Left = 88
            Top = 230
            Width = 65
            Height = 20
            Caption = #50724#45908#49373#49457
            TabOrder = 20
          end
          object pnExcept: TPanel
            Left = 157
            Top = 230
            Width = 65
            Height = 20
            Caption = #51217#49688
            TabOrder = 21
          end
          object pnUpload: TPanel
            Left = 229
            Top = 230
            Width = 65
            Height = 20
            Caption = #44208#44284#46321#47197
            TabOrder = 22
          end
          object pnOrdCode: TPanel
            Left = 119
            Top = 136
            Width = 85
            Height = 20
            BevelInner = bvLowered
            Color = clWhite
            TabOrder = 23
          end
          object btnSave: TBitBtn
            Left = 217
            Top = 72
            Width = 85
            Height = 56
            Caption = #51200#51109
            TabOrder = 13
            OnClick = btnSaveClick
          end
          object btnSend: TBitBtn
            Left = 217
            Top = 14
            Width = 85
            Height = 56
            Caption = #51116#51204#49569
            TabOrder = 14
            OnClick = btnSendClick
          end
          object btnClosePanel: TBitBtn
            Left = 217
            Top = 130
            Width = 85
            Height = 53
            Caption = #45803#44592
            TabOrder = 15
            OnClick = btnClosePanelClick
          end
          object Panel20: TPanel
            Left = 16
            Top = 186
            Width = 65
            Height = 17
            Caption = 'LotNo'
            Color = 14155775
            TabOrder = 24
          end
          object edLotNo: TEdit
            Left = 84
            Top = 183
            Width = 81
            Height = 24
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #44404#47548#52404
            Font.Style = [fsBold]
            ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
            ParentFont = False
            TabOrder = 12
            OnKeyDown = edLotNoKeyDown
            OnKeyPress = edLotNoKeyPress
          end
          object pnICode: TPanel
            Left = 167
            Top = 184
            Width = 85
            Height = 20
            BevelInner = bvLowered
            Color = clWhite
            TabOrder = 25
            Visible = False
          end
        end
      end
    end
    object Panel5: TPanel
      Left = 625
      Top = 1
      Width = 462
      Height = 466
      Align = alClient
      Caption = 'Panel5'
      TabOrder = 1
      object gdResult: TAdvStringGrid
        Left = 1
        Top = 1
        Width = 460
        Height = 464
        Cursor = crDefault
        Align = alClient
        ColCount = 11
        DefaultRowHeight = 21
        DefaultDrawing = False
        FixedCols = 0
        RowCount = 2
        FixedRows = 1
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        GridLineWidth = 1
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
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
        OnGetAlignment = gdResultGetAlignment
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
        Flat = False
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
        FixedColWidth = 58
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
          #44160#49324#53076#46300
          #44160#49324#47749
          #44208#44284
          #52280#44256#52824'('#54616')'
          #52280#44256#52824'('#49345')'
          'D'
          'P'
          'C'
          'LH'
          'UPCODE')
        Lookup = False
        LookupCaseSensitive = False
        LookupHistory = False
        BackGround.Top = 0
        BackGround.Left = 0
        BackGround.Display = bdTile
        BackGround.Cells = bcNormal
        Filter = <>
        ColWidths = (
          58
          119
          72
          68
          71
          23
          20
          21
          27
          64
          64)
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1088
    Height = 113
    Align = alTop
    TabOrder = 2
    object GroupBox1: TGroupBox
      Left = 46
      Top = 8
      Width = 493
      Height = 97
      Caption = #44208#44284' '#51204#49569
      TabOrder = 0
      object btnFind: TSpeedButton
        Left = 63
        Top = 51
        Width = 109
        Height = 37
        Caption = #44160#49353
        OnClick = btnFindClick
      end
      object btnDel: TSpeedButton
        Left = 184
        Top = 50
        Width = 109
        Height = 37
        Caption = #49325#51228
        OnClick = btnDelClick
      end
      object btnClose: TSpeedButton
        Left = 305
        Top = 50
        Width = 109
        Height = 37
        Caption = #45803#44592
        OnClick = btnCloseClick
      end
      object Panel10: TPanel
        Left = 29
        Top = 18
        Width = 88
        Height = 24
        Caption = #44160#52404#48264#54840
        Color = 16777207
        TabOrder = 0
      end
      object Panel11: TPanel
        Left = 245
        Top = 17
        Width = 88
        Height = 24
        Caption = #46321#47197#48264#54840
        Color = 16777207
        TabOrder = 1
      end
      object edBarCode: TEdit
        Left = 119
        Top = 18
        Width = 103
        Height = 24
        Color = 15269887
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 2
      end
      object edPatNo: TEdit
        Left = 335
        Top = 15
        Width = 103
        Height = 24
        Color = 15269887
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ParentFont = False
        TabOrder = 3
      end
      object rbtBcd: TRadioButton
        Left = 223
        Top = 21
        Width = 17
        Height = 17
        Checked = True
        TabOrder = 4
        TabStop = True
      end
      object rbtPat: TRadioButton
        Left = 439
        Top = 17
        Width = 17
        Height = 17
        TabOrder = 5
      end
    end
    object GroupBox2: TGroupBox
      Left = 544
      Top = 8
      Width = 497
      Height = 97
      Caption = #44208#44284' '#51312#54924
      TabOrder = 1
      object Label1: TLabel
        Left = 192
        Top = 22
        Width = 7
        Height = 13
        Caption = '~'
      end
      object btnView: TSpeedButton
        Left = 376
        Top = 16
        Width = 105
        Height = 73
        Caption = #51312#54924
        OnClick = btnViewClick
      end
      object Panel6: TPanel
        Left = 8
        Top = 18
        Width = 80
        Height = 20
        BevelInner = bvLowered
        Caption = #44160#49324#51068#51088
        Color = 9226104
        TabOrder = 0
      end
      object dtpFrom: TDateTimePicker
        Left = 93
        Top = 17
        Width = 94
        Height = 21
        Date = 39594.718523761580000000
        Time = 39594.718523761580000000
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        TabOrder = 1
      end
      object dtpTo: TDateTimePicker
        Left = 205
        Top = 16
        Width = 97
        Height = 21
        Date = 39594.718523761580000000
        Time = 39594.718523761580000000
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        TabOrder = 2
      end
      object Panel7: TPanel
        Left = 10
        Top = 42
        Width = 78
        Height = 20
        BevelInner = bvLowered
        Caption = 'ID '#44396#48516' '#51312#54924
        Color = 9226104
        TabOrder = 3
      end
      object cmbxID: TComboBox
        Left = 10
        Top = 65
        Width = 85
        Height = 21
        Color = 15532031
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 4
        Text = #51204#52404
        Items.Strings = (
          #51204#52404
          'Error'
          'BarCode'
          'POCT'
          'QC')
      end
      object Panel8: TPanel
        Left = 185
        Top = 42
        Width = 82
        Height = 20
        BevelInner = bvLowered
        Caption = #51204#49569' '#50640#47084
        Color = 9226104
        TabOrder = 5
      end
      object cmbxSend: TComboBox
        Left = 188
        Top = 64
        Width = 90
        Height = 21
        Color = 15532031
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 6
        Text = #51204#52404
        Items.Strings = (
          #51204#52404
          #51204#52404#50640#47084
          'Complete'
          'N:BarCode'#50640#47084
          'P:'#54872#51088#48264#54840' '#50640#47084
          'U:User'#50640#47084
          'F:Flag'#50640#47084
          'X:'#51204#49569#50640#47084)
      end
      object Panel9: TPanel
        Left = 98
        Top = 42
        Width = 78
        Height = 20
        BevelInner = bvLowered
        Caption = #44208#44284' '#51060#49345
        Color = 9226104
        TabOrder = 7
      end
      object cmbxResult: TComboBox
        Left = 99
        Top = 65
        Width = 86
        Height = 21
        Color = 15532031
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 8
        Text = #51204#52404
        Items.Strings = (
          #51204#52404
          #45944#53440#54056#45769
          #51221#49345#44160#52404)
      end
      object Panel12: TPanel
        Left = 277
        Top = 42
        Width = 79
        Height = 20
        BevelInner = bvLowered
        Caption = #44160#49324#48512#49436
        Color = 9226104
        TabOrder = 9
      end
      object cmbxLoc: TComboBox
        Left = 279
        Top = 64
        Width = 90
        Height = 21
        Color = 15532031
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        ItemHeight = 13
        TabOrder = 10
        Text = #51204#52404
        Items.Strings = (
          #51204#52404)
      end
    end
    object ckbxAll: TCheckBox
      Left = 8
      Top = 89
      Width = 33
      Height = 17
      Caption = 'All'
      TabOrder = 2
      OnClick = ckbxAllClick
    end
  end
end

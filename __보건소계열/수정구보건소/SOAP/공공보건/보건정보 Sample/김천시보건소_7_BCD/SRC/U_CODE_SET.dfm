object F_CodeSet: TF_CodeSet
  Left = 343
  Top = 261
  Width = 535
  Height = 398
  Caption = 'F_CodeSet'
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 527
    Height = 43
    Align = alTop
    TabOrder = 0
    object btnClose: TSpeedButton
      Left = 144
      Top = 5
      Width = 81
      Height = 33
      Caption = #45803#44592
      OnClick = btnCloseClick
    end
    object btnView: TSpeedButton
      Left = 58
      Top = 5
      Width = 81
      Height = 33
      Caption = #51312#54924
      OnClick = btnViewClick
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 43
    Width = 527
    Height = 328
    Align = alClient
    Caption = 'Panel3'
    TabOrder = 1
    object gdCodeM: TAdvStringGrid
      Left = 1
      Top = 1
      Width = 267
      Height = 326
      Cursor = crDefault
      Align = alClient
      ColCount = 8
      DefaultRowHeight = 21
      DefaultDrawing = False
      FixedCols = 0
      RowCount = 2
      FixedRows = 1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      GridLineWidth = 1
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 0
      OnClick = gdCodeMClick
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
      OnClickCell = gdCodeMClickCell
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
      FixedColWidth = 37
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
        #49692#48264
        #44160#49324#53076#46300
        #44160#49324#47749
        #44160#49324#50557#50612
        #51204#49569#53076#46300
        'SUB'#53076#46300
        #52280#44256#52824'('#54616')'
        #52280#44256#52824'('#49345')')
      Lookup = False
      LookupCaseSensitive = False
      LookupHistory = False
      BackGround.Top = 0
      BackGround.Left = 0
      BackGround.Display = bdTile
      BackGround.Cells = bcNormal
      Filter = <>
      ColWidths = (
        37
        64
        156
        79
        60
        69
        69
        42)
    end
    object Panel4: TPanel
      Left = 268
      Top = 1
      Width = 258
      Height = 326
      Align = alRight
      TabOrder = 1
      object btnAdd: TSpeedButton
        Left = 26
        Top = 31
        Width = 71
        Height = 33
        Caption = #52628#44032
        OnClick = btnAddClick
      end
      object btnDel: TSpeedButton
        Left = 170
        Top = 31
        Width = 71
        Height = 33
        Caption = #49325#51228
        OnClick = btnDelClick
      end
      object btnSave: TSpeedButton
        Left = 98
        Top = 31
        Width = 71
        Height = 33
        Caption = #51200#51109
        OnClick = btnSaveClick
      end
      object Panel5: TPanel
        Left = 25
        Top = 70
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #44160#49324#53076#46300
        Color = 15925222
        TabOrder = 0
      end
      object Panel6: TPanel
        Left = 25
        Top = 97
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #44160#49324#47749
        Color = 15925222
        TabOrder = 1
      end
      object Panel7: TPanel
        Left = 25
        Top = 177
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #44160#49324#50557#50612
        Color = 15925222
        TabOrder = 2
      end
      object Panel8: TPanel
        Left = 25
        Top = 203
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #52280#44256#52824'('#54616')'
        Color = 15925222
        TabOrder = 3
      end
      object Panel9: TPanel
        Left = 25
        Top = 230
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #52280#44256#52824'('#49345')'
        Color = 15925222
        TabOrder = 4
      end
      object edCode: TEdit
        Left = 106
        Top = 70
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 5
      end
      object edName: TEdit
        Left = 106
        Top = 98
        Width = 111
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 50
        TabOrder = 6
      end
      object edAbbr: TEdit
        Left = 106
        Top = 178
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 8
      end
      object Panel10: TPanel
        Left = 25
        Top = 256
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #49692#49436
        Color = 15925222
        TabOrder = 12
      end
      object Panel11: TPanel
        Left = 25
        Top = 124
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = #49688#49888#53076#46300
        Color = 15925222
        TabOrder = 9
      end
      object edUpCode: TEdit
        Left = 106
        Top = 124
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 7
      end
      object edSub: TEdit
        Left = 106
        Top = 150
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 10
      end
      object Panel12: TPanel
        Left = 25
        Top = 150
        Width = 78
        Height = 22
        BevelInner = bvLowered
        Caption = 'SUB'#53076#46300
        Color = 15925222
        TabOrder = 11
      end
      object edLow: TEdit
        Left = 106
        Top = 203
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 13
      end
      object edHigh: TEdit
        Left = 106
        Top = 229
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 14
      end
      object edSeq: TEdit
        Left = 106
        Top = 256
        Width = 95
        Height = 21
        ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
        MaxLength = 10
        TabOrder = 15
      end
    end
  end
end

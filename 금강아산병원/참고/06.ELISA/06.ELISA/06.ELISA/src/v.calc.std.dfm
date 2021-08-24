object vCalcStd: TvCalcStd
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsNone
  Caption = 'vCalcStd'
  ClientHeight = 543
  ClientWidth = 378
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 14
  object LabelResult: THTMLabel
    AlignWithMargins = True
    Left = 3
    Top = 4
    Width = 372
    Height = 68
    Margins.Top = 4
    Align = alTop
    BorderWidth = 1
    BorderStyle = bsSingle
    BorderColor = clSilver
    Color = clBtnHighlight
    HTMLText.Strings = (
      'LabelResult')
    ParentColor = False
    Transparent = False
    Version = '1.9.2.6'
    ExplicitLeft = 8
    ExplicitTop = 3
    ExplicitWidth = 358
  end
  object Grid: TAdvStringGrid
    AlignWithMargins = True
    Left = 3
    Top = 148
    Width = 372
    Height = 113
    Cursor = crDefault
    Margins.Bottom = 0
    Align = alTop
    BorderStyle = bsNone
    Ctl3D = False
    DefaultColWidth = 71
    DrawingStyle = gdsClassic
    RowCount = 5
    ParentCtl3D = False
    ScrollBars = ssNone
    TabOrder = 0
    HoverRowCells = [hcNormal, hcSelected]
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      ''
      'Conc'
      'Mean'
      '%CV'
      'QC Result')
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    DefaultAlignment = taCenter
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
    FilterDropDown.Font.Style = []
    FilterDropDown.TextChecked = 'Checked'
    FilterDropDown.TextUnChecked = 'Unchecked'
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Larger than'
      'Smaller than'
      'Clear')
    FixedColWidth = 62
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = [fsBold]
    Flat = True
    FloatFormat = '%.2f'
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HoverFixedCells = hfAll
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = DEFAULT_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'Tahoma'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'Tahoma'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'Tahoma'
    PrintSettings.FooterFont.Style = []
    PrintSettings.PageNumSep = '/'
    RowHeaders.Strings = (
      ''
      'S1'
      'S2'
      'S3'
      'S4')
    SearchFooter.FindNextCaption = 'Find &next'
    SearchFooter.FindPrevCaption = 'Find &previous'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SearchFooter.HighLightCaption = 'Highlight'
    SearchFooter.HintClose = 'Close'
    SearchFooter.HintFindNext = 'Find next occurrence'
    SearchFooter.HintFindPrev = 'Find previous occurrence'
    SearchFooter.HintHighlight = 'Highlight occurrences'
    SearchFooter.MatchCaseCaption = 'Match case'
    SearchFooter.ResultFormat = '(%d of %d)'
    ShowSelection = False
    ShowDesignHelper = False
    SortSettings.DefaultFormat = ssAutomatic
    Version = '8.2.4.0'
    ColWidths = (
      62
      71
      71
      71
      71)
    RowHeights = (
      22
      22
      22
      22
      22)
  end
  object Chart: TChart
    AlignWithMargins = True
    Left = 3
    Top = 264
    Width = 372
    Height = 276
    Legend.Visible = False
    Title.Alignment = taLeftJustify
    Title.Color = clBlack
    Title.Font.Color = clBlack
    Title.Font.Height = -12
    Title.Margins.Left = 9
    Title.Margins.Top = 8
    Title.Margins.Units = maPixels
    Title.Shadow.HorizSize = 0
    Title.Text.Strings = (
      'Calculated Plots:')
    Title.TextAlignment = taLeftJustify
    Title.VertMargin = 3
    BottomAxis.ExactDateTime = False
    BottomAxis.Grid.Style = psDash
    BottomAxis.MaximumOffset = 2
    BottomAxis.Ticks.Visible = False
    Hover.Visible = False
    LeftAxis.Axis.Color = clDefault
    LeftAxis.ExactDateTime = False
    LeftAxis.Grid.Style = psDash
    LeftAxis.Increment = 1.000000000000000000
    LeftAxis.MaximumOffset = 2
    LeftAxis.Ticks.Visible = False
    Panning.MouseWheel = pmwNone
    View3D = False
    Align = alClient
    BevelOuter = bvNone
    Color = clWhite
    TabOrder = 1
    DefaultCanvas = 'TGDIPlusCanvas'
    PrintMargins = (
      15
      13
      15
      13)
    ColorPaletteIndex = 13
    object SeriesStd: TPointSeries
      Selected.Hover.Visible = False
      Marks.Callout.Length = 8
      ClickableLine = False
      Pointer.Brush.Color = clBlack
      Pointer.HorizSize = 2
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.VertSize = 2
      XValues.Name = 'X'
      XValues.Order = loAscending
      YValues.Name = 'Y'
      YValues.Order = loNone
      Data = {
        01030000003D0AD7A3703DF6BFD7A3703D0AD707C000000000000000007B14AE
        47E17AF8BF3D0AD7A3703DF63F295C8FC2F528BCBF}
      Detail = {0000000000}
    end
    object SeriesFormula: TFastLineSeries
      LinePen.Color = 3513587
      XValues.Name = 'X'
      XValues.Order = loAscending
      YValues.Name = 'Y'
      YValues.Order = loNone
      object LinearFunc: TCustomTeeFunction
        CalcByValue = False
        Period = 1.000000000000000000
        NumPoints = 4
        OnCalculate = LinearFuncCalculate
      end
    end
  end
  object Panel1: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 75
    Width = 372
    Height = 70
    Margins.Top = 0
    Margins.Bottom = 0
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    DesignSize = (
      372
      70)
    object Label1: TLabel
      Left = 5
      Top = 4
      Width = 56
      Height = 14
      Caption = 'Intercept:'
    end
    object Label2: TLabel
      Left = 5
      Top = 27
      Width = 34
      Height = 14
      Caption = 'Slope:'
    end
    object Label3: TLabel
      Left = 5
      Top = 50
      Width = 124
      Height = 14
      Caption = 'Correlation Coefficient:'
    end
    object LabelCorreCoef: TLabel
      Left = 282
      Top = 50
      Width = 23
      Height = 14
      Alignment = taRightJustify
      Anchors = [akTop, akRight]
      Caption = 'Pass'
      StyleElements = [seClient, seBorder]
    end
    object EditCorreCoef: TRzNumericEdit
      Left = 307
      Top = 46
      Width = 62
      Height = 22
      Anchors = [akTop, akRight]
      Color = clHighlightText
      ReadOnly = True
      ReadOnlyColor = clHighlightText
      TabOrder = 0
      DisplayFormat = '0.#0'
    end
    object EditIntercept: TRzNumericEdit
      Left = 307
      Top = 0
      Width = 62
      Height = 22
      Anchors = [akTop, akRight]
      Color = clHighlightText
      ReadOnly = True
      ReadOnlyColor = clHighlightText
      TabOrder = 1
      IntegersOnly = False
      DisplayFormat = '0.###0'
    end
    object EditSlope: TRzNumericEdit
      Left = 307
      Top = 23
      Width = 62
      Height = 22
      Anchors = [akTop, akRight]
      Color = clHighlightText
      ReadOnly = True
      ReadOnlyColor = clHighlightText
      TabOrder = 2
      DisplayFormat = '0.###0'
    end
  end
end

object F_Main: TF_Main
  Left = 219
  Top = 160
  ActiveControl = gdIf
  Caption = '  SANSOFT ABL Interface Program'
  ClientHeight = 752
  ClientWidth = 939
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
    Height = 636
    Align = alClient
    BevelOuter = bvNone
    Caption = 'Panel2'
    TabOrder = 0
    ExplicitHeight = 616
    object gdIf: TAdvStringGrid
      Left = 0
      Top = 0
      Width = 939
      Height = 499
      Cursor = crDefault
      Align = alClient
      Color = clWhite
      Ctl3D = False
      DefaultColWidth = 50
      DefaultRowHeight = 21
      FixedColor = 16763594
      FixedCols = 0
      RowCount = 2
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing]
      ParentCtl3D = False
      ParentFont = False
      ParentShowHint = False
      ScrollBars = ssBoth
      ShowHint = True
      TabOrder = 0
      OnGetCellColor = gdIfGetCellColor
      OnGetAlignment = gdIfGetAlignment
      OnClickCell = gdIfClickCell
      OnDblClickCell = gdIfDblClickCell
      OnCanEditCell = gdIfCanEditCell
      ActiveCellFont.Charset = ANSI_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -12
      ActiveCellFont.Name = #44404#47548#52404
      ActiveCellFont.Style = [fsBold]
      Bands.PrimaryColor = 15663103
      Bands.SecondaryColor = clWhite
      Bands.Print = True
      CellNode.TreeColor = clSilver
      ColumnHeaders.Strings = (
        ''
        #48148#53076#46300
        #54872#51088#48264#54840
        #54872#51088#49457#47749
        #49345#53468)
      ControlLook.ControlStyle = csWinXP
      Filter = <>
      FilterDropDown.Font.Charset = DEFAULT_CHARSET
      FilterDropDown.Font.Color = clWindowText
      FilterDropDown.Font.Height = -11
      FilterDropDown.Font.Name = 'Tahoma'
      FilterDropDown.Font.Style = []
      FixedColWidth = 22
      FixedRowHeight = 25
      FixedFont.Charset = ANSI_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -12
      FixedFont.Name = 'Arial'
      FixedFont.Style = [fsBold]
      Flat = True
      FloatFormat = '%.2f'
      Navigation.AlwaysEdit = True
      Navigation.AdvanceOnEnter = True
      Navigation.AdvanceDirection = adTopBottom
      Navigation.AdvanceAuto = True
      PrintSettings.DateFormat = 'dd/mm/yyyy'
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
      PrintSettings.PageNumSep = '/'
      ScrollWidth = 16
      SearchFooter.FindNextCaption = 'Find &next'
      SearchFooter.FindPrevCaption = 'Find &previous'
      SearchFooter.Font.Charset = DEFAULT_CHARSET
      SearchFooter.Font.Color = clWindowText
      SearchFooter.Font.Height = -11
      SearchFooter.Font.Name = 'Tahoma'
      SearchFooter.Font.Style = []
      SearchFooter.HighLightCaption = 'Highlight'
      SearchFooter.HintClose = 'Close'
      SearchFooter.HintFindNext = 'Find next occurence'
      SearchFooter.HintFindPrev = 'Find previous occurence'
      SearchFooter.HintHighlight = 'Highlight occurences'
      SearchFooter.MatchCaseCaption = 'Match case'
      SortSettings.Column = 0
      VAlignment = vtaCenter
      ExplicitHeight = 479
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
        TabOrder = 1
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
      Top = 499
      Width = 939
      Height = 137
      Align = alBottom
      TabOrder = 1
      ExplicitTop = 479
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
          FixedColor = 16763620
          FixedCols = 0
          RowCount = 6
          Font.Charset = HANGEUL_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = #44404#47548#52404
          Font.Style = [fsBold]
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssBoth
          TabOrder = 0
          OnGetCellColor = gdResultGetCellColor
          ActiveCellFont.Charset = DEFAULT_CHARSET
          ActiveCellFont.Color = clWindowText
          ActiveCellFont.Height = -11
          ActiveCellFont.Name = 'Tahoma'
          ActiveCellFont.Style = [fsBold]
          CellNode.TreeColor = clSilver
          ColumnHeaders.Strings = (
            ''
            #44160#49324#47749
            #44208#44284
            #44160#49324#47749
            #44208#44284
            #44160#49324#47749
            #44208#44284)
          ControlLook.ControlStyle = csWinXP
          Filter = <>
          FilterDropDown.Font.Charset = DEFAULT_CHARSET
          FilterDropDown.Font.Color = clWindowText
          FilterDropDown.Font.Height = -11
          FilterDropDown.Font.Name = 'Tahoma'
          FilterDropDown.Font.Style = []
          FixedColWidth = 6
          FixedFont.Charset = DEFAULT_CHARSET
          FixedFont.Color = clWindowText
          FixedFont.Height = -11
          FixedFont.Name = 'Tahoma'
          FixedFont.Style = [fsBold]
          Flat = True
          FloatFormat = '%.2f'
          PrintSettings.DateFormat = 'dd/mm/yyyy'
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
          PrintSettings.PageNumSep = '/'
          ScrollWidth = 16
          SearchFooter.FindNextCaption = 'Find &next'
          SearchFooter.FindPrevCaption = 'Find &previous'
          SearchFooter.Font.Charset = DEFAULT_CHARSET
          SearchFooter.Font.Color = clWindowText
          SearchFooter.Font.Height = -11
          SearchFooter.Font.Name = 'Tahoma'
          SearchFooter.Font.Style = []
          SearchFooter.HighLightCaption = 'Highlight'
          SearchFooter.HintClose = 'Close'
          SearchFooter.HintFindNext = 'Find next occurence'
          SearchFooter.HintFindPrev = 'Find previous occurence'
          SearchFooter.HintHighlight = 'Highlight occurences'
          SearchFooter.MatchCaseCaption = 'Match case'
          SortSettings.Column = 0
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
    Top = 733
    Width = 939
    Height = 19
    Panels = <
      item
        Width = 500
      end>
    ExplicitTop = 713
  end
  object pnLog: TPanel
    Left = 0
    Top = 636
    Width = 939
    Height = 97
    Align = alBottom
    TabOrder = 2
    Visible = False
    ExplicitTop = 616
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

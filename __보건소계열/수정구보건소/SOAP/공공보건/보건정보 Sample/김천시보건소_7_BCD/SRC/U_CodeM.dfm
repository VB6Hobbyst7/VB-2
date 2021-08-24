object F_CodeM: TF_CodeM
  Left = 217
  Top = 251
  Caption = #44160#49324#53076#46300' '#49444#51221
  ClientHeight = 469
  ClientWidth = 652
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
    Width = 652
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
  object StatusBar1: TStatusBar
    Left = 0
    Top = 450
    Width = 652
    Height = 19
    Panels = <>
  end
  object Panel2: TPanel
    Left = 0
    Top = 43
    Width = 652
    Height = 407
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 2
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 650
      Height = 405
      Align = alClient
      Caption = 'Panel3'
      TabOrder = 0
      object gdCodeM: TAdvStringGrid
        Left = 1
        Top = 1
        Width = 361
        Height = 403
        Cursor = crDefault
        Align = alClient
        ColCount = 8
        DefaultRowHeight = 21
        FixedCols = 0
        RowCount = 2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected]
        ParentFont = False
        ScrollBars = ssBoth
        TabOrder = 0
        OnClick = gdCodeMClick
        OnClickCell = gdCodeMClickCell
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        CellNode.TreeColor = clSilver
        ColumnHeaders.Strings = (
          #49692#48264
          #44160#49324#53076#46300
          #44160#49324#47749
          #44160#49324#50557#50612
          #51204#49569#53076#46300
          'SUB'#53076#46300
          #52280#44256#52824'('#54616')'
          #52280#44256#52824'('#49345')')
        Filter = <>
        FixedColWidth = 37
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = clWindowText
        FixedFont.Height = -11
        FixedFont.Name = 'Tahoma'
        FixedFont.Style = [fsBold]
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
        Left = 362
        Top = 1
        Width = 287
        Height = 403
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
          OnKeyPress = MySelectNext
        end
        object edName: TEdit
          Left = 106
          Top = 98
          Width = 111
          Height = 21
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          MaxLength = 50
          TabOrder = 6
          OnKeyPress = MySelectNext
        end
        object edAbbr: TEdit
          Left = 106
          Top = 178
          Width = 95
          Height = 21
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          MaxLength = 10
          TabOrder = 8
          OnKeyPress = MySelectNext
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
          OnKeyPress = MySelectNext
        end
        object edSub: TEdit
          Left = 106
          Top = 150
          Width = 95
          Height = 21
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          MaxLength = 10
          TabOrder = 10
          OnKeyPress = MySelectNext
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
          OnKeyPress = MySelectNext
        end
        object edHigh: TEdit
          Left = 106
          Top = 229
          Width = 95
          Height = 21
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          MaxLength = 10
          TabOrder = 14
          OnKeyPress = MySelectNext
        end
        object edSeq: TEdit
          Left = 106
          Top = 256
          Width = 95
          Height = 21
          ImeName = #54620#44397#50612' '#51077#47141' '#49884#49828#53596' (IME 2000)'
          MaxLength = 10
          TabOrder = 15
          OnKeyPress = MySelectNext
        end
      end
    end
  end
end

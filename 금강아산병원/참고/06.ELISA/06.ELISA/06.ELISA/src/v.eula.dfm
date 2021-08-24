object vEula: TvEula
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsDialog
  Caption = 'ELISA Report - End User License Agreement'
  ClientHeight = 423
  ClientWidth = 459
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 453
    Height = 16
    Align = alTop
    Alignment = taCenter
    Caption = 'Review and accept the Software License Terms'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    ExplicitWidth = 308
  end
  object MemoEula: TRzRichEdit
    AlignWithMargins = True
    Left = 3
    Top = 25
    Width = 453
    Height = 360
    Align = alClient
    Color = clHighlightText
    Font.Charset = HANGEUL_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Lines.Strings = (
      'BioNote ELISA Report'
      ''
      'End User License Agreement'
      ''
      'NOTICE TO USER: --------')
    ParentFont = False
    ReadOnly = True
    TabOrder = 0
    Zoom = 100
    ReadOnlyColor = clHighlightText
  end
  object GridPanel1: TGridPanel
    Left = 0
    Top = 388
    Width = 459
    Height = 35
    Align = alBottom
    BevelOuter = bvNone
    ColumnCollection = <
      item
        Value = 50.000000000000000000
      end
      item
        Value = 50.000000000000000000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = ButtonAgree
        Row = 0
      end
      item
        Column = 1
        Control = Panel1
        Row = 0
      end>
    RowCollection = <
      item
        Value = 100.000000000000000000
      end>
    TabOrder = 1
    object ButtonAgree: TButton
      AlignWithMargins = True
      Left = 151
      Top = 3
      Width = 75
      Height = 29
      Align = alRight
      Caption = 'I &Agree'
      ModalResult = 1
      TabOrder = 0
      OnClick = ButtonAgreeClick
    end
    object Panel1: TPanel
      Left = 229
      Top = 0
      Width = 230
      Height = 35
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object ButtonCancel: TButton
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 75
        Height = 29
        Align = alLeft
        Caption = '&Cancel'
        ModalResult = 2
        TabOrder = 0
      end
      object ButtonOk: TButton
        AlignWithMargins = True
        Left = 152
        Top = 3
        Width = 75
        Height = 29
        Align = alRight
        Caption = '&OK'
        ModalResult = 1
        TabOrder = 1
      end
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      '.Caption'
      'ButtonAgree.Caption'
      'ButtonCancel.Caption'
      'ButtonOk.Caption'
      'Label1.Caption')
    Left = 40
    Top = 96
  end
end

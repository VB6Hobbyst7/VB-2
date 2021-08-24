object vI18n: TvI18n
  Left = 0
  Top = 0
  ActiveControl = ComboLang
  BiDiMode = bdLeftToRight
  BorderStyle = bsNone
  Caption = 'vI18n'
  ClientHeight = 50
  ClientWidth = 303
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  OnActivate = FormActivate
  OnCreate = FormCreate
  DesignSize = (
    303
    50)
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    AlignWithMargins = True
    Left = 42
    Top = 31
    Width = 72
    Height = 13
    Alignment = taRightJustify
    Caption = 'Date format:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    Layout = tlCenter
  end
  object Label1: TLabel
    AlignWithMargins = True
    Left = 56
    Top = 4
    Width = 58
    Height = 13
    Alignment = taRightJustify
    Caption = 'Language:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    Layout = tlCenter
  end
  object ComboLang: TCultureBox
    AlignWithMargins = True
    Left = 120
    Top = 0
    Width = 175
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    Color = clWindow
    DisplayName = cnNativeLanguageName
    Flags = svcI18n.FlagImg
    Items.Cultures = (
      '*')
    Localizer = svcI18n.Localizer
    TabOrder = 1
    OnKeyDown = CtrlKeyDown
    OnSelect = ComboLangSelect
  end
  object ComboDateFmt: TComboBox
    AlignWithMargins = True
    Left = 120
    Top = 27
    Width = 175
    Height = 21
    Style = csDropDownList
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 0
    OnClick = ComboDateFmtClick
    OnKeyDown = CtrlKeyDown
    Items.Strings = (
      'YYYY-MM-DD'
      'MM-DD-YYYY'
      'DD-MM-YYYY')
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      'Label1.Caption'
      'Label2.Caption')
    Left = 184
    Top = 8
  end
end

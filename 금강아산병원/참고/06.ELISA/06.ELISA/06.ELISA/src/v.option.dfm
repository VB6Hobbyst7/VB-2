object vOption: TvOption
  Left = 0
  Top = 0
  BiDiMode = bdLeftToRight
  BorderStyle = bsDialog
  Caption = 'ELISA Report Option'
  ClientHeight = 115
  ClientWidth = 357
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  DesignSize = (
    357
    115)
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 357
    Height = 73
    Align = alTop
    BevelInner = bvLowered
    Color = clWhite
    ParentBackground = False
    TabOrder = 0
    StyleElements = [seFont]
    DesignSize = (
      357
      73)
    object PanelI18n: TPanel
      Left = 8
      Top = 12
      Width = 337
      Height = 52
      Anchors = [akLeft, akTop, akRight]
      BevelOuter = bvNone
      TabOrder = 0
    end
  end
  object ButtonCancel: TButton
    Left = 270
    Top = 84
    Width = 75
    Height = 25
    Anchors = [akTop, akRight]
    Caption = '&Cancel'
    ModalResult = 2
    TabOrder = 1
    OnClick = ButtonCancelClick
  end
  object ButtonOk: TButton
    Left = 189
    Top = 84
    Width = 75
    Height = 25
    Anchors = [akTop, akRight]
    Caption = '&OK'
    ModalResult = 1
    TabOrder = 2
    OnClick = ButtonOkClick
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      '.Caption'
      'ButtonCancel.Caption'
      'ButtonOk.Caption')
    Left = 72
    Top = 56
  end
end

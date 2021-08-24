object vProgressDlg: TvProgressDlg
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Print...'
  ClientHeight = 67
  ClientWidth = 304
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object LabelProgress: TLabel
    Left = 239
    Top = 26
    Width = 42
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = '1/28'
  end
  object Progress: TProgressBar
    Left = 32
    Top = 24
    Width = 201
    Height = 17
    TabOrder = 0
  end
end

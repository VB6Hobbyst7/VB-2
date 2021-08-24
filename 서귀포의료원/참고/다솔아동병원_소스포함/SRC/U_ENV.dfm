object F_ENV: TF_ENV
  Left = 192
  Top = 107
  Width = 314
  Height = 326
  Caption = 'F_ENV'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 16
    Top = 160
    Width = 281
    Height = 129
    Caption = #44536#47000#54532' '#49353#49345' '#49444#51221
    TabOrder = 0
    object shMEAN: TShape
      Left = 186
      Top = 32
      Width = 33
      Height = 21
      Shape = stCircle
      OnMouseDown = shRESMouseDown
    end
    object shRES: TShape
      Left = 82
      Top = 30
      Width = 33
      Height = 21
      Shape = stCircle
      OnMouseDown = shRESMouseDown
    end
    object sh1SD: TShape
      Left = 82
      Top = 56
      Width = 33
      Height = 21
      Shape = stCircle
      OnMouseDown = shRESMouseDown
    end
    object sh2SD: TShape
      Left = 187
      Top = 58
      Width = 33
      Height = 21
      Shape = stCircle
      OnMouseDown = shRESMouseDown
    end
    object sh3SD: TShape
      Left = 82
      Top = 83
      Width = 33
      Height = 21
      Shape = stCircle
      OnMouseDown = shRESMouseDown
    end
    object Panel1: TPanel
      Left = 27
      Top = 30
      Width = 60
      Height = 20
      BevelInner = bvLowered
      Caption = #44208#44284
      TabOrder = 0
    end
    object Panel2: TPanel
      Left = 131
      Top = 31
      Width = 60
      Height = 20
      BevelInner = bvLowered
      Caption = 'MEAN'
      TabOrder = 1
    end
    object Panel3: TPanel
      Left = 27
      Top = 56
      Width = 60
      Height = 20
      BevelInner = bvLowered
      Caption = '1 SD'
      TabOrder = 2
    end
    object Panel4: TPanel
      Left = 132
      Top = 58
      Width = 60
      Height = 20
      BevelInner = bvLowered
      Caption = '2 SD'
      TabOrder = 3
    end
    object Panel5: TPanel
      Left = 27
      Top = 83
      Width = 60
      Height = 20
      BevelInner = bvLowered
      Caption = '3 SD'
      TabOrder = 4
    end
  end
  object GroupBox2: TGroupBox
    Left = 16
    Top = 48
    Width = 281
    Height = 105
    Caption = #44592#48376#51221#48372
    TabOrder = 1
    object Panel6: TPanel
      Left = 16
      Top = 24
      Width = 73
      Height = 22
      BevelInner = bvLowered
      Caption = #44160#49324#44592#44288#47749
      TabOrder = 0
    end
    object Panel7: TPanel
      Left = 16
      Top = 48
      Width = 73
      Height = 22
      BevelInner = bvLowered
      Caption = 'QC '#54924#49324#47749
      TabOrder = 1
    end
    object edHospNm: TEdit
      Left = 91
      Top = 24
      Width = 158
      Height = 21
      ImeName = 'Microsoft IME 2003'
      TabOrder = 2
    end
    object edCor: TEdit
      Left = 91
      Top = 48
      Width = 158
      Height = 21
      ImeName = 'Microsoft IME 2003'
      TabOrder = 3
    end
  end
  object btnSave: TButton
    Left = 136
    Top = 9
    Width = 75
    Height = 30
    Caption = #51200#51109
    TabOrder = 2
    OnClick = btnSaveClick
  end
  object Button1: TButton
    Left = 216
    Top = 9
    Width = 75
    Height = 30
    Caption = #45803#44592
    TabOrder = 3
    OnClick = Button1Click
  end
  object ColorDialog1: TColorDialog
    Left = 224
    Top = 128
  end
end

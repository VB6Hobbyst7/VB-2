object F_CommSet: TF_CommSet
  Left = 401
  Top = 345
  Width = 266
  Height = 262
  Caption = #52980#54252#53944' '#49444#51221
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 225
    Height = 219
    Caption = #54252#53944#49444#51221
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = #44404#47548#52404
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object cmbPortNum: TComboBox
      Left = 90
      Top = 24
      Width = 65
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ParentFont = False
      TabOrder = 0
      Items.Strings = (
        '1'
        '2'
        '3'
        '4'
        '5'
        '6'
        '7')
    end
    object cmbBaudrate: TComboBox
      Left = 90
      Top = 50
      Width = 65
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ParentFont = False
      TabOrder = 1
      Items.Strings = (
        '110'
        '300'
        '600'
        '1200'
        '2400'
        '4800'
        '9600'
        '14400'
        '19200'
        '38400'
        '56000'
        '57600'
        '115200'
        '128000'
        '256000')
    end
    object cmbDatabits: TComboBox
      Left = 90
      Top = 76
      Width = 65
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ItemIndex = 4
      ParentFont = False
      TabOrder = 2
      Text = '8'
      Items.Strings = (
        '4'
        '5'
        '6'
        '7'
        '8')
    end
    object cmbStopbits: TComboBox
      Left = 90
      Top = 102
      Width = 65
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ItemIndex = 0
      ParentFont = False
      TabOrder = 3
      Text = '1'
      Items.Strings = (
        '1'
        '1.5'
        '2')
    end
    object cmbParity: TComboBox
      Left = 90
      Top = 128
      Width = 65
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ItemIndex = 0
      ParentFont = False
      TabOrder = 4
      Text = 'None'
      Items.Strings = (
        'None'
        'Odd'
        'Even'
        'Mark'
        'Space')
    end
    object cmbHand: TComboBox
      Left = 90
      Top = 154
      Width = 110
      Height = 20
      Style = csDropDownList
      Font.Charset = HANGEUL_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ImeName = 'Microsoft IME 2003'
      ItemHeight = 12
      ParentFont = False
      TabOrder = 5
      Items.Strings = (
        '0 - comNone'
        '1 - comXOnXoff'
        '2 - comRTS'
        '3 - comRTSXOnXOff')
    end
    object Panel2: TPanel
      Left = 12
      Top = 24
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'PortNo'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 6
    end
    object Panel1: TPanel
      Left = 12
      Top = 128
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'Parity'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 7
    end
    object Panel3: TPanel
      Left = 12
      Top = 102
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'StopBit'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 8
    end
    object Panel4: TPanel
      Left = 12
      Top = 76
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'DataBit'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 9
    end
    object Panel5: TPanel
      Left = 12
      Top = 50
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'Boudrate'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 10
    end
    object Panel6: TPanel
      Left = 12
      Top = 154
      Width = 76
      Height = 20
      BevelInner = bvLowered
      Caption = 'HandShake'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 11
    end
    object cbxDTR: TCheckBox
      Left = 168
      Top = 100
      Width = 43
      Height = 17
      Caption = 'DTR'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 12
    end
    object cbxRTS: TCheckBox
      Left = 168
      Top = 77
      Width = 43
      Height = 17
      Caption = 'RTS'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 13
    end
    object Button1: TButton
      Left = 30
      Top = 183
      Width = 75
      Height = 25
      Caption = #51200#51109
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentFont = False
      TabOrder = 14
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 111
      Top = 183
      Width = 75
      Height = 25
      Caption = #52712#49548
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #44404#47548#52404
      Font.Style = []
      ParentFont = False
      TabOrder = 15
      OnClick = Button2Click
    end
  end
end

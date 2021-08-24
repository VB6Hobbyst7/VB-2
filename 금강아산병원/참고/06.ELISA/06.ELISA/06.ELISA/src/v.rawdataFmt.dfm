object vRawdataFmt: TvRawdataFmt
  Left = 0
  Top = 0
  Anchors = [akLeft, akTop, akRight]
  BiDiMode = bdLeftToRight
  BorderStyle = bsNone
  Caption = 'X'
  ClientHeight = 691
  ClientWidth = 301
  Color = 14671839
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ParentBiDiMode = False
  StyleElements = [seFont, seBorder]
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object PanelItr: TPanel
    Left = 0
    Top = 387
    Width = 301
    Height = 130
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object PanelM2: TPanel
      Left = 0
      Top = 88
      Width = 301
      Height = 44
      Align = alTop
      BevelOuter = bvNone
      DoubleBuffered = True
      ParentDoubleBuffered = False
      TabOrder = 0
      Visible = False
      DesignSize = (
        301
        44)
      object Shape3: TShape
        AlignWithMargins = True
        Left = 15
        Top = 0
        Width = 271
        Height = 1
        Margins.Left = 15
        Margins.Top = 0
        Margins.Right = 15
        Align = alTop
        Pen.Color = clBtnShadow
        ExplicitLeft = 0
        ExplicitTop = 40
        ExplicitWidth = 275
      end
      object Label4: TLabel
        Left = 15
        Top = 15
        Width = 37
        Height = 13
        Caption = 'Current'
      end
      object PanelM2Itr: TGridPanel
        Left = 224
        Top = 9
        Width = 62
        Height = 25
        Anchors = [akTop, akRight]
        BevelOuter = bvNone
        ColumnCollection = <
          item
            Value = 50.000000000000010000
          end
          item
            Value = 49.999999999999990000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = SpeedButton4
            Row = 0
          end
          item
            Column = 1
            Control = SpeedButton5
            Row = 0
          end>
        RowCollection = <
          item
            Value = 100.000000000000000000
          end>
        TabOrder = 0
        object SpeedButton4: TSpeedButton
          Tag = 2
          AlignWithMargins = True
          Left = 6
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'N'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitLeft = 7
        end
        object SpeedButton5: TSpeedButton
          Tag = 1
          AlignWithMargins = True
          Left = 37
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'A'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitLeft = 45
          ExplicitWidth = 24
        end
      end
    end
    object PanelM3: TPanel
      Tag = 1
      Left = 0
      Top = 44
      Width = 301
      Height = 44
      Align = alTop
      BevelOuter = bvNone
      DoubleBuffered = True
      ParentDoubleBuffered = False
      TabOrder = 1
      Visible = False
      DesignSize = (
        301
        44)
      object Shape2: TShape
        AlignWithMargins = True
        Left = 15
        Top = 0
        Width = 271
        Height = 1
        Margins.Left = 15
        Margins.Top = 0
        Margins.Right = 15
        Align = alTop
        Pen.Color = clBtnShadow
        ExplicitLeft = 0
        ExplicitTop = 40
        ExplicitWidth = 275
      end
      object Label3: TLabel
        Left = 15
        Top = 15
        Width = 37
        Height = 13
        Caption = 'Current'
      end
      object PanelM3Itr: TGridPanel
        Left = 193
        Top = 9
        Width = 93
        Height = 25
        Anchors = [akTop, akRight]
        BevelOuter = bvNone
        ColumnCollection = <
          item
            Value = 33.333333333333350000
          end
          item
            Value = 33.333333333333330000
          end
          item
            Value = 33.333333333333320000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = SpeedButton1
            Row = 0
          end
          item
            Column = 1
            Control = SpeedButton2
            Row = 0
          end
          item
            Column = 2
            Control = SpeedButton3
            Row = 0
          end>
        RowCollection = <
          item
            Value = 100.000000000000000000
          end>
        TabOrder = 0
        object SpeedButton1: TSpeedButton
          Tag = 3
          AlignWithMargins = True
          Left = 6
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Constraints.MaxHeight = 25
          Constraints.MaxWidth = 25
          Constraints.MinHeight = 25
          Constraints.MinWidth = 25
          Caption = 'N'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitLeft = 7
        end
        object SpeedButton2: TSpeedButton
          Tag = 2
          AlignWithMargins = True
          Left = 37
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Constraints.MaxHeight = 25
          Constraints.MaxWidth = 25
          Constraints.MinHeight = 25
          Constraints.MinWidth = 25
          Caption = 'A'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitLeft = 45
          ExplicitWidth = 24
        end
        object SpeedButton3: TSpeedButton
          Tag = 1
          AlignWithMargins = True
          Left = 67
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Constraints.MaxHeight = 25
          Constraints.MaxWidth = 25
          Constraints.MinHeight = 25
          Constraints.MinWidth = 25
          Caption = 'M'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitTop = -3
          ExplicitWidth = 26
        end
      end
    end
    object PanelStd: TPanel
      Tag = 2
      Left = 0
      Top = 0
      Width = 301
      Height = 44
      Align = alTop
      BevelOuter = bvNone
      DoubleBuffered = True
      ParentDoubleBuffered = False
      TabOrder = 2
      DesignSize = (
        301
        44)
      object Shape1: TShape
        AlignWithMargins = True
        Left = 15
        Top = 0
        Width = 271
        Height = 1
        Margins.Left = 15
        Margins.Top = 0
        Margins.Right = 15
        Align = alTop
        Pen.Color = clBtnShadow
        ExplicitLeft = 0
        ExplicitTop = 40
        ExplicitWidth = 275
      end
      object Label2: TLabel
        Left = 15
        Top = 15
        Width = 37
        Height = 13
        Caption = 'Current'
      end
      object PanelStdItr: TGridPanel
        Left = 162
        Top = 9
        Width = 124
        Height = 25
        Anchors = [akTop, akRight]
        BevelOuter = bvNone
        ColumnCollection = <
          item
            Value = 25.000000000000000000
          end
          item
            Value = 25.000000000000000000
          end
          item
            Value = 25.000000000000000000
          end
          item
            Value = 25.000000000000000000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = ButtonS1
            Row = 0
          end
          item
            Column = 1
            Control = ButtonS2
            Row = 0
          end
          item
            Column = 2
            Control = ButtonS3
            Row = 0
          end
          item
            Column = 3
            Control = ButtonS4
            Row = 0
          end>
        RowCollection = <
          item
            Value = 100.000000000000000000
          end>
        TabOrder = 0
        object ButtonS1: TSpeedButton
          Tag = 4
          AlignWithMargins = True
          Left = 6
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'S1'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitTop = -5
        end
        object ButtonS2: TSpeedButton
          Tag = 3
          AlignWithMargins = True
          Left = 37
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'S2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitTop = -3
        end
        object ButtonS3: TSpeedButton
          Tag = 2
          AlignWithMargins = True
          Left = 68
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'S3'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitLeft = 69
        end
        object ButtonS4: TSpeedButton
          Tag = 1
          AlignWithMargins = True
          Left = 99
          Top = 0
          Width = 25
          Height = 25
          Margins.Left = 6
          Margins.Top = 0
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alClient
          Caption = 'S4'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
          StyleElements = [seClient, seBorder]
          ExplicitTop = -3
        end
      end
    end
  end
  object Panel1: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 295
    Height = 25
    Align = alTop
    Alignment = taLeftJustify
    BevelOuter = bvNone
    Caption = 'Manual Data Format'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBtnShadow
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    DesignSize = (
      295
      25)
    object ButtonClose: TPngSpeedButton
      Left = 272
      Top = 2
      Width = 23
      Height = 22
      Anchors = [akTop, akRight]
      Flat = True
      OnClick = ButtonCloseClick
      PngImage.Data = {
        89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
        61000000017352474200AECE1CE9000000097048597300000EC300000EC301C7
        6FA864000000864944415478DA6364A0103052D380E9403C11886F10D02301C4
        DD401C8B6E8006102F07E2483C8680346F06E242203E82CD0BF80CC1D08C2B0C
        B019825533BE404436E4032ECDF80C4036E40F2ECD840C90801AC00075C90B52
        0C8069CE84F2A7E3328491806658206AE032849108CD0CF80C612452334E43C8
        49CA2043EAA18650373391050049232A11ECE012B00000000049454E44AE4260
        82}
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 31
    Width = 301
    Height = 122
    Align = alTop
    BevelOuter = bvNone
    ParentColor = True
    TabOrder = 2
    DesignSize = (
      301
      122)
    object Label1: TLabel
      Left = 8
      Top = 12
      Width = 59
      Height = 13
      Caption = 'Item Type'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object RadioSample: TRzRadioButton
      Left = 15
      Top = 68
      Width = 97
      Height = 15
      Caption = 'Subject Samples'
      TabOrder = 0
      OnClick = RadioClick
    end
    object RadioStd: TRzRadioButton
      Tag = 8
      Left = 15
      Top = 39
      Width = 68
      Height = 15
      Caption = 'Standards'
      Checked = True
      TabOrder = 1
      TabStop = True
      OnClick = RadioClick
    end
    object ButtonClear: TButton
      AlignWithMargins = True
      Left = 15
      Top = 94
      Width = 271
      Height = 25
      Margins.Left = 15
      Margins.Top = 0
      Margins.Right = 15
      Margins.Bottom = 0
      Anchors = [akLeft, akTop, akRight]
      Caption = 'Clear All Standards'
      TabOrder = 2
      OnClick = ButtonClearClick
    end
  end
  object PanelMaterial: TPanel
    Left = 0
    Top = 153
    Width = 301
    Height = 98
    Align = alTop
    BevelOuter = bvNone
    ParentColor = True
    TabOrder = 3
    Visible = False
    object Label5: TLabel
      Left = 8
      Top = 12
      Width = 79
      Height = 13
      Caption = 'Samples Type'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Image7: TImage
      Left = 15
      Top = 38
      Width = 24
      Height = 24
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000180000
        00180806000000E0773DF8000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000001
        D24944415478DAE5953128846118C79F2B034559D40D37503750D42917C35D91
        147506C560BB2BCA607083C144192483813A8372E5EA06571683A228068A2806
        E536CA0D06E50683E2FFF4FEBFEEED5CDD772775E5A95F7DEFFB3DDFFB7FDEE7
        7D9EF7F3C81F9BE75F08D48128E8053D1CDF824B9004F9DF0804C00EE804F7E0
        0A7C803ECEE5400C9C5623A08B9C71E118A3B6AD9DE2EA370A0E2A11A807378C
        36C868D78B7CE2143F027ED005DEDC0ACC81352EAE91F78393229F01A6C607EE
        C03698772BB007BC205CE25D94A97104D47699B2A05B8147E634CE71C04A9197
        8BD902BAE315D0043EDD08BC834DB0C0B193A2572994E524B8E0F30C488016FA
        9415D0037E16531DB6805653B284FF0698020D6E53B4CC6DB7312247E01C64E9
        B30A1E40A39843D61E99702BD0CA8F0EF99123609B730609461FB6525656402D
        C44553620EBBB8C6B557B49467C1B49832954A04D4C6415ACC959001D7629A4F
        EFA588986A4AF26CA41A8166F0C43CDB96B7E68EC150B5024B6091CF1929F485
        56D897E5D72D3FEFAAB2029AE317A6449B4B2BC8E96C1F7796A79F36E558A502
        5AA683621A2AC29D74F09DFE17F6C1085395669AB26E0534AA28D8B2E642DC85
        9A9FD1E738D61D0E4B8926AC895F666D0B7C03531C6319CA3E5F9E0000000049
        454E44AE426082}
    end
    object Image6: TImage
      Left = 15
      Top = 68
      Width = 24
      Height = 24
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000180000
        00180806000000E0773DF8000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000001
        C04944415478DAE5953128846118C79F2B03455994E186BBBAE11445B918CE24
        45312806DB5D51060383C144190C1645312857D40DAE6E31288A62A08862A328
        CA0D862BDF6050FCFFBD8FEEEB73F27EDF6590A77E75DFF7BDDFFB7FDEFFF33C
        DF85E49723F42F04AA400A748076BDBE04A720039C4A045AC1066806D7E00CBC
        824EBD570069701844809B1CE9C669CDDA1D7115E7BA01B0E347A01A5C68B609
        500B263D6B96D49E3D10032DA0682BC0CD167573661E01779E3551700FC2E00A
        AC83695B816DD008BACA3C4BA9359F028C4DB52C612B70A39E4EE935C5B2AEDF
        718F004FBC00EAC09B8DC00B5801339602E360153480671B0116F8514C77D858
        B40C46418DAD45F37AECA8661491EF8BCC0E63913923C3B602117D69575FAA97
        F26D5A546B983D1BE2C45680910407604B4CB1BD3DCE59612B4F8031316D2A7E
        041843628ACB4F420E9C8B193E7E97FAC5143B2366D2258800AD79509FDDE1B8
        EEED839EA0027360567FE7A43417ECB077D7BA36F9FAADFA51801E3FA925ECFD
        63294D76584FE6E83A0EE5A05F01764D371851BF7992267DC6FF853CE853ABB2
        6AD3ADAD00B34A8135D7BDA49E8211D3EC0B7ACD13F68A29B8AF1A541C7F5FE0
        033DBF5F19AAC49B020000000049454E44AE426082}
    end
    object RadioM2: TRzRadioButton
      Left = 53
      Top = 43
      Width = 120
      Height = 15
      Caption = 'In-Tube(Nil, Antigen)'
      TabOrder = 0
      OnClick = RadioClick
    end
    object RadioM3: TRzRadioButton
      Tag = 1
      Left = 53
      Top = 73
      Width = 165
      Height = 15
      Caption = 'In-Tube(Nil, Antigen, Mitogen)'
      Checked = True
      TabOrder = 1
      TabStop = True
      OnClick = RadioClick
    end
  end
  object PanelOrientation: TPanel
    Left = 0
    Top = 251
    Width = 301
    Height = 136
    Align = alTop
    BevelOuter = bvNone
    ParentColor = True
    TabOrder = 4
    object LabelOrientation: TLabel
      Left = 8
      Top = 12
      Width = 115
      Height = 13
      Caption = 'Samples Orientation'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Image1: TImage
      Left = 15
      Top = 38
      Width = 24
      Height = 24
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000180000
        00180806000000E0773DF8000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000000
        594944415478DA6364A031601C2C168800F17D20E681F2FF00B121105FA19605
        0A500B908123101F18B560F8582000C40568620B80F8C190F1C1A80523C082D1
        643A6A01E5164800F163206641123305E233D4B2806C40730B00B21729192989
        5BBD0000000049454E44AE426082}
    end
    object Image2: TImage
      Left = 15
      Top = 68
      Width = 24
      Height = 24
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000180000
        00180806000000E0773DF8000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000000
        694944415478DA6364A031601CB560D482510B502C1001621E2C6A3E00B1000E
        FDB8E47E00F10B740B3EE3B0A01188EB7158804BEE0F10B3A25BF09F444308C9
        31D2DD82DF40CC424B0B741820118D0E1E00B1020E4370C97D01E233E816D004
        8C5A306AC1A8050C0C004AA41519DF0529C60000000049454E44AE426082}
    end
    object Image5: TImage
      Left = 15
      Top = 98
      Width = 24
      Height = 24
      AutoSize = True
      Picture.Data = {
        0954506E67496D61676589504E470D0A1A0A0000000D49484452000000180000
        00180806000000E0773DF8000000017352474200AECE1CE90000000467414D41
        0000B18F0BFC6105000000097048597300000EC300000EC301C76FA864000000
        B04944415478DA6364A031601CB5801C0B348058028BF80D2036C021770488EF
        106BC17C204EC0229E08C4F140EC80436EC1A8052043DC81D8028B5C2910AF21
        D6025281030EF103E816FCC7A1501188F703B102163947A81C4EC78F5A409205
        0A38143E6180E45E162C722F18B0E76C1078806E012E70188865B088DB027137
        03F6641B09C42788B5E03E0EDF81820E94671C7004DD8191634101100B60119F
        00C401382C5FC04042245304462D200800C5762E1984FE53300000000049454E
        44AE426082}
    end
    object RadioVertical: TRzRadioButton
      Left = 53
      Top = 43
      Width = 54
      Height = 15
      Caption = 'Vertical'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = RadioItrClick
    end
    object RadioHorizontal: TRzRadioButton
      Tag = 1
      Left = 53
      Top = 73
      Width = 67
      Height = 15
      Caption = 'Horizontal'
      TabOrder = 1
      OnClick = RadioItrClick
    end
    object RadioRandom: TRzRadioButton
      Tag = 2
      Left = 53
      Top = 103
      Width = 58
      Height = 15
      Caption = 'Random'
      TabOrder = 2
      OnClick = RadioItrClick
    end
  end
  object Translator: TTranslator
    Localizer = svcI18n.Localizer
    Translatables.Properties = (
      'ButtonClear.Caption'
      'Label1.Caption'
      'Label2.Caption'
      'Label3.Caption'
      'Label4.Caption'
      'Label5.Caption'
      'Label6.Caption'
      'Panel1.Caption'
      'RadioHorizontal.Caption'
      'RadioRandom.Caption'
      'RadioSample.Caption'
      'RadioStd.Caption'
      'RadioVertical.Caption')
    Translatables.Literals = (
      '3BACFBD563FE308A0FB610FF1388A915'
      'DB8D503322C46D746776CCB54DBB91EB'
      '0877F321912912A59ABE8417841DA23D'
      '330E77D3673EEC3466BCF2F94C01BDC1')
    Left = 136
    Top = 264
  end
end

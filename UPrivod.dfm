object FPrivod: TFPrivod
  Left = 253
  Top = 110
  Width = 991
  Height = 492
  Caption = 'FPrivod'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 57
    Height = 342
  end
  object SG2: TStringGrid
    Left = 3
    Top = 57
    Width = 980
    Height = 342
    Align = alClient
    ColCount = 9
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    TabOrder = 0
    Visible = False
    ColWidths = (
      64
      131
      90
      64
      64
      64
      64
      64
      64)
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 983
    Height = 57
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 24
      Top = 56
      Width = 84
      Height = 13
      Caption = #1042#1089#1077#1075#1086' '#1085#1072' '#1089#1082#1083#1072#1076#1077
      Visible = False
    end
    object Label2: TLabel
      Left = 128
      Top = 56
      Width = 32
      Height = 13
      Caption = 'Label2'
      Visible = False
    end
    object Label3: TLabel
      Left = 312
      Top = 56
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object Label4: TLabel
      Left = 504
      Top = 56
      Width = 32
      Height = 13
      Caption = 'Label4'
      Visible = False
    end
    object Label5: TLabel
      Left = 232
      Top = 56
      Width = 66
      Height = 13
      Caption = #1055#1086#1090#1088#1077#1073#1085#1086#1089#1090#1100
      Visible = False
    end
    object Label6: TLabel
      Left = 392
      Top = 56
      Width = 100
      Height = 13
      Caption = #1057#1074#1086#1073#1086#1076#1085#1099#1081' '#1086#1089#1090#1072#1090#1086#1082
      Visible = False
    end
    object Label7: TLabel
      Left = 26
      Top = 8
      Width = 38
      Height = 13
      Caption = #1055#1088#1080#1074#1086#1076
      Visible = False
    end
    object ComboBox1: TComboBox
      Left = 25
      Top = 24
      Width = 121
      Height = 21
      ItemHeight = 13
      TabOrder = 0
      Text = 'BLF-24'
      Visible = False
      Items.Strings = (
        'BE230'
        'BE230-12'
        'BE24-12'
        'BF230-10'
        'BF230.1'
        'BF24-10'
        'BF24.1'
        'BLE230'
        'BLE230U-15'
        'BLE24'
        'BLF230-5'
        'BLF230.1'
        'BLF24-5'
        'BLF24.1'
        'GBB336.1E'
        'GEB136.1E'
        'GEB336.1E'
        #1069#1055#1042'-BF230'
        #1069#1055#1042'-BLE230'
        #1069#1055#1042'-BLF230')
    end
    object Button1: TButton
      Left = 167
      Top = 16
      Width = 75
      Height = 25
      Caption = #1040#1085#1072#1083#1080#1079
      TabOrder = 1
      Visible = False
      OnClick = Button1Click
    end
    object Edit1: TEdit
      Left = 304
      Top = 16
      Width = 121
      Height = 21
      TabOrder = 2
      Text = '0'
      Visible = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 399
    Width = 983
    Height = 66
    Align = alBottom
    TabOrder = 2
    object SpeedButton1: TSpeedButton
      Left = 520
      Top = 36
      Width = 23
      Height = 22
      Glyph.Data = {
        26040000424D2604000000000000360000002800000012000000120000000100
        180000000000F0030000000000000000000000000000000000000000FF0000FF
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF00000000FF0080008080800080008080
        8000800080808000800080808000800080808000800080808000800080808000
        80008080800000FF00000000FF80808000800080808000800080808000800080
        80800080008080800080008080800080008080800080008080800080000000FF
        00000000FF008000808080FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080008080
        800080008080808080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080008000808080008000FFFFFF00800080808000800000
        80008080800000FF00000000FF808080008000FFFFFF80808000800080808000
        8000808080008000FFFFFF008000808080008000FFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000808080008000808080008000FFFFFF
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFF008000808080008000FFFFFF0080008080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFFFFFF
        FFFFFFFF008000FFFFFF008000808080008000808080FFFFFFFFFFFFFFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFF008000FFFFFF00
        8000808080008000808080008000808080FFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000FFFFFF008000808080008000808080
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFF808080008000808080008000808080FFFFFF8080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080FFFFFFFFFFFFFFFFFF808080008000808080FFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080008000808080008000808080008000808080008000
        8080800080008080800080008080800080008080800000FF00000000FF808080
        0080008080800080008080800080008080800080008080800080008080800080
        008080800080008080800080000000FF00000000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
        00FF0000FF0000FF0000}
      OnClick = SpeedButton1Click
    end
    object Label8: TLabel
      Left = 552
      Top = 40
      Width = 85
      Height = 13
      Caption = #1047#1072#1103#1074#1082#1072' '#1085#1072' '#1089#1082#1083#1072#1076
    end
    object Button2: TButton
      Left = 792
      Top = 32
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 0
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 704
      Top = 32
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 1
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 360
      Top = 16
      Width = 75
      Height = 25
      Caption = 'Button4'
      TabOrder = 2
      Visible = False
      OnClick = Button4Click
    end
  end
  object SG1: TStringGrid
    Left = 3
    Top = 57
    Width = 980
    Height = 342
    Align = alClient
    ColCount = 8
    DefaultRowHeight = 18
    FixedCols = 4
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    ScrollBars = ssHorizontal
    TabOrder = 3
    OnSetEditText = SG1SetEditText
    ColWidths = (
      64
      100
      82
      94
      58
      73
      94
      355)
    RowHeights = (
      18
      18
      18
      18
      18)
  end
end

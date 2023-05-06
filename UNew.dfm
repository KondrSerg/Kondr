object FVozKl: TFVozKl
  Left = 75
  Top = 245
  Width = 1114
  Height = 715
  Caption = 'FVozKl'
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
  object Label1: TLabel
    Left = 24
    Top = 8
    Width = 76
    Height = 13
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
  end
  object Label2: TLabel
    Left = 160
    Top = 8
    Width = 7
    Height = 13
    Caption = 'A'
  end
  object Label3: TLabel
    Left = 232
    Top = 8
    Width = 7
    Height = 13
    Caption = 'B'
  end
  object Label4: TLabel
    Left = 304
    Top = 8
    Width = 19
    Height = 13
    Caption = #1051#1080#1090
  end
  object Label5: TLabel
    Left = 376
    Top = 8
    Width = 56
    Height = 13
    Caption = #1048#1089#1087#1086#1083#1077#1085#1080#1077
  end
  object Label6: TLabel
    Left = 448
    Top = 8
    Width = 38
    Height = 13
    Caption = #1055#1088#1080#1074#1086#1076
  end
  object Label7: TLabel
    Left = 608
    Top = 8
    Width = 22
    Height = 13
    Caption = #1069#1055#1042
  end
  object Name: TLabel
    Left = 32
    Top = 80
    Width = 28
    Height = 13
    Caption = 'Name'
  end
  object Al: TLabel
    Left = 96
    Top = 80
    Width = 9
    Height = 13
    Caption = 'Al'
  end
  object Bl: TLabel
    Left = 152
    Top = 80
    Width = 9
    Height = 13
    Caption = 'Bl'
  end
  object lPos2: TLabel
    Left = 264
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos2'
  end
  object lPos1: TLabel
    Left = 208
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos1'
  end
  object lPos3: TLabel
    Left = 336
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos3'
  end
  object lPos4: TLabel
    Left = 416
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos4'
  end
  object lPos5: TLabel
    Left = 504
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos5'
  end
  object lPos6: TLabel
    Left = 608
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos6'
  end
  object lPos7: TLabel
    Left = 728
    Top = 80
    Width = 26
    Height = 13
    Caption = 'lPos7'
  end
  object Edit1: TEdit
    Left = 48
    Top = 112
    Width = 561
    Height = 21
    TabOrder = 9
    Text = 'Edit1'
  end
  object Edit2: TEdit
    Left = 392
    Top = 144
    Width = 121
    Height = 21
    TabOrder = 10
    Text = 'Edit2'
  end
  object ComName: TComboBox
    Left = 24
    Top = 24
    Width = 113
    Height = 21
    ItemHeight = 13
    TabOrder = 0
    Text = #1043#1077#1088#1084#1080#1082
    Items.Strings = (
      #1043#1077#1088#1084#1080#1082
      #1056#1077#1075#1091#1083#1103#1088
      #1059#1042#1050
      #1050#1051
      #1050#1051#1040#1056#1040
      #1058#1102#1083#1100#1087#1072#1085'-1'
      #1058#1102#1083#1100#1087#1072#1085'-2'
      #1058#1102#1083#1100#1087#1072#1085'-3'
      #1050#1042#1059)
  end
  object ComIspol: TComboBox
    Left = 376
    Top = 24
    Width = 65
    Height = 21
    ItemHeight = 13
    TabOrder = 4
    Text = 'H'
    Items.Strings = (
      'H'
      'K'
      'B'
      'KB'
      'C')
  end
  object ComEPV: TComboBox
    Left = 608
    Top = 24
    Width = 81
    Height = 21
    ItemHeight = 13
    TabOrder = 6
    Items.Strings = (
      ''
      #1069#1055#1042)
  end
  object ComA: TComboBox
    Left = 160
    Top = 24
    Width = 65
    Height = 21
    ItemHeight = 13
    TabOrder = 1
    Text = '100'
    Items.Strings = (
      '150'
      '200'
      '250'
      '300'
      '400'
      '500'
      '600'
      '700'
      '800'
      '900'
      '1000'
      '1100'
      '1250')
  end
  object ComB: TComboBox
    Left = 232
    Top = 24
    Width = 65
    Height = 21
    ItemHeight = 13
    TabOrder = 2
    Text = '100'
    Items.Strings = (
      '150'
      '200'
      '250'
      '300'
      '400'
      '500'
      '600'
      '700'
      '800'
      '900'
      '1000'
      '1100'
      '1250')
  end
  object ComL: TComboBox
    Left = 304
    Top = 24
    Width = 57
    Height = 21
    ItemHeight = 13
    TabOrder = 3
    Items.Strings = (
      #1051
      #1055
      #1057)
  end
  object ComboBox1: TComboBox
    Left = 448
    Top = 24
    Width = 145
    Height = 21
    ItemHeight = 13
    TabOrder = 5
    Items.Strings = (
      'LM24A'
      'LM24A-S'
      'LM230A'
      'LM230A-S'
      'LM24A-SR'
      'NM24A'
      'NM24A-S'
      'NM230A'
      'NM24A-SR'
      'SM24A'
      'SM24A-S'
      'SM230A'
      'SM230A-S'
      'SM24A-SR'
      'GM24A'
      'GM230A'
      'GM24A-SR'
      'TF24'
      'TF24-S'
      'TF230'
      'TF230-S'
      'TF24-SR'
      'LF24'
      'LF24-S'
      'LF230'
      'LF230-S'
      'LF24-SR'
      'NF24A'
      'NF24A-S2'
      'NF230A '
      'NF230A-S2'
      'NF24A-SR '
      'NF24A-SR-S2'
      'SF24A'
      'SF24A-S2'
      'SF230A '
      'SF230A-S2'
      'SF24A-SR '
      'SF24A-SR-S2'
      '')
  end
  object SGL: TStringGrid
    Left = 24
    Top = 440
    Width = 1065
    Height = 152
    ColCount = 9
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ScrollBars = ssHorizontal
    TabOrder = 12
    ColWidths = (
      64
      59
      30
      53
      37
      45
      44
      345
      64)
  end
  object Button1: TButton
    Left = 736
    Top = 104
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 7
    OnClick = Button1Click
  end
  object SGR: TStringGrid
    Left = 24
    Top = 176
    Width = 1065
    Height = 256
    ColCount = 9
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ScrollBars = ssHorizontal
    TabOrder = 11
    OnClick = SGRClick
    ColWidths = (
      64
      59
      32
      46
      41
      31
      46
      479
      64)
  end
  object Button2: TButton
    Left = 824
    Top = 104
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 8
    OnClick = Button2Click
  end
  object OpenDialog1: TOpenDialog
    Left = 728
    Top = 24
  end
end

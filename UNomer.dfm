object FNomer: TFNomer
  Left = 763
  Top = 440
  Caption = 'FNomer'
  ClientHeight = 165
  ClientWidth = 279
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label3: TLabel
    Left = 32
    Top = 72
    Width = 32
    Height = 13
    Caption = 'Label3'
    Visible = False
  end
  object Label2: TLabel
    Left = 32
    Top = 56
    Width = 32
    Height = 13
    Caption = 'Label2'
    Visible = False
  end
  object Label5: TLabel
    Left = 192
    Top = 72
    Width = 32
    Height = 13
    Caption = 'Label5'
    Visible = False
  end
  object Label4: TLabel
    Left = 136
    Top = 72
    Width = 32
    Height = 13
    Caption = 'Label4'
    Visible = False
  end
  object Label1: TLabel
    Left = 8
    Top = 32
    Width = 38
    Height = 13
    Caption = #1055#1088#1080#1074#1086#1076
  end
  object Label6: TLabel
    Left = 8
    Top = 72
    Width = 34
    Height = 13
    Caption = #1050#1086#1083'-'#1074#1086
  end
  object Label7: TLabel
    Left = 48
    Top = 8
    Width = 32
    Height = 13
    Caption = 'Label7'
  end
  object Label8: TLabel
    Left = 12
    Top = 104
    Width = 32
    Height = 13
    Caption = 'Label8'
  end
  object Label9: TLabel
    Left = 96
    Top = 104
    Width = 32
    Height = 13
    Caption = 'Label9'
  end
  object Edit1: TEdit
    Left = 72
    Top = 24
    Width = 121
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 96
    Top = 128
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 5
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 184
    Top = 128
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 6
    OnClick = Button2Click
  end
  object DateTimePicker1: TDateTimePicker
    Left = 72
    Top = 24
    Width = 186
    Height = 21
    Date = 41226.384218125000000000
    Time = 41226.384218125000000000
    TabOrder = 2
  end
  object ComboBox1: TComboBox
    Left = 72
    Top = 24
    Width = 185
    Height = 21
    TabOrder = 1
    Text = 'BE230'
    OnKeyPress = ComboBox1KeyPress
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
      #1069#1055#1042'-BLF230'
      #1069#1052#1044'-220'
      #1069#1052#1044'-24'
      #1069#1052'-1.2'
      #1069#1052'-1.3'
      #1069#1052'-1.2 (220'#1042')'
      #1069#1052'-1.3 (24'#1042')')
  end
  object ComboBox2: TComboBox
    Left = 72
    Top = 64
    Width = 57
    Height = 21
    TabOrder = 3
    Text = '1'
    OnKeyPress = ComboBox2KeyPress
    Items.Strings = (
      '1'
      '2')
  end
  object Button3: TButton
    Left = 8
    Top = 128
    Width = 75
    Height = 25
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100
    TabOrder = 4
    Visible = False
    OnClick = Button3Click
  end
end

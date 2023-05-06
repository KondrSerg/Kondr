object FPolsRed: TFPolsRed
  Left = 0
  Top = 0
  Caption = 'FPolsRed'
  ClientHeight = 432
  ClientWidth = 480
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 104
    Width = 42
    Height = 13
    Caption = #1059#1095#1072#1089#1090#1086#1082
  end
  object Label2: TLabel
    Left = 24
    Top = 24
    Width = 33
    Height = 13
    Caption = #1054#1090#1076#1077#1083
  end
  object Label3: TLabel
    Left = 24
    Top = 168
    Width = 99
    Height = 13
    Caption = #1053#1072#1095#1072#1083#1100#1085#1080#1082' '#1091#1095#1072#1089#1090#1082#1072
  end
  object Label4: TLabel
    Left = 24
    Top = 8
    Width = 31
    Height = 13
    Caption = 'Label4'
  end
  object Button1: TButton
    Left = 382
    Top = 384
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object ComboBox1: TComboBox
    Left = 24
    Top = 43
    Width = 193
    Height = 21
    ItemIndex = 1
    TabOrder = 1
    Text = #1050#1058#1054
    Items.Strings = (
      #1055#1044#1054
      #1050#1058#1054
      #1055#1088#1086#1080#1079#1074#1086#1076#1089#1090#1074#1086
      #1054#1058#1050
      #1054#1048#1058
      #1044#1088#1091#1075#1086#1081)
  end
  object Button2: TButton
    Left = 264
    Top = 384
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 2
    OnClick = Button2Click
  end
  object ComboBox2: TComboBox
    Left = 24
    Top = 120
    Width = 193
    Height = 21
    ItemIndex = 1
    TabOrder = 3
    Text = #1050#1083#1072#1087#1072#1085#1072
    Items.Strings = (
      #1057#1074#1086#1073#1086#1076#1085#1099#1081
      #1050#1083#1072#1087#1072#1085#1072
      #1050#1062#1050#1055
      #1057#1058#1040#1052
      #1051#1070#1050
      #1047#1072#1075#1086#1090#1086#1074#1082#1072
      #1059#1087#1072#1082#1086#1074#1082#1072)
  end
  object ComboBox3: TComboBox
    Left = 24
    Top = 192
    Width = 193
    Height = 21
    TabOrder = 4
  end
  object Memo1: TMemo
    Left = 272
    Top = 192
    Width = 185
    Height = 41
    Lines.Strings = (
      #1052#1072#1089#1090#1077#1088#1072
      #1056#1091#1082#1086#1074#1086#1076#1080#1090#1077#1083#1080)
    TabOrder = 5
    Visible = False
  end
  object Memo2: TMemo
    Left = 0
    Top = 240
    Width = 1099
    Height = 89
    Lines.Strings = (
      'Memo2')
    TabOrder = 6
    Visible = False
  end
  object ADOConnection1: TADOConnection
    Provider = 'SQLOLEDB.1'
    Left = 17
    Top = 313
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 17
    Top = 257
  end
end

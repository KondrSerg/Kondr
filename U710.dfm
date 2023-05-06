object F710: TF710
  Left = 344
  Top = 371
  Caption = 'F710'
  ClientHeight = 292
  ClientWidth = 492
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
  object Label2: TLabel
    Left = 240
    Top = 24
    Width = 32
    Height = 13
    Caption = 'Label2'
  end
  object Label1: TLabel
    Left = 32
    Top = 48
    Width = 34
    Height = 13
    Caption = #1050#1086#1083'-'#1074#1086
  end
  object btn1: TButton
    Left = 384
    Top = 212
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 2
    OnClick = btn1Click
  end
  object btn2: TButton
    Left = 296
    Top = 212
    Width = 75
    Height = 25
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100
    TabOrder = 1
    OnClick = btn2Click
  end
  object dtp1: TDateTimePicker
    Left = 28
    Top = 20
    Width = 186
    Height = 21
    Date = 42619.685648761570000000
    Time = 42619.685648761570000000
    TabOrder = 0
  end
  object Button1: TButton
    Left = 296
    Top = 243
    Width = 75
    Height = 25
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100
    TabOrder = 3
    OnClick = Button1Click
  end
  object rg1: TRadioGroup
    Left = 296
    Top = 20
    Width = 130
    Height = 131
    Caption = #1060#1072#1084#1080#1083#1080#1103
    Items.Strings = (
      #1052#1086#1088#1086#1079#1086#1074#1072' '#1051'. '#1051'.'
      #1042#1072#1089#1080#1083#1100#1077#1074#1072' '#1045'. '#1042'.'
      #1055#1072#1085#1102#1096#1080#1085#1072' '#1040'. '#1053'.'
      #1061#1086#1084#1103#1082#1086#1074#1072' '#1040'.'#1042'.'
      #1052#1103#1082#1080#1096#1077#1074#1072' '#1045'.'#1054'.')
    TabOrder = 4
    Visible = False
    OnClick = rg1Click
  end
  object Memo1: TMemo
    Left = 28
    Top = 88
    Width = 240
    Height = 172
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 5
  end
  object Edit1: TEdit
    Left = 28
    Top = 61
    Width = 121
    Height = 21
    TabOrder = 6
    Text = '0'
  end
end

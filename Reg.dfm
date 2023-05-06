object Regist: TRegist
  Left = 468
  Top = 377
  Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103
  ClientHeight = 228
  ClientWidth = 320
  Color = clBtnFace
  DefaultMonitor = dmMainForm
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 40
    Top = 136
    Width = 83
    Height = 13
    Caption = #1055#1072#1088#1086#1083#1100' '#1085#1072' '#1087#1086#1095#1090#1091
  end
  object Label20: TLabel
    Left = 232
    Top = 88
    Width = 38
    Height = 13
    Caption = 'Label20'
    Visible = False
  end
  object Button1: TButton
    Left = 222
    Top = 166
    Width = 75
    Height = 25
    Caption = #1054#1082
    TabOrder = 4
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 136
    Top = 166
    Width = 75
    Height = 25
    Caption = #1042#1099#1093#1086#1076
    TabOrder = 3
    OnClick = Button2Click
  end
  object ComboBox1: TComboBox
    Left = 40
    Top = 32
    Width = 145
    Height = 21
    TabOrder = 1
    Visible = False
  end
  object Edit1: TEdit
    Left = 40
    Top = 88
    Width = 145
    Height = 21
    PasswordChar = '*'
    TabOrder = 2
    Visible = False
    OnKeyPress = Edit1KeyPress
  end
  object Memo1: TMemo
    Left = 26
    Top = 197
    Width = 185
    Height = 89
    ScrollBars = ssBoth
    TabOrder = 5
    Visible = False
  end
  object Memo2: TMemo
    Left = 152
    Top = 197
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo2')
    ScrollBars = ssBoth
    TabOrder = 0
    Visible = False
  end
  object Button3: TButton
    Left = 0
    Top = 166
    Width = 121
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1087#1072#1088#1086#1083#1100
    TabOrder = 6
    OnClick = Button3Click
  end
  object Edit2: TEdit
    Left = 152
    Top = 133
    Width = 145
    Height = 21
    TabOrder = 7
  end
end

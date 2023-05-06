object FOTK: TFOTK
  Left = 57
  Top = 209
  Caption = #1054#1058#1050
  ClientHeight = 220
  ClientWidth = 567
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 16
    Width = 136
    Height = 13
    Caption = #1050#1086#1083'-'#1074#1086' '#1087#1088#1080#1085#1103#1090#1099#1093' '#1082#1083#1072#1087#1072#1085#1086#1074
  end
  object Label2: TLabel
    Left = 184
    Top = 16
    Width = 32
    Height = 13
    Caption = 'Label2'
    Visible = False
  end
  object Label3: TLabel
    Left = 184
    Top = 40
    Width = 32
    Height = 13
    Caption = 'Label3'
    Visible = False
  end
  object Label4: TLabel
    Left = 184
    Top = 64
    Width = 32
    Height = 13
    Caption = 'Label4'
    Visible = False
  end
  object Label5: TLabel
    Left = 184
    Top = 88
    Width = 32
    Height = 13
    Caption = 'Label5'
    Visible = False
  end
  object Label6: TLabel
    Left = 184
    Top = 104
    Width = 32
    Height = 13
    Caption = 'Label6'
    Visible = False
  end
  object Label7: TLabel
    Left = 24
    Top = 80
    Width = 32
    Height = 13
    Caption = 'Label7'
    Visible = False
  end
  object Label8: TLabel
    Left = 24
    Top = 104
    Width = 32
    Height = 13
    Caption = 'Label8'
    Visible = False
  end
  object Label9: TLabel
    Left = 24
    Top = 128
    Width = 105
    Height = 13
    Caption = #1044#1072#1090#1072' '#1080#1079#1075#1086#1090#1086#1074#1083#1077#1085#1080#1103': '
  end
  object Label10: TLabel
    Left = 144
    Top = 128
    Width = 38
    Height = 13
    Caption = 'Label10'
  end
  object LabGP: TLabel
    Left = 288
    Top = 12
    Width = 33
    Height = 13
    Caption = 'LabGP'
    Visible = False
  end
  object LabKO: TLabel
    Left = 396
    Top = 12
    Width = 33
    Height = 13
    Caption = 'LabKO'
    Visible = False
  end
  object Label11: TLabel
    Left = 472
    Top = 12
    Width = 6
    Height = 13
    Caption = '0'
    Visible = False
  end
  object Label12: TLabel
    Left = 260
    Top = 84
    Width = 49
    Height = 13
    Caption = #1053#1072#1095#1047#1072#1074#8470
  end
  object Label13: TLabel
    Left = 320
    Top = 84
    Width = 6
    Height = 13
    Caption = '0'
  end
  object Label14: TLabel
    Left = 260
    Top = 104
    Width = 49
    Height = 13
    Caption = #1050#1086#1085#1047#1072#1074#8470
  end
  object Label15: TLabel
    Left = 320
    Top = 104
    Width = 6
    Height = 13
    Caption = '0'
  end
  object Button1: TButton
    Left = 480
    Top = 168
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 4
    OnClick = Button1Click
    OnKeyPress = Button1KeyPress
  end
  object Button2: TButton
    Left = 392
    Top = 168
    Width = 75
    Height = 25
    Caption = #1054#1082
    TabOrder = 3
    OnClick = Button2Click
    OnKeyPress = Button2KeyPress
  end
  object ME1: TEdit
    Left = 24
    Top = 48
    Width = 121
    Height = 21
    TabOrder = 0
    OnKeyPress = ME1KeyPress
  end
  object DateTimePicker1: TDateTimePicker
    Left = 256
    Top = 48
    Width = 105
    Height = 21
    Date = 41533.656474780090000000
    Time = 41533.656474780090000000
    TabOrder = 1
    OnKeyPress = DateTimePicker1KeyPress
  end
  object Button3: TButton
    Left = 304
    Top = 168
    Width = 75
    Height = 25
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100
    TabOrder = 2
    Visible = False
    OnClick = Button3Click
  end
  object rg1: TRadioGroup
    Left = 396
    Top = 16
    Width = 130
    Height = 146
    Caption = #1060#1072#1084#1080#1083#1080#1103
    Items.Strings = (
      #1052#1086#1088#1086#1079#1086#1074#1072' '#1051'. '#1051'.'
      #1042#1072#1089#1080#1083#1100#1077#1074#1072' '#1045'. '#1042'.'
      #1055#1072#1085#1102#1096#1080#1085#1072' '#1040'. '#1053'.'
      #1061#1086#1084#1103#1082#1086#1074#1072' '#1040'.'#1042'.'
      #1052#1103#1082#1080#1096#1077#1074#1072' '#1045'.'#1054'.'
      #1044#1077#1084#1077#1085#1077#1074#1072' '#1040'.'#1070'.')
    TabOrder = 5
    OnClick = rg1Click
  end
end

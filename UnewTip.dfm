object FNewTip: TFNewTip
  Left = 0
  Top = 0
  Caption = 'FNewTip'
  ClientHeight = 407
  ClientWidth = 806
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
  object Label2: TLabel
    Left = 216
    Top = 40
    Width = 6
    Height = 13
    Caption = '0'
  end
  object Label1: TLabel
    Left = 24
    Top = 22
    Width = 63
    Height = 13
    Caption = #1058#1080#1087' '#1082#1083#1072#1087#1072#1085#1072
  end
  object Label3: TLabel
    Left = 240
    Top = 21
    Width = 30
    Height = 13
    Caption = #1064#1080#1092#1088
  end
  object Label4: TLabel
    Left = 360
    Top = 21
    Width = 77
    Height = 13
    Caption = #1053#1086#1074#1099#1081' '#1087#1088#1080#1079#1085#1072#1082
  end
  object Label5: TLabel
    Left = 536
    Top = 40
    Width = 6
    Height = 13
    Caption = '0'
  end
  object rg2: TRadioGroup
    Left = 335
    Top = 3
    Width = 435
    Height = 351
    Caption = #1042#1085#1077#1089#1077#1085#1080#1077' '#1087#1088#1080#1079#1085#1072#1082#1086#1074
    TabOrder = 6
  end
  object rg1: TRadioGroup
    Left = 8
    Top = 3
    Width = 321
    Height = 353
    Caption = #1042#1085#1077#1089#1077#1085#1080#1077' '#1085#1086#1074#1086#1075#1086' '#1090#1080#1087#1072' '#1082#1083#1072#1087#1072#1085#1072
    TabOrder = 4
  end
  object Edit1: TEdit
    Left = 24
    Top = 35
    Width = 161
    Height = 21
    TabOrder = 0
  end
  object btn1: TButton
    Left = 216
    Top = 248
    Width = 75
    Height = 25
    Caption = #1042#1085#1077#1089#1090#1080
    TabOrder = 1
    OnClick = btn1Click
  end
  object btn2: TButton
    Left = 704
    Top = 360
    Width = 75
    Height = 25
    Caption = #1042#1099#1093#1086#1076
    TabOrder = 2
    OnClick = btn2Click
  end
  object Memo1: TMemo
    Left = 24
    Top = 67
    Width = 161
    Height = 206
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object Edit2: TEdit
    Left = 240
    Top = 35
    Width = 81
    Height = 21
    TabOrder = 5
  end
  object Edit3: TEdit
    Left = 360
    Top = 35
    Width = 161
    Height = 21
    TabOrder = 7
  end
  object Edit4: TEdit
    Left = 576
    Top = 35
    Width = 81
    Height = 21
    TabOrder = 8
  end
  object Memo2: TMemo
    Left = 360
    Top = 67
    Width = 161
    Height = 206
    PopupMenu = pm1
    ScrollBars = ssBoth
    TabOrder = 9
  end
  object btn3: TButton
    Left = 552
    Top = 248
    Width = 75
    Height = 25
    Caption = #1042#1085#1077#1089#1090#1080
    TabOrder = 10
    OnClick = btn3Click
  end
  object btn4: TButton
    Left = 424
    Top = 376
    Width = 75
    Height = 25
    Caption = 'btn4'
    TabOrder = 11
    Visible = False
    OnClick = btn4Click
  end
  object pm1: TPopupMenu
    Left = 616
    Top = 120
    object N1: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N1Click
    end
  end
end

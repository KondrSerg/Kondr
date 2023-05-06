object Form5: TForm5
  Left = 698
  Top = 322
  Width = 368
  Height = 464
  Caption = 'Form5'
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
  object Label3: TLabel
    Left = 8
    Top = 132
    Width = 70
    Height = 13
    Caption = #1044#1072#1090#1072' '#1079#1072#1087#1091#1089#1082#1072
    Visible = False
  end
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 360
    Height = 153
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 704
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object Label2: TLabel
      Left = 136
      Top = 32
      Width = 32
      Height = 13
      Caption = 'Label2'
    end
    object Label5: TLabel
      Left = 24
      Top = 12
      Width = 37
      Height = 13
      Caption = #1047#1072#1082#1072#1079': '
    end
    object Label6: TLabel
      Left = 24
      Top = 32
      Width = 47
      Height = 13
      Caption = #1048#1079#1076#1077#1083#1080#1077':'
    end
    object Label7: TLabel
      Left = 24
      Top = 52
      Width = 74
      Height = 13
      Caption = #1054#1073#1097#1077#1077' '#1082#1086#1083'-'#1074#1086':'
    end
    object Label8: TLabel
      Left = 24
      Top = 72
      Width = 84
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086' '#1074' '#1088#1072#1073#1086#1090#1077':'
    end
    object Label9: TLabel
      Left = 24
      Top = 92
      Width = 81
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086' '#1075#1086#1090#1086#1074#1099#1093':'
    end
    object Label10: TLabel
      Left = 136
      Top = 12
      Width = 38
      Height = 13
      Caption = 'Label10'
    end
    object Label11: TLabel
      Left = 136
      Top = 52
      Width = 38
      Height = 13
      Caption = 'Label11'
    end
    object Label12: TLabel
      Left = 136
      Top = 72
      Width = 38
      Height = 13
      Caption = 'Label12'
    end
    object Label13: TLabel
      Left = 136
      Top = 92
      Width = 38
      Height = 13
      Caption = 'Label13'
    end
    object LabKO1: TLabel
      Left = 24
      Top = 112
      Width = 92
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086' '#1074' '#1073#1088#1080#1082#1077#1090#1072#1093
    end
    object LabKO2: TLabel
      Left = 136
      Top = 112
      Width = 39
      Height = 13
      Caption = 'LabKO2'
    end
  end
  object mmo1: TMemo
    Left = 8
    Top = 240
    Width = 769
    Height = 89
    Lines.Strings = (
      'mmo1')
    TabOrder = 4
    Visible = False
  end
  object mmo2: TMemo
    Left = 4
    Top = 268
    Width = 769
    Height = 89
    Lines.Strings = (
      'mmo2')
    TabOrder = 5
    Visible = False
  end
  object btn1: TButton
    Left = 264
    Top = 384
    Width = 75
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100
    TabOrder = 6
    Visible = False
    OnClick = btn1Click
  end
  object SpecGrid1: TStringGrid
    Left = 12
    Top = 156
    Width = 325
    Height = 221
    ColCount = 2
    DefaultRowHeight = 18
    FixedCols = 0
    TabOrder = 3
    ColWidths = (
      107
      157)
  end
  object mmo3: TMemo
    Left = 340
    Top = 144
    Width = 181
    Height = 89
    Lines.Strings = (
      'mmo3')
    ScrollBars = ssBoth
    TabOrder = 2
    Visible = False
  end
  object mmo4: TMemo
    Left = 532
    Top = 140
    Width = 181
    Height = 89
    Lines.Strings = (
      'mmo3')
    ScrollBars = ssBoth
    TabOrder = 1
    Visible = False
  end
end

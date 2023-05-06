object FKolZap: TFKolZap
  Left = 378
  Top = 423
  Width = 745
  Height = 506
  Caption = 'FKolZap'
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
  object Panel1: TPanel
    Left = 0
    Top = 408
    Width = 737
    Height = 64
    Align = alBottom
    TabOrder = 0
    object Button1: TButton
      Left = 544
      Top = 15
      Width = 89
      Height = 25
      Caption = #1055#1088#1080#1085#1103#1090#1100' '#1054#1058#1050
      TabOrder = 0
      Visible = False
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 648
      Top = 15
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 1
      OnClick = Button2Click
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 737
    Height = 65
    Align = alTop
    TabOrder = 1
    object Label2: TLabel
      Left = 56
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label2'
      Visible = False
    end
    object Label1: TLabel
      Left = 16
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object Label3: TLabel
      Left = 104
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object Label4: TLabel
      Left = 152
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label4'
      Visible = False
    end
    object Label5: TLabel
      Left = 192
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label5'
      Visible = False
    end
    object Label6: TLabel
      Left = 240
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label6'
      Visible = False
    end
    object Label7: TLabel
      Left = 496
      Top = 8
      Width = 83
      Height = 13
      Caption = #1044#1072#1090#1072' '#1079#1076#1072#1095#1080' '#1054#1058#1050
      Visible = False
    end
    object Edit1: TEdit
      Left = 16
      Top = 24
      Width = 121
      Height = 21
      TabOrder = 0
    end
    object Button3: TButton
      Left = 160
      Top = 24
      Width = 75
      Height = 25
      Caption = #1055#1086#1082#1072#1079#1072#1090#1100
      TabOrder = 1
      OnClick = Button3Click
    end
    object DateTimePicker1: TDateTimePicker
      Left = 496
      Top = 24
      Width = 186
      Height = 21
      Date = 41480.342078344910000000
      Time = 41480.342078344910000000
      TabOrder = 2
      Visible = False
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 65
    Width = 737
    Height = 343
    Align = alClient
    TabOrder = 2
    object SGS: TStringGrid
      Left = 1
      Top = 1
      Width = 735
      Height = 341
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
      OnSelectCell = SGSSelectCell
      OnSetEditText = SGSSetEditText
      ColWidths = (
        64
        81
        64
        78
        64)
    end
  end
end

object FZapuskVozd: TFZapuskVozd
  Left = 238
  Top = 252
  Width = 509
  Height = 455
  Caption = 'FZapuskVozd'
  Color = clBtnFace
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
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 501
    Height = 49
    Align = alTop
    TabOrder = 0
    object Label3: TLabel
      Left = 0
      Top = 0
      Width = 34
      Height = 13
      Caption = #1053#1086#1084#1077#1088
      Visible = False
    end
    object Label4: TLabel
      Left = 40
      Top = 0
      Width = 6
      Height = 13
      Caption = '0'
      Visible = False
    end
    object Label1: TLabel
      Left = 32
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object Label5: TLabel
      Left = 184
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label5'
      Visible = False
    end
    object Label6: TLabel
      Left = 344
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label6'
      Visible = False
    end
    object Label2: TLabel
      Left = 88
      Top = 8
      Width = 70
      Height = 13
      Caption = #1044#1072#1090#1072' '#1079#1072#1087#1091#1089#1082#1072
    end
    object DateTimePicker1: TDateTimePicker
      Left = 88
      Top = 24
      Width = 186
      Height = 21
      Date = 41443.312102870370000000
      Time = 41443.312102870370000000
      TabOrder = 0
      OnEnter = DateTimePicker1Enter
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 49
    Width = 501
    Height = 152
    Align = alTop
    TabOrder = 1
    object Label7: TLabel
      Left = 104
      Top = 176
      Width = 31
      Height = 13
      Caption = #1047#1072#1082#1072#1079
      Visible = False
    end
    object Label8: TLabel
      Left = 40
      Top = 184
      Width = 52
      Height = 13
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      Visible = False
    end
    object Label11: TLabel
      Left = 24
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label11'
      Visible = False
    end
    object Label9: TLabel
      Left = 216
      Top = 136
      Width = 32
      Height = 13
      Caption = 'Label9'
      Visible = False
    end
    object Label15: TLabel
      Left = 248
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label15'
      Visible = False
    end
    object SG2: TStringGrid
      Left = 8
      Top = 9
      Width = 485
      Height = 120
      ColCount = 8
      DefaultRowHeight = 18
      FixedCols = 0
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = SG2DblClick
      OnSelectCell = SG2SelectCell
      ColWidths = (
        102
        79
        139
        64
        64
        64
        64
        64)
    end
    object Button3: TButton
      Left = 96
      Top = 128
      Width = 99
      Height = 25
      Caption = #1042' '#1101#1090#1086#1084' '#1087#1072#1082#1077#1090#1077'?'
      TabOrder = 1
      Visible = False
      OnClick = Button3Click
    end
    object Edit1: TEdit
      Left = 136
      Top = 152
      Width = 65
      Height = 21
      TabOrder = 2
      Visible = False
    end
    object Edit2: TEdit
      Left = 224
      Top = 152
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 3
      Visible = False
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 201
    Width = 501
    Height = 220
    Align = alClient
    TabOrder = 2
    object SG1: TStringGrid
      Left = 24
      Top = 8
      Width = 469
      Height = 153
      DefaultRowHeight = 18
      FixedCols = 0
      ScrollBars = ssVertical
      TabOrder = 0
      ColWidths = (
        104
        77
        58
        3
        3)
    end
    object Button1: TButton
      Left = 112
      Top = 184
      Width = 75
      Height = 25
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      TabOrder = 2
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 198
      Top = 184
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 3
      OnClick = Button2Click
    end
    object Memo1: TMemo
      Left = 40
      Top = 40
      Width = 201
      Height = 97
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 1
      Visible = False
    end
  end
end

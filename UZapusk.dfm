object FZapusk: TFZapusk
  Left = 410
  Top = 158
  Width = 327
  Height = 508
  Caption = #1047#1072#1087#1091#1089#1082
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 201
    Width = 319
    Height = 273
    Align = alClient
    TabOrder = 0
    object Label3: TLabel
      Left = 8
      Top = 192
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object Label4: TLabel
      Left = 8
      Top = 216
      Width = 32
      Height = 13
      Caption = 'Label4'
      Visible = False
    end
    object SG1: TStringGrid
      Left = 8
      Top = 8
      Width = 297
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
      Left = 144
      Top = 232
      Width = 75
      Height = 25
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      TabOrder = 1
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 230
      Top = 232
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
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
      TabOrder = 3
      Visible = False
    end
    object ProgressBar1: TProgressBar
      Left = 8
      Top = 168
      Width = 297
      Height = 17
      TabOrder = 4
    end
    object SD1: TStringGrid
      Left = 328
      Top = 8
      Width = 689
      Height = 553
      TabOrder = 5
      Visible = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 319
    Height = 49
    Align = alTop
    TabOrder = 1
    object Label14: TLabel
      Left = 0
      Top = 0
      Width = 34
      Height = 13
      Caption = #1053#1086#1084#1077#1088
      Visible = False
    end
    object Label15: TLabel
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
    object Label9: TLabel
      Left = 184
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label9'
      Visible = False
    end
    object Label11: TLabel
      Left = 344
      Top = 16
      Width = 38
      Height = 13
      Caption = 'Label11'
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
    Width = 319
    Height = 152
    Align = alTop
    TabOrder = 2
    object Label13: TLabel
      Left = 104
      Top = 176
      Width = 31
      Height = 13
      Caption = #1047#1072#1082#1072#1079
      Visible = False
    end
    object Label12: TLabel
      Left = 40
      Top = 184
      Width = 52
      Height = 13
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      Visible = False
    end
    object SG2: TStringGrid
      Left = 8
      Top = 9
      Width = 297
      Height = 120
      DefaultRowHeight = 18
      FixedCols = 0
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = SG2DblClick
      ColWidths = (
        102
        79
        139
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
    object ComboBox1: TComboBox
      Left = 464
      Top = 48
      Width = 145
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      Text = 'ComboBox1'
      Visible = False
      Items.Strings = (
        #1048#1089#1093#1086#1076#1085#1099#1077' '#1076#1072#1085#1085#1099#1077
        #1088' 25'
        #1047#1072#1076#1072#1085#1080#1077' '#1085#1072' '#1058#1056#1059#1052#1055#1060
        #1047#1072#1076#1072#1085#1080#1077' '#1085#1072' '#1075#1080#1073#1082#1091
        #1047#1072#1076#1072#1085#1080#1077' '#1085#1072' '#1085#1086#1078#1085#1080#1094#1099
        #1050#1086#1084#1087#1083#1077#1082#1090#1086#1074#1086#1095#1085#1099#1081' '#1083#1080#1089#1090)
    end
    object Edit3: TEdit
      Left = 464
      Top = 80
      Width = 121
      Height = 21
      TabOrder = 5
      Text = 'Edit3'
      Visible = False
    end
    object Edit4: TEdit
      Left = 464
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 6
      Text = 'Edit4'
      Visible = False
    end
    object Edit5: TEdit
      Left = 320
      Top = 72
      Width = 121
      Height = 21
      TabOrder = 7
      Text = '0'
      Visible = False
    end
    object Edit6: TEdit
      Left = 320
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 8
      Text = '0'
      Visible = False
    end
  end
end

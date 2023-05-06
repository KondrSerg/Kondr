object FNAklVoz: TFNAklVoz
  Left = 451
  Top = 224
  Caption = 'FNAklVoz'
  ClientHeight = 447
  ClientWidth = 563
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
  object Panel1: TPanel
    Left = 0
    Top = 225
    Width = 563
    Height = 222
    Align = alClient
    TabOrder = 2
    ExplicitWidth = 571
    ExplicitHeight = 227
    object Label2: TLabel
      Left = 8
      Top = 192
      Width = 32
      Height = 13
      Caption = 'Label2'
      Visible = False
    end
    object Label3: TLabel
      Left = 8
      Top = 216
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object SG1: TStringGrid
      Left = 8
      Top = 8
      Width = 541
      Height = 153
      DefaultRowHeight = 18
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
      ColWidths = (
        104
        77
        58
        14
        18)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
    object Button1: TButton
      Left = 296
      Top = 192
      Width = 75
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 13
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 382
      Top = 192
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 14
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
      TabOrder = 7
      Visible = False
    end
    object ProgressBar1: TProgressBar
      Left = 8
      Top = 168
      Width = 297
      Height = 17
      TabOrder = 8
    end
    object SD1: TStringGrid
      Left = 572
      Top = 28
      Width = 265
      Height = 161
      TabOrder = 5
      Visible = False
      ColWidths = (
        64
        64
        64
        64
        64)
      RowHeights = (
        24
        24
        24
        24
        24)
    end
    object Button4: TButton
      Left = 232
      Top = 280
      Width = 75
      Height = 25
      Caption = 'Button4'
      TabOrder = 17
      Visible = False
    end
    object Button5: TButton
      Left = 144
      Top = 280
      Width = 75
      Height = 25
      Caption = 'Button5'
      TabOrder = 16
      Visible = False
    end
    object SD: TStringGrid
      Left = 461
      Top = 176
      Width = 384
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 9
      Visible = False
      ColWidths = (
        64
        64
        186
        161
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD2: TStringGrid
      Left = 624
      Top = 188
      Width = 506
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 10
      Visible = False
      ColWidths = (
        64
        64
        191
        169
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD9: TStringGrid
      Left = 505
      Top = 300
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 19
      Visible = False
      ColWidths = (
        64
        64
        134
        146
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD5: TStringGrid
      Left = 501
      Top = 408
      Width = 609
      Height = 105
      ColCount = 21
      DefaultRowHeight = 20
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 21
      ColWidths = (
        64
        64
        129
        216
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        20
        20)
    end
    object SD6: TStringGrid
      Left = 5
      Top = 300
      Width = 492
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 18
      ColWidths = (
        64
        64
        132
        141
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD7: TStringGrid
      Left = 501
      Top = 520
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 23
      ColWidths = (
        64
        64
        134
        146
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object Stenki: TStringGrid
      Left = 493
      Top = 632
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 25
      ColWidths = (
        64
        64
        134
        146
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD8: TStringGrid
      Left = 4
      Top = 632
      Width = 481
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 24
      ColWidths = (
        64
        64
        134
        146
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD4: TStringGrid
      Left = 840
      Top = 28
      Width = 305
      Height = 149
      ColCount = 17
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 6
      ColWidths = (
        64
        64
        203
        158
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SD3: TStringGrid
      Left = 4
      Top = 520
      Width = 489
      Height = 105
      ColCount = 16
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 22
      Visible = False
      ColWidths = (
        64
        64
        196
        155
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object SSvarka: TStringGrid
      Left = 4
      Top = 408
      Width = 489
      Height = 105
      ColCount = 16
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 20
      Visible = False
      ColWidths = (
        64
        64
        196
        155
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object btn1: TButton
      Left = 320
      Top = 232
      Width = 75
      Height = 25
      Caption = 'btn1'
      TabOrder = 15
      Visible = False
      OnClick = btn1Click
    end
    object RSG1: TStringGrid
      Left = 552
      Top = 12
      Width = 447
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 1
      Visible = False
      ColWidths = (
        64
        64
        186
        161
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object RSG2: TStringGrid
      Left = 552
      Top = 116
      Width = 447
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 2
      Visible = False
      ColWidths = (
        64
        64
        186
        161
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object RSG3: TStringGrid
      Left = 552
      Top = 220
      Width = 447
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 3
      Visible = False
      ColWidths = (
        64
        64
        186
        161
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object RSG4: TStringGrid
      Left = 552
      Top = 324
      Width = 447
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 4
      Visible = False
      ColWidths = (
        64
        64
        186
        161
        64
        64
        64
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18)
    end
    object btn2: TButton
      Left = 60
      Top = 192
      Width = 75
      Height = 25
      Caption = 'TEST'
      TabOrder = 11
      Visible = False
      OnClick = btn2Click
    end
    object btn3: TButton
      Left = 160
      Top = 192
      Width = 89
      Height = 25
      Caption = #1050#1086#1087#1080' '#1082' '#1089#1077#1073#1077
      TabOrder = 12
      OnClick = btn3Click
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 49
    Width = 563
    Height = 176
    Align = alTop
    TabOrder = 1
    ExplicitWidth = 571
    object Label4: TLabel
      Left = 104
      Top = 176
      Width = 31
      Height = 13
      Caption = #1047#1072#1082#1072#1079
      Visible = False
    end
    object Label9: TLabel
      Left = 40
      Top = 184
      Width = 52
      Height = 13
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      Visible = False
    end
    object Label10: TLabel
      Left = 216
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label10'
      Visible = False
    end
    object Label11: TLabel
      Left = 264
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label11'
    end
    object SG2: TStringGrid
      Left = 8
      Top = 9
      Width = 541
      Height = 120
      ColCount = 8
      DefaultRowHeight = 18
      FixedCols = 0
      TabOrder = 3
      OnClick = SG2Click
      OnDblClick = SG2DblClick
      ColWidths = (
        102
        79
        139
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
    object Button3: TButton
      Left = 96
      Top = 128
      Width = 99
      Height = 25
      Caption = #1042' '#1101#1090#1086#1084' '#1087#1072#1082#1077#1090#1077'?'
      TabOrder = 9
      Visible = False
      OnClick = Button3Click
    end
    object Edit1: TEdit
      Left = 136
      Top = 152
      Width = 65
      Height = 21
      TabOrder = 11
      Visible = False
    end
    object Edit2: TEdit
      Left = 224
      Top = 152
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 12
      Visible = False
    end
    object ComboBox1: TComboBox
      Left = 464
      Top = 48
      Width = 145
      Height = 21
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
      TabOrder = 7
      Text = 'Edit3'
      Visible = False
    end
    object Edit4: TEdit
      Left = 464
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 8
      Text = 'Edit4'
      Visible = False
    end
    object Edit5: TEdit
      Left = 320
      Top = 72
      Width = 121
      Height = 21
      TabOrder = 5
      Text = '0'
      Visible = False
    end
    object Edit6: TEdit
      Left = 320
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 6
      Text = '0'
      Visible = False
    end
    object Memo2: TMemo
      Left = 592
      Top = 4
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo2')
      ScrollBars = ssBoth
      TabOrder = 0
      Visible = False
    end
    object mmo1: TMemo
      Left = 804
      Top = 8
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo2')
      ScrollBars = ssBoth
      TabOrder = 1
    end
    object mmo2: TMemo
      Left = 996
      Top = 8
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo2')
      ScrollBars = ssBoth
      TabOrder = 2
    end
    object chk1: TCheckBox
      Left = 368
      Top = 148
      Width = 97
      Height = 17
      Caption = #1042#1089#1077' '#1085#1072' Trumpf'
      Checked = True
      State = cbChecked
      TabOrder = 10
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 563
    Height = 49
    Align = alTop
    TabOrder = 0
    ExplicitWidth = 571
    object Label5: TLabel
      Left = 0
      Top = 0
      Width = 34
      Height = 13
      Caption = #1053#1086#1084#1077#1088
      Visible = False
    end
    object Label6: TLabel
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
    object Label7: TLabel
      Left = 184
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label7'
      Visible = False
    end
    object Label8: TLabel
      Left = 344
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label8'
      Visible = False
    end
    object Label15: TLabel
      Left = 520
      Top = 16
      Width = 38
      Height = 13
      Caption = 'Label15'
    end
  end
  object SQLConnector1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=111;Persist Security Info=True;User' +
      ' ID=testUser;Initial Catalog=Klapan;Data Source=DINAMO\SQLEXPRES' +
      'S'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 284
    Top = 24
  end
  object ADOQuery2: TADOQuery
    Connection = SQLConnector1
    Parameters = <>
    Left = 328
    Top = 24
  end
end

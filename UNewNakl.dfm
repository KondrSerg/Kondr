object FNewNakl: TFNewNakl
  Left = 429
  Top = 201
  Caption = 'FNewNakl'
  ClientHeight = 533
  ClientWidth = 1170
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
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1170
    Height = 49
    Align = alTop
    TabOrder = 0
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
    object Label2: TLabel
      Left = 96
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label2'
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 49
    Width = 1170
    Height = 152
    Align = alTop
    TabOrder = 1
    object Label9: TLabel
      Left = 104
      Top = 176
      Width = 31
      Height = 13
      Caption = #1047#1072#1082#1072#1079
      Visible = False
    end
    object Label10: TLabel
      Left = 40
      Top = 184
      Width = 52
      Height = 13
      Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
      Visible = False
    end
    object Label11: TLabel
      Left = 216
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label11'
      Visible = False
    end
    object Label15: TLabel
      Left = 264
      Top = 136
      Width = 38
      Height = 13
      Caption = 'Label15'
    end
    object SpeedButton1: TSpeedButton
      Left = 720
      Top = 120
      Width = 23
      Height = 22
      Glyph.Data = {
        26040000424D2604000000000000360000002800000012000000120000000100
        180000000000F0030000000000000000000000000000000000000000FF0000FF
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF00000000FF0080008080800080008080
        8000800080808000800080808000800080808000800080808000800080808000
        80008080800000FF00000000FF80808000800080808000800080808000800080
        80800080008080800080008080800080008080800080008080800080000000FF
        00000000FF008000808080FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080008080
        800080008080808080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080008000808080008000FFFFFF00800080808000800000
        80008080800000FF00000000FF808080008000FFFFFF80808000800080808000
        8000808080008000FFFFFF008000808080008000FFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000808080008000808080008000FFFFFF
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFF008000808080008000FFFFFF0080008080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFFFFFF
        FFFFFFFF008000FFFFFF008000808080008000808080FFFFFFFFFFFFFFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFF008000FFFFFF00
        8000808080008000808080008000808080FFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000FFFFFF008000808080008000808080
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFF808080008000808080008000808080FFFFFF8080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080FFFFFFFFFFFFFFFFFF808080008000808080FFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080008000808080008000808080008000808080008000
        8080800080008080800080008080800080008080800000FF00000000FF808080
        0080008080800080008080800080008080800080008080800080008080800080
        008080800080008080800080000000FF00000000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
        00FF0000FF0000FF0000}
      Visible = False
      OnClick = SpeedButton1Click
    end
    object SG2: TStringGrid
      Left = 8
      Top = 1
      Width = 393
      Height = 120
      ColCount = 6
      DefaultRowHeight = 18
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
      TabOrder = 0
      OnDblClick = SG2DblClick
      ColWidths = (
        102
        79
        139
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
      TabOrder = 8
      Visible = False
      OnClick = Button3Click
    end
    object Edit1: TEdit
      Left = 136
      Top = 152
      Width = 65
      Height = 21
      TabOrder = 9
      Visible = False
    end
    object Edit2: TEdit
      Left = 224
      Top = 152
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 10
      Visible = False
    end
    object ComboBox1: TComboBox
      Left = 464
      Top = 48
      Width = 145
      Height = 21
      TabOrder = 3
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
      TabOrder = 6
      Text = 'Edit3'
      Visible = False
    end
    object Edit4: TEdit
      Left = 464
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 7
      Text = 'Edit4'
      Visible = False
    end
    object Edit5: TEdit
      Left = 320
      Top = 72
      Width = 121
      Height = 21
      TabOrder = 4
      Text = '0'
      Visible = False
    end
    object Edit6: TEdit
      Left = 320
      Top = 112
      Width = 121
      Height = 21
      TabOrder = 5
      Text = '0'
      Visible = False
    end
    object Edit7: TEdit
      Left = 464
      Top = 16
      Width = 121
      Height = 21
      TabOrder = 1
      Visible = False
    end
    object SD: TStringGrid
      Left = 608
      Top = 16
      Width = 497
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 2
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
  end
  object Panel1: TPanel
    Left = 0
    Top = 201
    Width = 1170
    Height = 332
    Align = alClient
    TabOrder = 2
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
      Width = 393
      Height = 153
      DefaultRowHeight = 18
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
      TabOrder = 0
      ColWidths = (
        46
        50
        64
        51
        398)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
    object Button2: TButton
      Left = 230
      Top = 192
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 17
      OnClick = Button2Click
    end
    object Memo1: TMemo
      Left = 328
      Top = 368
      Width = 201
      Height = 97
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 23
      Visible = False
    end
    object ProgressBar1: TProgressBar
      Left = 8
      Top = 168
      Width = 297
      Height = 17
      TabOrder = 13
    end
    object SD1: TStringGrid
      Left = 496
      Top = 16
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 2
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
    object Button4: TButton
      Left = 140
      Top = 192
      Width = 75
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 16
      OnClick = Button4Click
    end
    object Button5: TButton
      Left = 408
      Top = 120
      Width = 75
      Height = 25
      Caption = 'Button5'
      TabOrder = 12
      Visible = False
      OnClick = Button5Click
    end
    object SD2: TStringGrid
      Left = 496
      Top = 128
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 3
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
    object SD3: TStringGrid
      Left = 0
      Top = 360
      Width = 489
      Height = 105
      ColCount = 16
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 21
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
    object SD4: TStringGrid
      Left = 8
      Top = 472
      Width = 481
      Height = 105
      ColCount = 17
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 25
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
    object SD5: TStringGrid
      Left = 496
      Top = 344
      Width = 609
      Height = 105
      ColCount = 22
      DefaultRowHeight = 20
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 5
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
        64
        64)
      RowHeights = (
        20
        20)
    end
    object SD6: TStringGrid
      Left = 496
      Top = 448
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 6
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
      Left = 496
      Top = 560
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 7
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
      Left = 8
      Top = 584
      Width = 481
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 26
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
    object SD9: TStringGrid
      Left = 496
      Top = 232
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 4
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
      Left = 496
      Top = 664
      Width = 609
      Height = 97
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 8
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
    object Memo2: TMemo
      Left = 112
      Top = 384
      Width = 297
      Height = 97
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 24
      Visible = False
    end
    object SSvarka: TStringGrid
      Left = 0
      Top = 256
      Width = 489
      Height = 105
      ColCount = 16
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 20
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
    object MemVG: TMemo
      Left = 500
      Top = 208
      Width = 369
      Height = 209
      ScrollBars = ssBoth
      TabOrder = 19
      Visible = False
    end
    object RSG1: TStringGrid
      Left = 434
      Top = 12
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 1
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
      Left = 436
      Top = 116
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 9
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
      Left = 436
      Top = 220
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 10
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
      Left = 436
      Top = 324
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 11
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
    object chk1: TCheckBox
      Left = 315
      Top = 168
      Width = 97
      Height = 17
      Caption = #1042#1089#1077' '#1085#1072' Trumpf'
      Checked = True
      State = cbChecked
      TabOrder = 14
    end
    object sg3: TStringGrid
      Left = 255
      Top = 324
      Width = 609
      Height = 105
      ColCount = 13
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 22
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
    object btn1: TButton
      Left = 332
      Top = 192
      Width = 75
      Height = 25
      Caption = 'Test'
      TabOrder = 18
      OnClick = btn1Click
    end
    object btn2: TButton
      Left = 44
      Top = 192
      Width = 89
      Height = 25
      Caption = #1050#1086#1087#1080' '#1082' '#1089#1077#1073#1077
      TabOrder = 15
      OnClick = btn2Click
    end
  end
  object ADOQuery2: TADOQuery
    Connection = SQLConnector1
    Parameters = <>
    Left = 328
    Top = 24
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
end

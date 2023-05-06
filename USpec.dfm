object FSpec: TFSpec
  Left = 0
  Top = 81
  Caption = 'FSpec'
  ClientHeight = 748
  ClientWidth = 1264
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
    Top = 552
    Width = 1264
    Height = 196
    Align = alBottom
    TabOrder = 1
    object Label1: TLabel
      Left = 56
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object Label2: TLabel
      Left = 128
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label2'
      Visible = False
    end
    object SpeedButton1: TSpeedButton
      Left = 16
      Top = 6
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
      OnClick = SpeedButton1Click
    end
    object Label3: TLabel
      Left = 192
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object lbl1: TLabel
      Left = 56
      Top = 24
      Width = 16
      Height = 13
      Caption = 'lbl1'
      Visible = False
    end
    object GP: TLabel
      Left = 664
      Top = 16
      Width = 15
      Height = 13
      Caption = 'GP'
    end
    object KO: TLabel
      Left = 720
      Top = 16
      Width = 15
      Height = 13
      Caption = 'KO'
    end
    object Edit1: TEdit
      Left = 296
      Top = 8
      Width = 249
      Height = 21
      TabOrder = 0
    end
    object Button1: TButton
      Left = 568
      Top = 6
      Width = 75
      Height = 25
      Caption = #1053#1072#1081#1090#1080
      TabOrder = 1
      OnClick = Button1Click
    end
    object btn1: TButton
      Left = 840
      Top = 12
      Width = 75
      Height = 25
      Caption = 'btn1'
      TabOrder = 2
      Visible = False
      OnClick = btn1Click
    end
    object Memo3: TMemo
      Left = 1
      Top = 43
      Width = 1262
      Height = 152
      Align = alBottom
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 3
      OnDblClick = Memo3DblClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1264
    Height = 552
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 0
    object SpecGrid: TStringGrid
      Left = 1
      Top = 1
      Width = 1262
      Height = 550
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      PopupMenu = PopupMenu1
      TabOrder = 0
      OnDblClick = SpecGridDblClick
      OnDrawCell = SpecGridDrawCell
      OnKeyPress = SpecGridKeyPress
      OnMouseWheelDown = SpecGridMouseWheelDown
      OnMouseWheelUp = SpecGridMouseWheelUp
      OnSelectCell = SpecGridSelectCell
      ColWidths = (
        64
        64
        64
        237
        64)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
    object ComboBox1: TComboBox
      Left = 256
      Top = 440
      Width = 97
      Height = 21
      TabOrder = 3
      Text = 'FALSE'
      Visible = False
      OnClick = ComboBox1Click
      Items.Strings = (
        'FALSE'
        'TRUE')
    end
    object Memo1: TMemo
      Left = 976
      Top = 448
      Width = 129
      Height = 241
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 2
      Visible = False
    end
    object Memo2: TMemo
      Left = 360
      Top = 416
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 1
      Visible = False
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 504
    Top = 272
    object N1: TMenuItem
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      ShortCut = 16451
      OnClick = N1Click
    end
    object N3: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1089#1090#1088#1086#1082#1091
      Visible = False
      OnClick = N3Click
    end
    object N4: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1089#1090#1088#1086#1082#1091
      Visible = False
      OnClick = N4Click
    end
    object N2: TMenuItem
      Caption = #1042' '#1087#1077#1088#1077#1079#1072#1087#1091#1089#1082
      OnClick = N2Click
    end
  end
  object ADOQuery3: TADOQuery
    Parameters = <>
    Left = 537
    Top = 188
  end
  object ADOConnection: TADOConnection
    Left = 657
    Top = 184
  end
end

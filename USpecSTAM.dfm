object FSTAM: TFSTAM
  Left = 218
  Top = 482
  Caption = 'FSTAM'
  ClientHeight = 589
  ClientWidth = 1109
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
  object lbl1: TLabel
    Left = 196
    Top = 168
    Width = 16
    Height = 13
    Caption = 'lbl1'
  end
  object SpecGrid: TStringGrid
    Left = 0
    Top = 0
    Width = 1109
    Height = 440
    Align = alClient
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    PopupMenu = pm1
    TabOrder = 0
    OnDrawCell = SpecGridDrawCell
    OnMouseWheelDown = SpecGridMouseWheelDown
    OnMouseWheelUp = SpecGridMouseWheelUp
    OnSelectCell = SpecGridSelectCell
    ExplicitHeight = 506
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
  object cbb1: TComboBox
    Left = 256
    Top = 440
    Width = 97
    Height = 21
    TabOrder = 5
    Text = 'FALSE'
    Visible = False
    OnClick = cbb1Click
    Items.Strings = (
      'FALSE'
      'TRUE')
  end
  object mmo1: TMemo
    Left = 368
    Top = 416
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 3
    Visible = False
  end
  object mmo2: TMemo
    Left = 576
    Top = 416
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 4
    Visible = False
  end
  object pnl1: TPanel
    Left = 0
    Top = 440
    Width = 1109
    Height = 149
    Align = alBottom
    TabOrder = 1
    object lbl2: TLabel
      Left = 56
      Top = 8
      Width = 16
      Height = 13
      Caption = 'lbl2'
    end
    object lbl3: TLabel
      Left = 128
      Top = 8
      Width = 16
      Height = 13
      Caption = 'lbl3'
    end
    object btn1: TSpeedButton
      Left = 784
      Top = 11
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
      OnClick = btn1Click
    end
    object lbl4: TLabel
      Left = 192
      Top = 8
      Width = 16
      Height = 13
      Caption = 'lbl4'
    end
    object lbl5: TLabel
      Left = 56
      Top = 24
      Width = 16
      Height = 13
      Caption = 'lbl1'
    end
    object Label3: TLabel
      Left = 764
      Top = 48
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object edt1: TEdit
      Left = 296
      Top = 8
      Width = 249
      Height = 21
      TabOrder = 0
    end
    object btn2: TButton
      Left = 568
      Top = 8
      Width = 75
      Height = 25
      Caption = #1053#1072#1081#1090#1080
      TabOrder = 1
      OnClick = btn2Click
    end
    object Memo1: TMemo
      Left = 0
      Top = 35
      Width = 466
      Height = 117
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 2
    end
  end
  object Memo2: TMemo
    Left = 40
    Top = 408
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo2')
    ScrollBars = ssHorizontal
    TabOrder = 2
    Visible = False
  end
  object pm1: TPopupMenu
    Left = 504
    Top = 272
    object N1: TMenuItem
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      ShortCut = 16451
      OnClick = N1Click
    end
  end
end

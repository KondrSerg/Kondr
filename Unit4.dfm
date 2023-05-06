object Form4: TForm4
  Left = 108
  Top = 286
  Width = 1050
  Height = 184
  Caption = 'Form4'
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
    Left = 432
    Top = 88
    Width = 32
    Height = 13
    Caption = 'Label3'
  end
  object Label1: TLabel
    Left = 320
    Top = 88
    Width = 32
    Height = 13
    Caption = 'Label1'
  end
  object SpecGrid: TStringGrid
    Left = 0
    Top = 0
    Width = 1042
    Height = 81
    Align = alTop
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    ScrollBars = ssHorizontal
    TabOrder = 0
    OnSelectCell = SpecGridSelectCell
  end
  object ComboBox1: TComboBox
    Left = 400
    Top = 4
    Width = 97
    Height = 21
    ItemHeight = 13
    TabOrder = 1
    Text = 'FALSE'
    Visible = False
    OnClick = ComboBox1Click
    Items.Strings = (
      'FALSE'
      'TRUE')
  end
  object Button2: TButton
    Left = 651
    Top = 104
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button1: TButton
    Left = 755
    Top = 104
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 3
    OnClick = Button1Click
  end
end

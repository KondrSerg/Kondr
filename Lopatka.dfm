object FLopatka: TFLopatka
  Left = 87
  Top = 320
  Width = 1193
  Height = 640
  Caption = #1047#1072#1076#1072#1085#1080#1077
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
  object Splitter1: TSplitter
    Left = 0
    Top = 249
    Width = 1185
    Height = 3
    Cursor = crVSplit
    Align = alTop
  end
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 1185
    Height = 249
    Align = alTop
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 512
    Width = 1185
    Height = 94
    Align = alBottom
    TabOrder = 2
    object Label1: TLabel
      Left = 32
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object Button1: TButton
      Left = 24
      Top = 56
      Width = 75
      Height = 25
      Caption = 'Excel'
      TabOrder = 0
      OnClick = Button1Click
    end
  end
  object StringGrid2: TStringGrid
    Left = 0
    Top = 252
    Width = 1185
    Height = 260
    Align = alClient
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 1
  end
end

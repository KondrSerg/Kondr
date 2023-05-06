object FLop: TFLop
  Left = 186
  Top = -1
  Width = 1067
  Height = 995
  Caption = 'Lopatki'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 857
    Height = 281
    Caption = 'Panel1'
    TabOrder = 0
    object Label1: TLabel
      Left = 72
      Top = 272
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object StringGrid1: TStringGrid
      Left = 1
      Top = 1
      Width = 855
      Height = 279
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
      ColWidths = (
        64
        137
        156
        64
        64)
    end
  end
  object Panel2: TPanel
    Left = -8
    Top = 288
    Width = 857
    Height = 441
    Caption = 'Panel1'
    TabOrder = 1
    object Label2: TLabel
      Left = 72
      Top = 272
      Width = 32
      Height = 13
      Caption = 'Label2'
    end
    object StringGrid2: TStringGrid
      Left = 1
      Top = 1
      Width = 855
      Height = 439
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
      ColWidths = (
        64
        210
        135
        150
        169)
    end
  end
  object Panel3: TPanel
    Left = 8
    Top = 736
    Width = 857
    Height = 225
    Caption = 'Panel1'
    TabOrder = 2
    object Label3: TLabel
      Left = 72
      Top = 272
      Width = 32
      Height = 13
      Caption = 'Label2'
    end
    object StringGrid3: TStringGrid
      Left = 1
      Top = 1
      Width = 855
      Height = 223
      Align = alClient
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
    end
  end
end

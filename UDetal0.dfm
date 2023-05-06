object FDetal0: TFDetal0
  Left = 206
  Top = 411
  Width = 1003
  Height = 83
  Caption = 'FDetal0'
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
  object Label1: TLabel
    Left = 304
    Top = 128
    Width = 32
    Height = 13
    Caption = 'Label1'
  end
  object Label2: TLabel
    Left = 560
    Top = 344
    Width = 32
    Height = 13
    Caption = 'Label2'
  end
  object SpecGrid: TStringGrid
    Left = 0
    Top = 0
    Width = 978
    Height = 357
    Align = alClient
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ScrollBars = ssHorizontal
    TabOrder = 0
  end
end

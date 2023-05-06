object FPDM: TFPDM
  Left = 253
  Top = 110
  Width = 449
  Height = 597
  Caption = 'FPDM'
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 441
    Height = 41
    Align = alTop
    TabOrder = 0
  end
  object Panel2: TPanel
    Left = 0
    Top = 41
    Width = 441
    Height = 481
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object SG1: TStringGrid
      Left = 1
      Top = 1
      Width = 439
      Height = 479
      Align = alClient
      ColCount = 2
      DefaultRowHeight = 18
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
      TabOrder = 0
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 522
    Width = 441
    Height = 41
    Align = alBottom
    TabOrder = 2
  end
end

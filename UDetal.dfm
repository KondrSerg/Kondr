object FDetal: TFDetal
  Left = 160
  Top = 369
  Width = 465
  Height = 190
  Caption = 'FDetal'
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
    Left = 16
    Top = 8
    Width = 32
    Height = 13
    Caption = 'Label1'
  end
  object Label2: TLabel
    Left = 200
    Top = 8
    Width = 32
    Height = 13
    Caption = 'Label2'
  end
  object StringGrid3: TStringGrid
    Left = 480
    Top = 64
    Width = 369
    Height = 97
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    PopupMenu = PopupMenu1
    TabOrder = 1
    Visible = False
  end
  object Edit1: TEdit
    Left = 16
    Top = 40
    Width = 121
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 368
    Top = 104
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 3
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 256
    Top = 104
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 2
    OnClick = Button2Click
  end
  object PopupMenu1: TPopupMenu
    Left = 52
    Top = 88
    object True1: TMenuItem
      Caption = 'True'
    end
    object False1: TMenuItem
      Caption = 'False'
    end
  end
end

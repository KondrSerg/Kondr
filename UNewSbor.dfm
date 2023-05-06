object FNewSbor: TFNewSbor
  Left = 316
  Top = 147
  Caption = 'FNewSbor'
  ClientHeight = 601
  ClientWidth = 1064
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
  object SGSB: TStringGrid
    Left = 0
    Top = 41
    Width = 1064
    Height = 498
    Align = alClient
    ColCount = 7
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    PopupMenu = PopupMenu1
    TabOrder = 0
    OnKeyPress = SGSBKeyPress
    ExplicitWidth = 604
    ColWidths = (
      64
      172
      171
      194
      106
      105
      64)
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1064
    Height = 41
    Align = alTop
    TabOrder = 1
    ExplicitWidth = 604
  end
  object Panel2: TPanel
    Left = 0
    Top = 539
    Width = 1064
    Height = 62
    Align = alBottom
    TabOrder = 2
    ExplicitWidth = 604
    object Button1: TButton
      Left = 520
      Top = 24
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 432
      Top = 24
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 1
      OnClick = Button2Click
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 216
    Top = 296
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N2Click
    end
  end
end

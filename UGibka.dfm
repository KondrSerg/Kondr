object FGibka: TFGibka
  Left = 0
  Top = 0
  Caption = 'FGibka'
  ClientHeight = 559
  ClientWidth = 961
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 961
    Height = 41
    Align = alTop
    Caption = 'pnl1'
    TabOrder = 0
  end
  object pnl2: TPanel
    Left = 0
    Top = 518
    Width = 961
    Height = 41
    Align = alBottom
    Caption = 'pnl2'
    TabOrder = 1
  end
  object pgc1: TPageControl
    Left = 0
    Top = 41
    Width = 961
    Height = 477
    ActivePage = ts1
    Align = alClient
    TabOrder = 2
    object ts1: TTabSheet
      Caption = #1043#1080#1073#1082#1072
      object Label1: TLabel
        Left = 3
        Top = 40
        Width = 51
        Height = 13
        Caption = #1054#1089#1085#1086#1074#1085#1099#1077
      end
      object Label2: TLabel
        Left = 428
        Top = 40
        Width = 88
        Height = 13
        Caption = #1044#1086#1087#1086#1083#1085#1080#1090#1077#1083#1100#1085#1099#1077
      end
      object StringGrid1: TStringGrid
        Left = 3
        Top = 59
        Width = 399
        Height = 278
        FixedCols = 0
        TabOrder = 0
        OnSelectCell = StringGrid1SelectCell
        ColWidths = (
          64
          64
          64
          64
          64)
        RowHeights = (
          24
          24
          24
          24
          24)
      end
      object pnl3: TPanel
        Left = 0
        Top = 395
        Width = 953
        Height = 54
        Align = alBottom
        UseDockManager = False
        ParentBackground = False
        TabOrder = 1
        object Button1: TButton
          Left = 384
          Top = 16
          Width = 75
          Height = 25
          Caption = #1054#1073#1085#1086#1074#1080#1090#1100
          TabOrder = 0
          OnClick = Button1Click
        end
        object Button2: TButton
          Left = 488
          Top = 16
          Width = 75
          Height = 25
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 1
          OnClick = Button2Click
        end
      end
      object StringGrid2: TStringGrid
        Left = 408
        Top = 59
        Width = 542
        Height = 278
        ColCount = 6
        FixedCols = 0
        PopupMenu = pm1
        TabOrder = 2
        OnKeyPress = StringGrid2KeyPress
        OnSelectCell = StringGrid2SelectCell
        ColWidths = (
          64
          64
          64
          64
          64
          64)
        RowHeights = (
          24
          24
          24
          24
          24)
      end
      object ComboBox1: TComboBox
        Left = 632
        Top = 128
        Width = 145
        Height = 21
        TabOrder = 3
        Text = 'ComboBox1'
        Visible = False
      end
    end
    object ts2: TTabSheet
      Caption = #1056#1077#1079#1082#1072
      ImageIndex = 1
    end
  end
  object pm1: TPopupMenu
    Left = 780
    Top = 353
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N1Click
    end
  end
end

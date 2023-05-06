object FFile: TFFile
  Left = 0
  Top = 0
  Caption = 'FFile'
  ClientHeight = 468
  ClientWidth = 632
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ListView1: TListView
    Left = 15
    Top = 79
    Width = 530
    Height = 314
    Columns = <
      item
        AutoSize = True
        Caption = #1048#1084#1103
      end>
    RowSelect = True
    PopupMenu = pm4
    TabOrder = 0
    ViewStyle = vsReport
    OnDblClick = ListView1DblClick
  end
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 632
    Height = 73
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 424
      Top = 48
      Width = 31
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object btn2: TSpeedButton
      Left = 96
      Top = 40
      Width = 23
      Height = 22
    end
    object Edit3: TEdit
      Left = 15
      Top = 12
      Width = 329
      Height = 21
      TabOrder = 0
      Text = 'V:\'#1050#1058#1054'\_'#1056#1072#1079#1074#1077#1088#1090#1082#1080' new\'#1057#1058#1040#1052' 18.05.2018\*.*'
    end
    object btn1: TButton
      Left = 15
      Top = 39
      Width = 75
      Height = 25
      Caption = #1042#1077#1088#1085#1091#1090#1100#1089#1103
      TabOrder = 1
      OnClick = btn1Click
    end
  end
  object Memo1: TMemo
    Left = 15
    Top = 399
    Width = 530
    Height = 64
    ScrollBars = ssBoth
    TabOrder = 2
  end
  object dlgSave1: TSaveDialog
    Left = 476
    Top = 224
  end
  object dlgOpen2: TOpenDialog
    Left = 508
    Top = 224
  end
  object pm4: TPopupMenu
    Left = 340
    Top = 224
    object N12: TMenuItem
      Caption = #1054#1090#1082#1088#1099#1090#1100
      Visible = False
      OnClick = N12Click
    end
    object N9: TMenuItem
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      OnClick = N9Click
    end
    object N10: TMenuItem
      Caption = #1042' '#1072#1088#1093#1080#1074
      Visible = False
    end
    object N11: TMenuItem
      Caption = #1057#1086#1079#1076#1072#1090#1100
      OnClick = N11Click
    end
  end
end

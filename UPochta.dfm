object FPochta: TFPochta
  Left = 0
  Top = 0
  Caption = 'FPochta'
  ClientHeight = 556
  ClientWidth = 997
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object spl1: TSplitter
    Left = 361
    Top = 89
    Height = 467
    ExplicitLeft = 368
    ExplicitTop = 248
    ExplicitHeight = 100
  end
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 997
    Height = 89
    Align = alTop
    TabOrder = 0
    object Button1: TButton
      Left = 344
      Top = 16
      Width = 75
      Height = 25
      Caption = #1053#1072#1081#1090#1080
      TabOrder = 0
      OnClick = Button1Click
    end
    object dtp1: TDateTimePicker
      Left = 16
      Top = 16
      Width = 186
      Height = 21
      Date = 43622.427164236120000000
      Time = 43622.427164236120000000
      TabOrder = 1
    end
  end
  object lst1: TListBox
    Left = 0
    Top = 89
    Width = 361
    Height = 467
    Align = alLeft
    ItemHeight = 13
    TabOrder = 1
    OnDblClick = lst1DblClick
  end
  object redt1: TRichEdit
    Left = 364
    Top = 89
    Width = 633
    Height = 467
    Align = alClient
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Lines.Strings = (
      'redt1')
    ParentFont = False
    PopupMenu = pm1
    ScrollBars = ssBoth
    TabOrder = 2
    Zoom = 100
  end
  object pm1: TPopupMenu
    Left = 488
    Top = 224
    object N2: TMenuItem
      Caption = #1047#1072#1074#1053#1086#1084#1077#1088'('#1050#1083#1072#1087#1072#1085#1072')'
      OnClick = N2Click
    end
    object N1: TMenuItem
      Caption = #1053#1086#1084#1077#1088'('#1041#1088#1072#1082')'
      OnClick = N1Click
    end
  end
end

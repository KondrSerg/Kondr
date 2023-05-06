object FServer: TFServer
  Left = 457
  Top = 360
  Caption = 'FServer'
  ClientHeight = 471
  ClientWidth = 660
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 272
    Top = 152
    Width = 75
    Height = 25
    Caption = 'Create'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 360
    Top = 152
    Width = 75
    Height = 25
    Caption = 'Send'
    TabOrder = 3
    OnClick = Button2Click
  end
  object Memo1: TMemo
    Left = 56
    Top = 56
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 1
  end
  object Memo2: TMemo
    Left = 264
    Top = 48
    Width = 185
    Height = 89
    Lines.Strings = (
      'Memo2')
    ScrollBars = ssBoth
    TabOrder = 0
  end
  object ListBox1: TListBox
    Left = 80
    Top = 200
    Width = 345
    Height = 113
    ItemHeight = 13
    TabOrder = 5
  end
  object Edit1: TEdit
    Left = 120
    Top = 168
    Width = 121
    Height = 21
    TabOrder = 4
    Text = '123454'
  end
  object Button3: TButton
    Left = 320
    Top = 328
    Width = 75
    Height = 25
    Caption = 'KilAll'
    TabOrder = 6
    OnClick = Button3Click
  end
end

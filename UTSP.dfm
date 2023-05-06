object FTSP: TFTSP
  Left = 122
  Top = 55
  Caption = 'FTSP'
  ClientHeight = 322
  ClientWidth = 502
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 232
    Top = 264
    Width = 75
    Height = 25
    Caption = 'Create'
    TabOrder = 6
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 320
    Top = 264
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 7
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 152
    Top = 200
    Width = 75
    Height = 25
    Caption = 'Send'
    TabOrder = 5
    OnClick = Button3Click
  end
  object Memo1: TMemo
    Left = 28
    Top = 35
    Width = 185
    Height = 89
    ScrollBars = ssBoth
    TabOrder = 2
  end
  object ListBox1: TListBox
    Left = 236
    Top = 27
    Width = 177
    Height = 97
    ItemHeight = 13
    TabOrder = 1
  end
  object Memo2: TMemo
    Left = 244
    Top = 163
    Width = 185
    Height = 89
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object Edit1: TEdit
    Left = 32
    Top = 168
    Width = 193
    Height = 21
    TabOrder = 4
    Text = 'Edit1'
  end
  object Edit2: TEdit
    Left = 32
    Top = 8
    Width = 121
    Height = 21
    TabOrder = 0
    Text = '192.168.7.146:12345'
  end
end

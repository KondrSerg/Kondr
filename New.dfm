object FNew: TFNew
  Left = 321
  Top = 182
  Width = 600
  Height = 635
  Caption = 'FNew'
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
    Width = 78
    Height = 13
    Caption = #1047#1072' '#1082#1072#1082#1086#1077' '#1095#1080#1089#1083#1086
  end
  object DateTimePicker1: TDateTimePicker
    Left = 8
    Top = 32
    Width = 186
    Height = 21
    Date = 41747.637392511580000000
    Time = 41747.637392511580000000
    TabOrder = 0
  end
  object Button1: TButton
    Left = 384
    Top = 560
    Width = 75
    Height = 25
    Caption = #1047#1072#1083#1080#1090#1100
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 480
    Top = 560
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 2
    OnClick = Button2Click
  end
  object Memo1: TMemo
    Left = 8
    Top = 64
    Width = 553
    Height = 201
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object Memo2: TMemo
    Left = 8
    Top = 280
    Width = 553
    Height = 201
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 4
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=111;Persist Security Info=True;User' +
      ' ID=semin;Initial Catalog=VezaDB;Data Source=192.168.5.2'
    Provider = 'SQLOLEDB.1'
    Left = 264
    Top = 24
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 304
    Top = 24
  end
end

object FSborAll: TFSborAll
  Left = 297
  Top = 248
  Caption = #1057#1073#1086#1088#1097#1080#1082
  ClientHeight = 350
  ClientWidth = 738
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 6
    Height = 13
    Caption = '0'
  end
  object Label2: TLabel
    Left = 40
    Top = 24
    Width = 43
    Height = 13
    Caption = #1057#1073#1086#1088#1082#1072'2'
  end
  object Label3: TLabel
    Left = 40
    Top = 70
    Width = 62
    Height = 13
    Caption = #1055#1086#1076#1089#1073#1086#1088#1082#1072'1'
  end
  object Label4: TLabel
    Left = 40
    Top = 117
    Width = 68
    Height = 13
    Caption = #1057#1073#1086#1088#1051#1086#1087#1072#1090#1086#1082
  end
  object Label5: TLabel
    Left = 40
    Top = 163
    Width = 67
    Height = 13
    Caption = #1057#1073#1086#1088#1050#1086#1088#1087#1091#1089#1072
  end
  object Label6: TLabel
    Left = 440
    Top = 24
    Width = 75
    Height = 13
    Caption = #1057#1073#1086#1088#1082#1072' '#1043#1086#1090#1086#1074#1072
  end
  object Label7: TLabel
    Left = 32
    Top = 240
    Width = 32
    Height = 13
    Caption = 'Label7'
  end
  object ComboBox1: TComboBox
    Left = 32
    Top = 43
    Width = 313
    Height = 21
    TabOrder = 0
    OnKeyPress = ComboBox1KeyPress
  end
  object Button1: TButton
    Left = 144
    Top = 280
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 270
    Top = 280
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 2
    OnClick = Button2Click
    OnKeyPress = Button2KeyPress
  end
  object CBB1: TComboBox
    Left = 32
    Top = 88
    Width = 313
    Height = 21
    TabOrder = 3
    Visible = False
    OnKeyPress = ComboBox1KeyPress
  end
  object CBB2: TComboBox
    Left = 32
    Top = 136
    Width = 313
    Height = 21
    TabOrder = 4
    OnKeyPress = ComboBox1KeyPress
  end
  object CBB3: TComboBox
    Left = 32
    Top = 184
    Width = 313
    Height = 21
    TabOrder = 5
    OnKeyPress = ComboBox1KeyPress
  end
  object btn1: TBitBtn
    Left = 368
    Top = 88
    Width = 33
    Height = 25
    Caption = 'V'
    TabOrder = 6
    OnClick = btn1Click
  end
  object btn2: TBitBtn
    Left = 368
    Top = 134
    Width = 33
    Height = 25
    Caption = 'V'
    TabOrder = 7
    OnClick = btn2Click
  end
  object btn3: TBitBtn
    Left = 368
    Top = 182
    Width = 33
    Height = 25
    Caption = 'V'
    TabOrder = 8
    OnClick = btn3Click
  end
  object dtp1: TDateTimePicker
    Left = 440
    Top = 43
    Width = 186
    Height = 21
    Date = 45028.379801180560000000
    Time = 45028.379801180560000000
    TabOrder = 9
  end
  object dtp2: TDateTimePicker
    Left = 440
    Top = 88
    Width = 186
    Height = 21
    Date = 45028.379801180560000000
    Time = 45028.379801180560000000
    TabOrder = 10
  end
  object dtp3: TDateTimePicker
    Left = 440
    Top = 134
    Width = 186
    Height = 21
    Date = 45028.379801180560000000
    Time = 45028.379801180560000000
    TabOrder = 11
  end
  object dtp4: TDateTimePicker
    Left = 440
    Top = 184
    Width = 186
    Height = 21
    Date = 45028.379801180560000000
    Time = 45028.379801180560000000
    TabOrder = 12
  end
end

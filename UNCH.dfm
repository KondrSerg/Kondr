object Fnch: TFnch
  Left = 0
  Top = 0
  Caption = 'Fnch'
  ClientHeight = 472
  ClientWidth = 791
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
  object pnl1: TPanel
    Left = 40
    Top = 64
    Width = 721
    Height = 345
    TabOrder = 0
    object Label1: TLabel
      Left = 24
      Top = 20
      Width = 60
      Height = 13
      Caption = #1048#1089#1087#1086#1083#1085#1077#1085#1080#1077
    end
    object Label2: TLabel
      Left = 231
      Top = 20
      Width = 81
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086' '#1051#1086#1087#1072#1090#1086#1082
    end
    object Label3: TLabel
      Left = 348
      Top = 20
      Width = 132
      Height = 13
      Caption = #1055#1086#1083#1091#1087#1077#1088#1080#1084#1077#1090#1088' '#1080#1083#1080' '#1088#1072#1079#1084#1077#1088
    end
    object Label4: TLabel
      Left = 504
      Top = 20
      Width = 87
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086' '#1087#1088#1080#1074#1086#1076#1086#1074
    end
    object Label5: TLabel
      Left = 616
      Top = 20
      Width = 31
      Height = 13
      Caption = 'Label5'
    end
    object Label6: TLabel
      Left = 24
      Top = 109
      Width = 37
      Height = 13
      Caption = #1057#1073#1086#1088#1082#1072
    end
    object Label7: TLabel
      Left = 184
      Top = 109
      Width = 55
      Height = 13
      Caption = #1055#1086#1076#1089#1073#1086#1088#1082#1072
    end
    object Button1: TButton
      Left = 504
      Top = 296
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 400
      Top = 296
      Width = 75
      Height = 25
      Caption = 'Select'
      TabOrder = 1
      OnClick = Button2Click
    end
    object ComboBox1: TComboBox
      Left = 24
      Top = 51
      Width = 81
      Height = 21
      TabOrder = 2
      Items.Strings = (
        #1054','#1047
        #1044)
    end
    object ComboBox2: TComboBox
      Left = 348
      Top = 51
      Width = 132
      Height = 21
      TabOrder = 3
    end
    object rg1: TRadioGroup
      Left = 128
      Top = 20
      Width = 97
      Height = 77
      Caption = #1050#1086#1083'-'#1074#1086' '#1060#1083#1072#1085#1094#1077#1074
      ItemIndex = 2
      Items.Strings = (
        '0'
        '1'
        '2')
      TabOrder = 4
    end
    object Edit1: TEdit
      Left = 231
      Top = 51
      Width = 82
      Height = 21
      NumbersOnly = True
      TabOrder = 5
      Text = '0'
    end
    object Edit2: TEdit
      Left = 504
      Top = 51
      Width = 82
      Height = 21
      NumbersOnly = True
      TabOrder = 6
      Text = '0'
    end
    object Edit3: TEdit
      Left = 24
      Top = 128
      Width = 121
      Height = 21
      TabOrder = 7
      Text = 'Edit3'
    end
    object Edit4: TEdit
      Left = 184
      Top = 128
      Width = 121
      Height = 21
      TabOrder = 8
      Text = 'Edit4'
    end
    object Memo1: TMemo
      Left = 348
      Top = 128
      Width = 231
      Height = 153
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 9
    end
    object Edit5: TEdit
      Left = 616
      Top = 51
      Width = 57
      Height = 21
      TabOrder = 10
      Text = '0'
    end
  end
end

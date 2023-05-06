object FPrim: TFPrim
  Left = 148
  Top = 268
  Width = 581
  Height = 306
  Caption = 'FPrim'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 231
    Width = 573
    Height = 41
    Align = alBottom
    TabOrder = 0
    object Label1: TLabel
      Left = 24
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object Label2: TLabel
      Left = 120
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label2'
      Visible = False
    end
    object Label3: TLabel
      Left = 192
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label3'
      Visible = False
    end
    object Label4: TLabel
      Left = 248
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label4'
      Visible = False
    end
    object Label5: TLabel
      Left = 296
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label5'
      Visible = False
    end
    object Button1: TButton
      Left = 464
      Top = 8
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 376
      Top = 8
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 1
      OnClick = Button2Click
    end
  end
  object Memo1: TMemo
    Left = 0
    Top = 0
    Width = 573
    Height = 231
    Align = alClient
    TabOrder = 1
  end
end

object FBriket: TFBriket
  Left = 491
  Top = 261
  Caption = 'FBriket'
  ClientHeight = 104
  ClientWidth = 198
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label3: TLabel
    Left = 16
    Top = 96
    Width = 3
    Height = 13
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 198
    Height = 104
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 206
    ExplicitHeight = 109
    object Label4: TLabel
      Left = 8
      Top = 8
      Width = 34
      Height = 13
      Caption = #1053#1086#1084#1077#1088
    end
    object Label1: TLabel
      Left = 64
      Top = 8
      Width = 32
      Height = 13
      Caption = 'Label1'
      Visible = False
    end
    object GradientLabel1: TGradientLabel
      Left = 122
      Top = 32
      Width = 65
      Height = 17
      AutoSize = False
      Caption = 'GradientLabel1'
      Color = clYellow
      ParentColor = False
      ColorTo = clAqua
      EllipsType = etNone
      GradientType = gtFullHorizontal
      GradientDirection = gdLeftToRight
      Indent = 0
      Orientation = goHorizontal
      TransparentText = False
      VAlignment = vaTop
      Version = '1.2.0.1'
    end
    object GradientLabel2: TGradientLabel
      Left = 104
      Top = 32
      Width = 17
      Height = 17
      AutoSize = False
      Caption = #8470
      Color = clAqua
      ParentColor = False
      ColorTo = clYellow
      EllipsType = etNone
      GradientType = gtFullHorizontal
      GradientDirection = gdLeftToRight
      Indent = 0
      Orientation = goHorizontal
      TransparentText = False
      VAlignment = vaTop
      Version = '1.2.0.1'
    end
    object Button1: TButton
      Left = 8
      Top = 64
      Width = 81
      Height = 25
      Caption = 'Ok'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Edit3: TEdit
      Left = 8
      Top = 32
      Width = 89
      Height = 21
      TabOrder = 1
      OnKeyPress = Edit3KeyPress
    end
    object Button2: TButton
      Left = 104
      Top = 64
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
      OnClick = Button2Click
      OnKeyPress = Button2KeyPress
    end
  end
end

object FS: TFS
  Left = 291
  Top = 306
  Caption = 'FS'
  ClientHeight = 422
  ClientWidth = 950
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
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 950
    Height = 105
    Align = alTop
    TabOrder = 0
    object jvgrdnt1: TJvGradient
      Left = 1
      Top = 1
      Width = 948
      Height = 103
      StartColor = clAqua
      EndColor = clWhite
      Steps = 50
      ExplicitWidth = 956
    end
    object lbl1: TGradientLabel
      Left = 192
      Top = 24
      Width = 105
      Height = 29
      AutoSize = False
      Caption = '0'
      Color = clAqua
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Transparent = True
      EllipsType = etNone
      GradientType = gtFullHorizontal
      GradientDirection = gdLeftToRight
      Indent = 0
      Orientation = goHorizontal
      TransparentText = False
      VAlignment = vaTop
      Version = '1.2.0.1'
    end
    object lbl2: TGradientLabel
      Left = 20
      Top = 24
      Width = 161
      Height = 29
      AutoSize = False
      Caption = #1053#1086#1084#1077#1088' '#1085#1072#1082#1083#1072#1076#1085#1086#1081
      Color = clCream
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Transparent = True
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
    object lbl3: TLabel
      Left = 20
      Top = 68
      Width = 91
      Height = 13
      Caption = #1053'\'#1095' '#1057#1074#1072#1088#1082#1072' '#1042#1089#1077#1075#1086
    end
    object lbl4: TLabel
      Left = 124
      Top = 68
      Width = 6
      Height = 13
      Caption = '0'
    end
    object lbl5: TLabel
      Left = 248
      Top = 68
      Width = 90
      Height = 13
      Caption = #1053'\'#1095' '#1057#1073#1086#1088#1082#1072' '#1074#1089#1077#1075#1086
    end
    object lbl6: TLabel
      Left = 352
      Top = 68
      Width = 16
      Height = 13
      Caption = 'lbl6'
    end
    object lbl7: TLabel
      Left = 707
      Top = 12
      Width = 74
      Height = 13
      Caption = #1055#1083#1072#1085#1080#1088#1086#1074#1072#1085#1080#1077
    end
    object dtp1: TDateTimePicker
      Left = 708
      Top = 28
      Width = 186
      Height = 21
      Date = 42025.356357210650000000
      Time = 42025.356357210650000000
      TabOrder = 4
      OnClick = dtp1Click
    end
    object btn1: TButton
      Left = 404
      Top = 68
      Width = 75
      Height = 25
      Caption = 'btn1'
      TabOrder = 5
      Visible = False
      OnClick = btn1Click
    end
    object btn2: TButton
      Left = 320
      Top = 24
      Width = 75
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 1
      OnClick = btn2Click
    end
    object btn3: TButton
      Left = 416
      Top = 24
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
      OnClick = btn3Click
    end
    object btn4: TButton
      Left = 512
      Top = 68
      Width = 75
      Height = 25
      Caption = #1047#1072#1075#1086#1090#1086#1074#1082#1072
      TabOrder = 6
      Visible = False
    end
    object Edit: TEdit
      Left = 176
      Top = 24
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 0
      Text = 'Edit'
    end
    object btn5: TButton
      Left = 512
      Top = 24
      Width = 75
      Height = 25
      Caption = 'TEST'
      TabOrder = 3
      Visible = False
    end
  end
  object SG21: TAdvStringGrid
    Left = 0
    Top = 105
    Width = 950
    Height = 317
    Cursor = crDefault
    Align = alClient
    ColCount = 31
    DrawingStyle = gdsClassic
    FixedCols = 5
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goThumbTracking]
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 1
    OnSetEditText = SG21SetEditText
    ActiveRowShow = True
    ActiveRowColor = clSkyBlue
    HoverRowCells = [hcNormal, hcSelected]
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    AutoNumAlign = True
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'MS Sans Serif'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'MS Sans Serif'
    FilterDropDown.Font.Style = []
    FilterDropDown.TextChecked = 'Checked'
    FilterDropDown.TextUnChecked = 'Unchecked'
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Clear')
    FixedRightCols = 25
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = [fsBold]
    FloatFormat = '%.2f'
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    MouseActions.SizeFixedCol = True
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = DEFAULT_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'MS Sans Serif'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'MS Sans Serif'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'MS Sans Serif'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'MS Sans Serif'
    PrintSettings.FooterFont.Style = []
    PrintSettings.PageNumSep = '/'
    ScrollSynch = True
    SearchFooter.FindNextCaption = 'Find &next'
    SearchFooter.FindPrevCaption = 'Find &previous'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'MS Sans Serif'
    SearchFooter.Font.Style = []
    SearchFooter.HighLightCaption = 'Highlight'
    SearchFooter.HintClose = 'Close'
    SearchFooter.HintFindNext = 'Find next occurrence'
    SearchFooter.HintFindPrev = 'Find previous occurrence'
    SearchFooter.HintHighlight = 'Highlight occurrences'
    SearchFooter.MatchCaseCaption = 'Match case'
    SortSettings.DefaultFormat = ssAutomatic
    Version = '8.1.3.0'
    ColWidths = (
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      75
      135
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64)
    RowHeights = (
      22
      22
      22
      22
      22
      22
      22
      22
      22
      22)
  end
end

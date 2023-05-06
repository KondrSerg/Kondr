object FRabota: TFRabota
  Left = 0
  Top = 0
  Caption = 'FRabota'
  ClientHeight = 702
  ClientWidth = 1126
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
    Left = 0
    Top = 0
    Width = 1126
    Height = 121
    Align = alTop
    TabOrder = 0
    object btn1: TSpeedButton
      Left = 16
      Top = 27
      Width = 23
      Height = 22
      Glyph.Data = {
        26040000424D2604000000000000360000002800000012000000120000000100
        180000000000F0030000000000000000000000000000000000000000FF0000FF
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF00000000FF0080008080800080008080
        8000800080808000800080808000800080808000800080808000800080808000
        80008080800000FF00000000FF80808000800080808000800080808000800080
        80800080008080800080008080800080008080800080008080800080000000FF
        00000000FF008000808080FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080008080
        800080008080808080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080008000808080008000FFFFFF00800080808000800000
        80008080800000FF00000000FF808080008000FFFFFF80808000800080808000
        8000808080008000FFFFFF008000808080008000FFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000808080008000808080008000FFFFFF
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFFFFFFFF008000808080008000FFFFFF0080008080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFFFFFF
        FFFFFFFF008000FFFFFF008000808080008000808080FFFFFFFFFFFFFFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFF008000FFFFFF00
        8000808080008000808080008000808080FFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080FFFFFF008000FFFFFF008000808080008000808080
        008000808080008000808080FFFFFF0080008080800000FF00000000FF808080
        008000FFFFFF808080008000808080008000808080FFFFFF8080800080008080
        80008000FFFFFF8080800080000000FF00000000FF008000808080FFFFFF0080
        00808080008000808080FFFFFFFFFFFFFFFFFF808080008000808080FFFFFF00
        80008080800000FF00000000FF808080008000FFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8080800080000000FF
        00000000FF008000808080008000808080008000808080008000808080008000
        8080800080008080800080008080800080008080800000FF00000000FF808080
        0080008080800080008080800080008080800080008080800080008080800080
        008080800080008080800080000000FF00000000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
        00FF0000FF0000FF0000}
      OnClick = btn1Click
    end
    object Label1: TLabel
      Left = 20
      Top = 96
      Width = 99
      Height = 13
      Caption = #1057#1091#1084#1084#1072' '#1074#1099#1076#1077#1083#1077#1085#1085#1099#1093
    end
    object Label2: TLabel
      Left = 88
      Top = 16
      Width = 31
      Height = 13
      Caption = 'Label2'
    end
    object Label3: TLabel
      Left = 160
      Top = 16
      Width = 31
      Height = 13
      Caption = 'Label3'
    end
    object Label4: TLabel
      Left = 263
      Top = 7
      Width = 68
      Height = 13
      Caption = #1056#1072#1089#1082#1083#1102#1095#1077#1085#1080#1077
    end
    object Label5: TLabel
      Left = 455
      Top = 8
      Width = 44
      Height = 13
      Caption = #1054#1073#1074#1072#1088#1082#1072
    end
    object Label6: TLabel
      Left = 656
      Top = 29
      Width = 200
      Height = 13
      Caption = #1042#1086#1079#1076#1091#1096#1085#1099#1077': '#1057#1073#1086#1088#1082#1072'1+'#1057#1073#1051#1086#1087'+'#1054#1073#1074#1072#1088#1082#1072
    end
    object Label7: TLabel
      Left = 656
      Top = 48
      Width = 229
      Height = 13
      Caption = #1042#1086#1079#1076#1091#1096#1085#1099#1077': '#1057#1073#1086#1088#1082#1072'2('#1057#1074#1072#1088#1082#1072')+'#1056#1072#1089#1082#1083#1102#1095#1077#1085#1080#1077
    end
    object Button1: TButton
      Left = 128
      Top = 48
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 0
      Visible = False
      OnClick = Button1Click
    end
    object Edit1: TEdit
      Left = 128
      Top = 94
      Width = 121
      Height = 21
      TabOrder = 1
      Text = '0'
    end
    object Memo1: TMemo
      Left = 263
      Top = 26
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 2
    end
    object Memo2: TMemo
      Left = 454
      Top = 26
      Width = 185
      Height = 89
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 3
    end
  end
  object SG: TAdvStringGrid
    Left = 0
    Top = 121
    Width = 1126
    Height = 581
    Cursor = crDefault
    Align = alClient
    ColCount = 20
    DrawingStyle = gdsClassic
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    GridLineWidth = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ParentFont = False
    PopupMenu = pm1
    ScrollBars = ssBoth
    TabOrder = 1
    OnMouseUp = SGMouseUp
    ActiveRowShow = True
    ActiveRowColor = clSkyBlue
    HoverRowCells = [hcNormal, hcSelected]
    OnGetCellColor = SGGetCellColor
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
    RowIndicator.Data = {
      C6000000424DC60000000000000076000000280000000A0000000A0000000100
      0400000000005000000000000000000000001000000000000000000000000000
      8000008000000080800080000000800080008080000080808000C0C0C0000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00CCCCCCCCCC00
      0000CCCCCCCCCC000000CCCCCCCCCC000000CCCCCCCCCC000000CCCCCCCCCC00
      0000CCCCCCCCCC000000CCCCCCCCCC000000CCCCCCCCCC000000CCCCCCCCCC00
      0000CCCCCCCCCC000000}
    ScrollWidth = 12
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
      79
      82
      350
      40
      66
      109
      55
      86
      52
      76
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
  object pm1: TPopupMenu
    Left = 456
    Top = 209
    object Copy1: TMenuItem
      Caption = 'Copy'
      ShortCut = 16451
      OnClick = Copy1Click
    end
  end
end

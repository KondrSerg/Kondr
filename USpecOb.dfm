object FSpecOb: TFSpecOb
  Left = 258
  Top = 235
  Caption = 'FSpecOb'
  ClientHeight = 560
  ClientWidth = 970
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object SpecGrid: TAdvStringGrid
    Left = 0
    Top = 85
    Width = 970
    Height = 475
    Cursor = crDefault
    Align = alClient
    ColCount = 20
    DrawingStyle = gdsClassic
    FixedCols = 18
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    GridLineWidth = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ParentFont = False
    PopupMenu = pm1
    ScrollBars = ssBoth
    TabOrder = 1
    OnDblClick = SpecGridDblClick
    OnDrawCell = SpecGridDrawCell
    ActiveRowShow = True
    ActiveRowColor = clSilver
    GridLineColor = clBlack
    GridFixedLineColor = clBlack
    HoverRowCells = [hcNormal, hcSelected]
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    AutoNumAlign = True
    Bands.Active = True
    Bands.PrimaryColor = clMenu
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
    ControlLook.ProgressMarginX = 1
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
    FixedColWidth = 20
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -12
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = [fsBold]
    FloatFormat = '%.2f'
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    MouseActions.SizeFixedCol = True
    Multilinecells = True
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
    SelectionColor = clLime
    SelectionColorTo = clInactiveCaptionText
    SelectionColorMixer = True
    SortSettings.DefaultFormat = ssAutomatic
    Version = '8.1.3.0'
    ColWidths = (
      20
      20
      20
      20
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
      64)
    RowHeights = (
      22
      22)
    object Memo3: TMemo
      Left = 1004
      Top = 56
      Width = 109
      Height = 89
      Lines.Strings = (
        'Memo6')
      ScrollBars = ssBoth
      TabOrder = 2
      Visible = False
    end
    object Memo1: TMemo
      Left = 288
      Top = 80
      Width = 469
      Height = 89
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 3
      Visible = False
    end
    object Memo2: TMemo
      Left = 288
      Top = 232
      Width = 469
      Height = 89
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 4
      Visible = False
    end
  end
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 970
    Height = 85
    Align = alTop
    Caption = 'pnl1'
    TabOrder = 0
    object Label1: TLabel
      Left = 636
      Top = 32
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object Memo4: TMemo
      Left = 1
      Top = 1
      Width = 968
      Height = 83
      Align = alClient
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssBoth
      TabOrder = 0
      Visible = False
    end
  end
  object pm1: TPopupMenu
    Left = 288
    Top = 213
    object NNCopy1: TMenuItem
      Caption = 'Copy'
      OnClick = NNCopy1Click
    end
  end
  object actlst1: TActionList
    Left = 720
    Top = 44
    object act1: TAction
      Caption = 'act1'
      ShortCut = 16454
      OnExecute = act1Execute
    end
  end
end

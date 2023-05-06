object FSborNCH: TFSborNCH
  Left = 0
  Top = 0
  Caption = #1057#1073#1086#1088#1082#1072
  ClientHeight = 439
  ClientWidth = 889
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pgc1: TPageControl
    Left = 0
    Top = 0
    Width = 889
    Height = 439
    ActivePage = ts1
    Align = alClient
    TabOrder = 0
    object ts1: TTabSheet
      Caption = #1058#1080#1087#1099' '#1050#1083#1072#1087#1072#1085#1086#1074
      object pnl2: TPanel
        Left = 0
        Top = 0
        Width = 881
        Height = 145
        Align = alTop
        TabOrder = 0
        object Label1: TLabel
          Left = 152
          Top = 96
          Width = 31
          Height = 13
          Caption = 'Label1'
        end
        object Label2: TLabel
          Left = 16
          Top = 8
          Width = 43
          Height = 13
          Caption = #1048#1079#1076#1077#1083#1080#1077
        end
        object Label3: TLabel
          Left = 16
          Top = 74
          Width = 64
          Height = 13
          Caption = #1058#1080#1087' '#1050#1083#1072#1087#1072#1085#1072
        end
        object Label4: TLabel
          Left = 648
          Top = 8
          Width = 61
          Height = 13
          Caption = #1056#1072#1079#1088#1077#1096#1077#1085#1080#1103
        end
        object Label5: TLabel
          Left = 16
          Top = 120
          Width = 6
          Height = 13
          Caption = '0'
        end
        object btn4: TButton
          Left = 455
          Top = 54
          Width = 75
          Height = 25
          Caption = #1054#1095#1080#1089#1090#1080#1090#1100
          TabOrder = 0
          OnClick = btn4Click
        end
        object btn5: TButton
          Left = 552
          Top = 23
          Width = 75
          Height = 25
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 1
          OnClick = btn5Click
        end
        object btn6: TButton
          Left = 552
          Top = 54
          Width = 75
          Height = 25
          Caption = #1042#1099#1093#1086#1076
          TabOrder = 2
          OnClick = btn6Click
        end
        object btndt1: TButtonedEdit
          Left = 16
          Top = 27
          Width = 433
          Height = 21
          TabOrder = 3
          OnKeyDown = btndt1KeyDown
        end
        object btn7: TButton
          Left = 456
          Top = 24
          Width = 75
          Height = 25
          Caption = #1054#1073#1088#1072#1073#1086#1090#1072#1090#1100
          TabOrder = 4
          OnClick = btn7Click
        end
        object ComboBox1: TComboBox
          Left = 648
          Top = 23
          Width = 145
          Height = 21
          TabOrder = 5
          OnSelect = ComboBox1Select
          Items.Strings = (
            #1042#1057#1045
            #1042#1089#1077' '#1079#1072#1073#1083#1086#1082#1080#1088#1086#1074#1072#1085#1085#1099#1077
            '')
        end
        object ComboBox2: TComboBox
          Left = 16
          Top = 93
          Width = 161
          Height = 21
          TabOrder = 6
          OnSelect = ComboBox2Select
        end
        object btn8: TButton
          Left = 648
          Top = 56
          Width = 75
          Height = 25
          Caption = #1042#1099#1073#1086#1088#1082#1072
          TabOrder = 7
          Visible = False
          OnClick = btn8Click
        end
        object btn1: TButton
          Left = 456
          Top = 85
          Width = 75
          Height = 25
          Caption = 'btn1'
          TabOrder = 8
          Visible = False
          OnClick = btn1Click
        end
      end
      object SG1: TAdvStringGrid
        Left = 0
        Top = 145
        Width = 881
        Height = 266
        Cursor = crDefault
        Align = alClient
        ColCount = 3
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
          565
          100)
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
        object Memo2: TMemo
          Left = 660
          Top = 80
          Width = 185
          Height = 89
          Lines.Strings = (
            'mmo1')
          TabOrder = 2
          Visible = False
        end
      end
    end
    object ts2: TTabSheet
      Caption = #1047#1072#1087#1088#1077#1090' '#1087#1088#1086#1080#1079#1074#1086#1076#1089#1090#1074#1072
      ImageIndex = 1
    end
  end
  object pm1: TPopupMenu
    Left = 324
    Top = 193
    object N1: TMenuItem
      Caption = #1053#1086#1074#1099#1081' '#1090#1080#1087
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1047#1072#1087#1088#1077#1090' '#1087#1088#1086#1080#1079#1074#1086#1076#1089#1090#1074#1072
      OnClick = N2Click
    end
    object N3: TMenuItem
      Caption = #1057#1085#1103#1090#1100' '#1047#1072#1087#1088#1077#1090
      OnClick = N3Click
    end
    object N4: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N4Click
    end
    object N5: TMenuItem
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      ShortCut = 16451
      OnClick = N5Click
    end
  end
end

object DialogMain: TDialogMain
  Left = 192
  Top = 103
  BorderStyle = bsDialog
  Caption = 'Комплексная обработка ошибок'
  ClientHeight = 248
  ClientWidth = 343
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object btnErrorRaise: TButton
    Left = 134
    Top = 220
    Width = 75
    Height = 25
    Caption = 'Ошибка'
    TabOrder = 0
    OnClick = OnRaiseError
  end
  object rgrpErrorType: TRadioGroup
    Left = 0
    Top = 0
    Width = 343
    Height = 213
    Align = alTop
    Caption = 'Типы генерируемых ошибок'
    ItemIndex = 0
    Items.Strings = (
      'Неверный указатель'
      'Необрабатываемое открытие файла (надежная функция)'
      'Обрабатываемое открытие файла (ненадежная функция)'
      'Ошибка WIN32 API'
      'Традиционный стиль'
      'Стиль комплексной обработки ошибок'
      'Проверка аргументов и целостности'
      'Ошибочная функция')
    TabOrder = 1
  end
end

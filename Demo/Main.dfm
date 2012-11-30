object Form1: TForm1
  Left = 330
  Top = 106
  Caption = 'Form1'
  ClientHeight = 339
  ClientWidth = 528
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 168
    Top = 40
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 0
    OnClick = Button1Click
  end
  object RS: TRSPrinter
    PageSize = pzLegal
    PageLength = 0
    FastPrinter = Epson_FX
    FastFont = []
    FastPort = 'LPT1'
    SaveConfToRegistry = False
    Preview = ppYes
    Left = 88
    Top = 136
  end
end

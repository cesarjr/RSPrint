object FrmMain: TFrmMain
  Left = 330
  Top = 106
  Caption = 'RSPrint Demo App'
  ClientHeight = 166
  ClientWidth = 197
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object lblMode: TLabel
    Left = 8
    Top = 17
    Width = 27
    Height = 13
    Caption = 'Mode'
  end
  object btnPreview: TButton
    Left = 8
    Top = 96
    Width = 177
    Height = 25
    Caption = 'Preview'
    TabOrder = 0
    OnClick = btnPreviewClick
  end
  object btnPrint: TButton
    Left = 8
    Top = 127
    Width = 177
    Height = 25
    Caption = 'Print'
    TabOrder = 1
    OnClick = btnPrintClick
  end
  object cmbMode: TComboBox
    Left = 8
    Top = 32
    Width = 177
    Height = 21
    Style = csDropDownList
    ItemIndex = 1
    TabOrder = 2
    Text = 'rmWindows'
    Items.Strings = (
      'rmFast'
      'rmWindows'
      'rmDefault')
  end
  object RSPrint: TRSPrinter
    PageSize = pzLegal
    PageLength = 0
    FastPrinter = Epson_FX
    FastFont = []
    Preview = ppYes
    Left = 96
    Top = 64
  end
end

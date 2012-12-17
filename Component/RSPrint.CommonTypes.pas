unit RSPrint.CommonTypes;

interface

uses
  Classes, Graphics;

type

  TPageSize = (pzDefault, pzContinuous, pzLetter, pzLegal, pzA4, pzLetterSmall, pzTabloid, pzLedger, pzStatement,
    pzExecutive, pzA3, pzA4small, pzA5, pzB4, pzB5, pzFolio, pzQuarto, pz10x14, pz11x17, pzNote, pzEnv9, pzEnv10,
    pzEnv11, pzEnv12, pzEnv14, pzEnvDl, pzEnvC5, pzEnvC3, pzEnvC4, pzEnvC6, pzEnvC65, pzEnvB4, pzEnvB5, pzEnvB6,
    pzEnvItaly, pzEnvMonarch, pzEnvPersonal, pzFanfoldUS, pzFanfoldStd, pzFanfoldLgl);

  TFontType = (Bold, Italic, Underline, Compress, DobleWide);

  TFastFont = Set of TFontType;

  TPrinterModel = (Cannon_F60, Cannon_Laser, Epson_FX, Epson_LX, Epson_Stylus, HP_Deskjet, HP_Laserjet, HP_Thinkjet,
    IBM_Color_Jet, IBM_PC_Graphics, IBM_Proprinter, NEC_3500, NEC_Pinwriter, Mp20Mi);

  TPrinterMode = (pmFast, pmWindows);

  TLineType = (ltSingle, ltDouble);

  TTbAlign = (alLeft, alCenter, alRight);

  TInitialZoom = (zWidth, zHeight);

  TRSPrinterPreview = (ppYes, ppNo);

  TTbPreviewType = (pYes, pNo, pDefault);

  TRSPrinterMode = (rmFast, rmWindows, rmDefault);

  TControlCodes = record
    Normal: string;
    Bold: string;
    Wide: string;
    Italic: string;
    UnderlineON: string;
    UnderlineOFF: string;
    CondensedON: string;
    CondensedOFF: string;
    Setup: string;
    Reset: string;
    SelLength: string;
  end;

  PPage = ^TPage;
  TPage = record
    Writed: TList;
    VerticalLines: TList;
    HorizLines: TList;
    PrintedLines: Byte;
    Graphics: TList;
  end;

  PGraphic = ^TGraphic;
  TGraphic = record
    Col: Single;
    Line: Single;
    Picture: TPicture;
  end;

  PWrite = ^TWrite;
  TWrite = record
    Col: Byte;
    Line: Byte;
    Font: TFastFont;
    Text: string;
  end;

  PHorizLine = ^THorizLine;
  THorizLine = record
    Col1: Byte;
    Col2: Byte;
    Line: Byte;
    Kind: TLineType;
  end;

  PVertLine = ^TVertLine;
  TVertLine = record
    Col: Byte;
    Line1: Byte;
    Line2: Byte;
    Kind: TLineType;
  end;

  PPrintJob = ^TPrintJob;
  TPrintJob = record
    Name: string;
    PageSize: TPageSize;
    PageContinuousJump: Byte;
    PageLength: Byte;
    LasPaginas: TList;
    Copias: Integer;
    DefaultFont: TFastFont;
    Lineas: Integer;
    Transliterate: Boolean;
    ControlCodes: TControlCodes;
  end;

  TPrinterStatus = record
    PrintingCanceled: Boolean;
    PrintingCancelAll: Boolean;
    PrintingPaused: Boolean;
    PrintingJobName: string;
    CurrentlyPrinting: boolean;

    procedure CancelPrinting;
    procedure CancelAllPrinting;
    procedure PausePrinting;
    procedure RestorePrinting;

    procedure StartPrinting;
  end;

implementation

procedure TPrinterStatus.CancelAllPrinting;
begin
  PrintingCancelAll := True;
end;

procedure TPrinterStatus.CancelPrinting;
begin
  PrintingCanceled := True;
end;

procedure TPrinterStatus.PausePrinting;
begin
  PrintingPaused := True;
end;

procedure TPrinterStatus.StartPrinting;
begin
  PrintingCanceled := False;
  PrintingPaused:= False;
  PrintingCancelAll := False;
  CurrentlyPrinting := True;
end;

procedure TPrinterStatus.RestorePrinting;
begin
  PrintingPaused := False;
end;

end.

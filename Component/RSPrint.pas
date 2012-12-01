unit RSPrint;

interface

//{$R RSP.dres}

uses
  Windows, Messages, SysUtils, CommDlg, Classes, Graphics, Controls, ExtCtrls, StdCtrls, Consts, ShellAPI, menus,
  Printers, ComCtrls, Forms, Dialogs;

const
  RSPrintVersion = 'Versão 2.0';
  WM_TASKICON = WM_USER + 10;

type
  TRSPrinter = class;

  TPrinterModel = (Cannon_F60, Cannon_Laser, Epson_FX, Epson_LX, Epson_Stylus, HP_Deskjet, HP_Laserjet, HP_Thinkjet,
    IBM_Color_Jet, IBM_PC_Graphics, IBM_Proprinter, NEC_3500, NEC_Pinwriter,Mp20Mi);

  TPrinterMode = (pmFast, pmWindows);
  TLineType = (ltSingle, ltDouble);

  TFontType = (Bold, Italic, Underline, Compress, DobleWide);
  TFastFont = Set of TFontType;

  TTbAlign = (alLeft, alCenter, alRight);

  PPage = ^TPage;
  TPage = record
    Writed: TList;
    VerticalLines: TList;
    HorizLines: TList;
    PrintedLines: byte;
    Graphics: TList;
  end;

  PGraphic = ^TGraphic;
  TGraphic = record
    Col: single;
    Line: single;
    Picture: TPicture;
  end;

  PWrite = ^TWrite;
  TWrite = record
    Col: byte;
    Line: byte;
    FastFont: TFastFont;
    Text: string;
  end;

  PHorizLine = ^THorizLine;
  THorizLine = record
    Col1: byte;
    Col2: byte;
    Line: byte;
    Kind: TLineType;
  end;

  PVertLine = ^TVertLine;
  TVertLine = record
    Col: byte;
    Line1: byte;
    Line2: byte;
    Kind: TLineType;
  end;

  EPrinterError = class(Exception);

  TInitialZoom = (zWidth, zHeight);

  TPageSize = (pzDefault, pzContinuous, pzLetter, pzLegal, pzA4, pzLetterSmall, pzTabloid, pzLedger, pzStatement,
    pzExecutive, pzA3, pzA4small, pzA5, pzB4, pzB5, pzFolio, pzQuarto, pz10x14, pz11x17, pzNote, pzEnv9, pzEnv10,
    pzEnv11, pzEnv12, pzEnv14, pzEnvDl, pzEnvC5, pzEnvC3, pzEnvC4, pzEnvC6, pzEnvC65, pzEnvB4, pzEnvB5, pzEnvB6,
    pzEnvItaly, pzEnvMonarch, pzEnvPersonal, pzFanfoldUS, pzFanfoldStd, pzFanfoldLgl);

  PTbPrn = ^TTbPrn;
  TTbPrn = record
    Name: string;
    FastModel: string;
    WinModel: string;
    FastPort: string;
    WinTopMargin: Integer;
    WinBottomMargin: Integer;
    DefaultMode: TPrinterMode;
  end;

  TTbSaveMode = (smNoSave, smFile, smRegistry);

  PPrintJob = ^TPrintJob;
  TPrintJob = record
    Name: string;
    PageSize: TPageSize;
    PageLength: byte;
    LasPaginas: TList;
    FCopias: integer;
    FFuente: TFastFont;
    FPort: string;
    FLineas: integer;
    FTransliterate: boolean;
    PRNNormal: string;
    PRNBold: string;
    PRNWide: string;
    PRNItalics: string;
    PRNULineON: string;
    PRNULineOFF: string;
    PRNCompON: string;
    PRNCompOFF: string;
    PRNSetup: string;
    PRNReset: string;
    PRNSelLength: string;
  end;

  TPrintThread = class(TThread)
  private
    Job: PPrintJob;
    procedure SetJobName;

  protected
    procedure Execute; override;
    function PrintFast(Number : integer) : boolean;

  public
    RSPrinter: TRSPrinter;
    Jobs: TThreadList;
  end;

  TRSPrintTray = class(TComponent)
  private
    RSPrinter: TRSPrinter;
    FPopupMenuL: TPopupMenu;
    FIcon: TIcon;
    FTip: string;
    FWnd: HWnd;
    procedure DoPopup(i: integer);
    procedure TBChange;
    procedure SetTip(s: string);
    procedure PauseClick(Sender: TObject);
    procedure CancelClick(Sender: TObject);
    procedure CancelAllClick(Sender: TObject);

  protected
    procedure WndProc(var Msg: TMessage);
    procedure DoOnLeftClick; virtual;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  published
    property Tip: string read FTip write SetTip;
  end;

  TRSPrinterPreview = (ppYes,ppNo);

  TRSPrinter = class(TComponent)
  private
    FOwner: TForm;
    fOnPrinterError: TNotifyEvent;
    FDatosEmpresa: String;      // INFORMACION QUE SE IMPRIME EN TODOS LOS REPORTES
    FModelo: TPrinterModel;           // MODELO DE IMPRESORA
    FFastPuerto: string;         // PUERTO DE LA IMPRESORA
    FLineas: byte;              // CANTIDAD DE LINEAS POR PAGINA
    FColumnas: byte;            // CANTIDAD DE COLUMNAS EN LA PAGINA
    FFuente: TFastFont;         // TIPO DE LETRA
    FModo: TPrinterMode;        // MODO DE IMPRESION (NORMAL/MEJORADO)
    LasPaginas: TList;             // Almacenamiento de las páginas
    FCopias: integer;
    FWinSupMargin: integer;     // MARGEN SUPERIOR PARA EL MODO WINDOWS
    FWinBotMargin: integer;     // MARGEN INFERIOR PARA EL MODO WINDOWS
    PRNNormal: string;
    PRNBold: string;
    PRNWide: string;
    PRNItalics: string;
    PRNULineON: string;
    PRNULineOFF: string;
    PRNCompON: string;
    PRNCompOFF: string;
    PRNSetup: string;
    PRNReset: string;
    PRNSelLength: string;
    PaginaActual: Integer;
    FPreview: TRSPrinterPreview;
    FZoom: TInitialZoom;
    FTitulo: string;
    PreviewForm: TForm;
    FWinPrinter: string;        // NOMBRE DE LA IMPRESORA EN WINDOWS
    FWinPort: string;
    FTransliterate: Boolean;
    FWinFont: string;
    OldCloseQuery: procedure (Sender: TObject; var CanClose: Boolean) of object;
    TrayIcon: TRSPrintTray;
    PrintThread: TPrintThread;
    FPageSize: TPageSize;
    FPageLength: byte;
    FReportName: String;
    FSaveConfToRegistry: boolean;
    FContinuousJump: byte;
    Cancelado: boolean;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    function GetModelRealName(Model : TPrinterModel) : String;
    procedure SetModel(Name: TPrinterModel);
    procedure SetFastPuerto(Puerto: string);
    function GetPaginas: integer;
    procedure Clear;
    procedure PrintThreadTerminate(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);

  protected
    procedure Loaded; override;

  public
    GridFrm: TForm;
    Fonts: TStringList;
    PageNo: Integer;
    PageWidth: Integer;         // ANCHO DE PAGINA EN PIXELS
    PageHeight: Integer;  // ALTO DE PAGINA EN PIXELS
    PageWidthP: Double;        // ANCHO DE PAGINA EN PULGADAS
    PageHeightP: Double;  // ALTO DE PAGINA EN PULGADAS
    PageOrientation: TPrinterOrientation; // ORIENTACION DE LA PAGINA
    ReGenerate: Procedure of object;
    CurrentlyPrinting : boolean;
    procedure PreviewReal;                            // MUESTRA LA IMPRESION EN PANTALLA
    procedure CancelAllPrinting;
    procedure PausePrinting;
    procedure RestorePrinting;
    procedure CancelPrinting;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure BeginDoc;  // PARA COMENZAR UNA NUEVA IMPRESION
    procedure EndDoc; // ELIMINA LA INFORMACION DE LAS PAGINAS AL FINALIZAR
    procedure Write(Line, Col: byte; Text: string); // ESCRIBE UN TEXTO EN LA PAGINA
    procedure WriteFont(Line, Col: byte; Text: string; Font: TFastFont); // ESCRIBE UN TEXTO EN LA PAGINA
    procedure DrawMetafile(Col, Line: single; Picture: TMetafile); // Hace un dibujo
    procedure Box(Line1,Col1,Line2,Col2 : byte; Kind : TLineType);                // ESCRIBE UN RECTANGULO
    procedure LineH(Line,Col1,Col2 : byte; Kind : TLineType);           // ESCRIBE UNA LINEA HORIZONTAL
    procedure LineV(Line1,Line2,Col : byte; Kind : TLineType);             // ESCRIBE UNA LINEA VERTICAL
    procedure Print;                           // MANDA LA IMPRESION A LA IMPRESORA
    procedure NewPage;                        // CREA UNA NUEVA PAGINA
    procedure SetModelName(Name : String);
    procedure GetModels(Models : TStrings);
    function GetModelName : String;
    procedure BuildPage(Number : integer; Page : TMetaFile;ToPrint:boolean;Mode : TPrinterMode;ShowDialog : boolean);
    property Lines : byte read FLineas write FLineas default 66; // CANTIDAD DE LINEAS POR PAGINA
    function PrintPage(Number:integer) : boolean;// MANDA UNA PAGINA A LA IMPRESORA
    procedure PrintAll;                           // MANDA LA IMPRESION A LA IMPRESORA
    function GetPrintingWidth : integer;
    function GetPrintingHeight : integer;
    property WinPrinter : string read fWinPrinter write fWinPrinter;
    property WinPort : string read fWinPort write fWinPort;

  published
    property PageContinuousJump : byte read FContinuousJump write FContinuousJump default 5;
    property PageSize : TPageSize read FPageSize write FPageSize;
    property PageLength : byte read FPageLength write FPageLength;
    property FastPrinter : TPrinterModel read FModelo Write SetModel; // MODELO DE LA IMPRESORA
    property FastFont : TFastFont read FFuente write FFuente;
    property FastPort : string read FFastPuerto Write SetFastPuerto; // PUERTO DE LA IMPRESORA
    property Mode : TPrinterMode read FModo write FModo default pmFast; // MODO DE IMPRESION
    property Columnas : byte read FColumnas default 80; // CANTIDAD DE COLUMNAS EN LA PAGINA
    property Paginas : Integer read GetPaginas default 0;             // CANTIDAD DE PAGINAS
    property CompanyData : string read FDatosEmpresa write FDatosEmpresa; // SE IMPRIME EN TODOS LOS REPORTES
    property ReportName : string read FReportName  write FReportName;
    property SaveConfToRegistry : boolean read FSaveConfToRegistry write FSaveConfToRegistry;
    property Zoom : TInitialZoom read FZoom write FZoom default zWidth;
    property Preview : TRSPrinterPreview read FPreview write FPreview;
    property Title : string read FTitulo write FTitulo;
    property Copies : integer read fCopias write fCopias default 1;
    property OnPrinterError : TNotifyEvent read fOnPrinterError write fOnPrinterError;
    property PrintingWidth : integer read GetPrintingWidth;
    property PrintingHeight : integer read GetPrintingHeight;
    property Transliterate : boolean read FTransliterate write FTransliterate default True;
    property WinMarginTop : integer read FWinSupMargin write FWinSupMargin default 0;
    property WinMarginBottom : integer read FWinBotMargin write FWinBotMargin default 0;
  end;

  TTbPreviewType = (pYes, pNo, pDefault);
  TRSPrinterMode = (rmFast, rmWindows, rmDefault);

  TTbProcedure = procedure of object;

  TTbGetTextEvent = procedure(var Text: string) of object;

procedure Register;

implementation

uses
  RSPrintV, ComObj, splprint;

var
  PrintingCanceled: boolean;
  PrintingCancelAll: boolean;
  PrintingPaused: boolean;
  PrintingJobName: string;

procedure Register;
begin
  RegisterComponents('RS Componentes', [TRSPrinter]);
end;

function Min(Val1, Val2: Integer):integer;
begin
  if Val1<Val2 then
    Min := Val1
  else
    Min := Val2;
end;

function Spaces(N : integer) : string;
var
  S: string;
begin
  S := '';
  while Length(S) < N do
    S := S + ' ';
  Spaces := S;
end;

(* IMPRESSORA *)

constructor TRSPrinter.Create(AOwner: TComponent);
var
  ADevice: array [0..255] of Char;
  ADriver: array [0..255] of Char;
  APort: array [0..255] of Char;
  DeviceMode: THandle;
begin
  inherited Create(AOwner);

  if (AOwner is TForm) and not (CsDesigning in ComponentState) then
  begin
    FOwner := TForm(AOwner);
    OldCloseQuery := TForm(AOwner).OnCloseQuery;
    TForm(AOwner).OnCloseQuery := FormCloseQuery;
  end;

  PrintThread := nil;
  SetModel(EPSON_FX);
  SetFastPuerto('LPT1');
  FColumnas := 80;
  PaginaActual := 0;
  PageNo:= 1;
  fCopias := 1;
  FZoom := zWidth;
  FFuente := [];
  FPageSize := pzLegal;
  FContinuousJump := 5;
  LasPaginas := TList.Create;

  if (Printer <> nil) and (Printer.Printers.Count > 0) then
  begin
    PageOrientation := Printer.Orientation;
    PageWidth := GetDeviceCaps(Printer.Handle, PHYSICALWIDTH);
    PageHeight := GetDeviceCaps(Printer.Handle, PHYSICALHEIGHT);
    PageWidthP := PageWidth/GetDeviceCaps(Printer.Handle, LOGPIXELSX);
    PageHeightP := PageHeight/GetDeviceCaps(Printer.Handle, LOGPIXELSY);
    Lines := Trunc(PageHeightP*6)-4;
    Printer.GetPrinter(ADevice, ADriver, APort, DeviceMode);
    WinPort := APort;
    WinPrinter := ADevice;
  end;

  ReGenerate := nil;
  FTransliterate := True;
  FWinFont := 'courier new';
end;

destructor TRSPrinter.Destroy;
begin
  Clear;

  if not (CsDesigning in ComponentState) then
    FOwner.OnCloseQuery := OldCloseQuery;

  inherited Destroy;
end;

procedure TRSPrinter.Loaded;
begin
  FLineas := FLineas - (WinMarginTop div 12) - (WinMarginBottom div 12);
end;

procedure TRSPrinter.Clear;
var
  Pag: PPage;
  HLin: PHorizLine;
  VLin: PVertLine;
  Esc: PWrite;
  Graf: PGraphic;
begin
  while LasPaginas.Count > 0 do
  begin
    Application.ProcessMessages;
    Pag := LasPaginas.Items[0];
    while Pag.Writed.Count > 0 do
    begin
      Esc := Pag.Writed.Items[0];
      Dispose(Esc);
      Pag.Writed.Delete(0);
      Pag.Writed.Pack;
    end;
    Pag.Writed.Free;

    Application.ProcessMessages;
    while Pag.HorizLines.Count > 0 do
    begin
      HLin := Pag.HorizLines.Items[0];
      Dispose(HLin);
      Pag.HorizLines.Delete(0);
      Pag.HorizLines.Pack;
    end;

    Application.ProcessMessages;
    while Pag.VerticalLines.Count > 0 do
    begin
      VLin := Pag.VerticalLines.Items[0];
      Dispose(VLin);
      Pag.VerticalLines.Delete(0);
      Pag.VerticalLines.Pack;
    end;
    Pag.VerticalLines.Free;

    Application.ProcessMessages;
    while Pag.Graphics.Count > 0 do
    begin
      Graf := Pag.Graphics[0];
      Graf^.Picture.Free;
      Dispose(Graf);
      Pag.Graphics.Delete(0);
      Pag.Graphics.Pack;
    end;
    Pag.Graphics.Free;

    Dispose(Pag);
    LasPaginas.Delete(0);
    LasPaginas.Pack;
  end;
end;

procedure TRSPrinter.SetModel(Name : TPrinterModel);
var
  Nombre: TPrinterModel;
begin
  Nombre := Name;
  FModelo := Nombre;

  case Nombre of
    Cannon_F60 :
    begin
      PRNNormal := '27 70 27 54 00';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '27 54 01';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    Cannon_Laser :
    begin
      PRNNormal := '27 38';
      PRNBold := '27 79';
      PRNWide := '27 87 01';
      PRNItalics := '';
      PRNULineON := '27 69';
      PRNULineOFF := '27 82';
      PRNCompON := '27 31 09';
      PRNCompOFF := '27 31 13';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    HP_Deskjet :
    begin
      PRNNormal := '27 40 115 48 66 27 40 115 48 83';
      PRNBold := '27 40 115 51 66';
      PRNWide := '27 87 01';
      PRNItalics := '27 40 115 49 83';
      PRNULineON := '27 38 100 48 68';
      PRNULineOFF := '27 38 100 64';
      PRNCompON := '27 40 115 49 54 46 54 72';
      PRNCompOFF := '27 40 115 49 48 72';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    HP_Laserjet :
    begin
      PRNNormal := '27 40 115 48 66 27 40 115 48 83';
      PRNBold := '27 40 115 53 66';
      PRNWide := '27 87 01';
      PRNItalics := '27 40 115 49 83';
      PRNULineON := '27 38 100 68';
      PRNULineOFF := '27 38 100 64';
      PRNCompON := '27 40 115 49 54 46 54 72';
      PRNCompOFF := '27 40 115 49 48 72';
      PRNSetup := '';
      PRNReset := '12';
      PRNSelLength := '27 67';
    end;

    HP_Thinkjet :
    begin
      PRNNormal := '27 70';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '';
      PRNULineON := '27 45 49';
      PRNULineOFF := '27 45 48';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    IBM_Color_Jet :
    begin
      PRNNormal := '27 72';
      PRNBold := '27 71';
      PRNWide := '27 87 01';
      PRNItalics := '';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    IBM_PC_Graphics :
    begin
      PRNNormal := '27 70 27 55';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '27 54';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    IBM_Proprinter :
    begin
      PRNNormal := '27 70';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    NEC_3500 :
    begin
      PRNNormal := '27 72';
      PRNBold := '27 71';
      PRNWide := '27 87 01';
      PRNItalics := '';
      PRNULineON := '27 45';
      PRNULineOFF := '27 39';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    NEC_Pinwriter:
    begin
      PRNNormal := '27 70 27 53';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '27 52';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '';
      PRNSelLength := '27 67';
    end;

    Epson_Stylus:
    begin
      PRNNormal := '27 70 27 53 27 45 00';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '27 52';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '27 70 15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '27 64';
      PRNSelLength := '27 67';
    end;

    Mp20Mi:
    begin
      PRNNormal := '27 77';
      PRNBold := '27 69';
      PRNWide := '27 87 01';
      PRNItalics := '27 52';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '27 15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '27 64';
      PRNSelLength := '27 67';
    end;
    else
    begin
      PRNNormal := '27 70 27 53';
      PRNBold := '27 69';
      PRNWide := '27 14';
      PRNItalics := '27 52';
      PRNULineON := '27 45 01';
      PRNULineOFF := '27 45 00';
      PRNCompON := '15';
      PRNCompOFF := '18';
      PRNSetup := '';
      PRNReset := '27 64';
      PRNSelLength := '27 67';
    end;
  end;
end;

procedure TRSPrinter.SetFastPuerto(Puerto : string);
begin
  if (Puerto <> fFastPuerto) then
    FFastPuerto := Puerto;
end;

function TRSPrinter.GetPaginas : Integer;
begin
  GetPaginas := LasPaginas.Count;
end;

procedure TRSPrinter.BeginDoc;  // PARA COMENZAR UNA NUEVA IMPRESION
begin
  Clear;
  NewPage;
end;

procedure TRSPrinter.EndDoc; // ELIMINA LA INFORMACION DE LAS PAGINAS AL FINALIZAR
begin
  Clear;
end;

procedure TRSPrinter.WriteFont(Line,Col : byte; Text:string;Font : TFastFont); // ESCRIBE UN TEXTO EN LA PAGINA
var
  Txt: PWrite;
  Pag: PPage;
  P: integer;
begin
  if LasPaginas.Count > 0 then
    begin
      Pag := LasPaginas.Items[PaginaActual];
      P := 0;

      if Pag.Writed.Count > 0 then
      begin
        Txt := Pag.Writed.Items[0];
        while (P<Pag.Writed.Count) and
              ((Txt^.Line < Line) or
              ((Txt^.Line = Line)and(Txt^.Col<=Col))) do
        begin
          Inc(P);
          if P<Pag.Writed.Count then
            Txt := Pag.Writed.Items[P];
        end;
      end;

      New(Txt);
      Pag.Writed.Insert(P,Txt);
      Txt^.Col := Col;
      Txt^.Line := Line;
      Txt^.Text := Text;
      Txt^.FastFont := Font;

      if Line > Pag.PrintedLines then
        Pag.PrintedLines := Line;
    end;
end;

procedure TRSPrinter.Write(Line,Col : byte; Text:string); // ESCRIBE UN TEXTO EN LA PAGINA
var
  Txt: PWrite;
  Pag: PPage;
  P: integer;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];
    P := 0;

    if Pag.Writed.Count>0 then
    begin
      Txt := Pag.Writed.Items[0];
      while (P<Pag.Writed.Count) and
            ((Txt^.Line < Line) or
            ((Txt^.Line = Line)and(Txt^.Col<=Col))) do
      begin
        Inc(P);
        if P<Pag.Writed.Count then
          Txt := Pag.Writed.Items[P];
      end;
    end;

    New(Txt);
    Pag.Writed.Insert(P,Txt);
    Txt^.Col := Col;
    Txt^.Line := Line;
    Txt^.Text := Text;
    Txt^.FastFont := FFuente;
    if Line > Pag.PrintedLines then
      Pag.PrintedLines := Line;
  end;
end;

procedure TRSPrinter.DrawMetafile(Col,Line : single; Picture : TMetafile); // Hace un dibujo
var
  Pag: PPage;
  Dib: TPicture;
  Graf: PGraphic;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];
    Dib := TPicture.Create;
    New(Graf);
    Graf^.Col := Col;
    Graf^.Line := Line;
    Graf^.Picture := Dib;
    Dib.Metafile.Assign(Picture);
    Pag^.Graphics.Add(Graf);
  end;
end;

procedure TRSPrinter.Box(Line1,Col1,Line2,Col2 : byte; Kind : TLineType);// ESCRIBE UN RECTANGULO
var
  Pag: PPage;
  VLine: PVertLine;
  HLine: PHorizLine;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];
    New(VLine);
    VLine^.Col := Col1;
    VLine^.Line1 := Line1;
    VLine^.Line2 := Line2;
    VLine^.Kind := Kind;
    Pag^.VerticalLines.Add(VLine);
    New(VLine);
    VLine^.Col := Col2;
    VLine^.Line1 := Line1;
    VLine^.Line2 := Line2;
    VLine^.Kind := Kind;
    Pag^.VerticalLines.Add(VLine);
    New(HLine);
    HLine^.Col1 := Col1;
    HLine^.Col2 := Col2;
    HLine^.Line := Line1;
    HLine^.Kind := Kind;
    Pag^.HorizLines.Add(HLine);
    New(HLine);
    HLine^.Col1 := Col1;
    HLine^.Col2 := Col2;
    HLine^.Line := Line2;
    HLine^.Kind := Kind;
    Pag^.HorizLines.Add(HLine);
    if Line2 > Pag^.PrintedLines then
      Pag^.PrintedLines := Line2;
  end;
end;

procedure TRSPrinter.LineH(Line,Col1,Col2 : byte; Kind : TLineType); // ESCRIBE UNA LINEA HORIZONTAL
var
  HLine: PHorizLine;
  Pag: PPage;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];
    New(HLine);
    Pag.HorizLines.Add(HLine);
    HLine^.Col1 := Col1;
    HLine^.Col2 := Col2;
    HLine^.Line := Line;
    HLine^.Kind := Kind;

    if Line > Pag^.PrintedLines then
      Pag^.PrintedLines := Line;
  end;
end;

procedure TRSPrinter.LineV(Line1,Line2,Col : byte; Kind : TLineType);             // ESCRIBE UNA LINEA VERTICAL
var
  VLine: PVertLine;
  Pag: PPage;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];
    New(VLine);
    Pag.VerticalLines.Add(VLine);
    VLine^.Col := Col;
    VLine^.Line1 := Line1;
    VLine^.Line2 := Line2;
    VLine^.Kind := Kind;

    if Line2 > Pag^.PrintedLines then
      Pag^.PrintedLines := Line2;
  end;
end;

procedure TRSPrinter.PreviewReal;                            // MUESTRA LA IMPRESION EN PANTALLA
begin
  PreviewForm := TPrintPreview.Create(self);
  TPrintPreview(PreviewForm).RSPrinter := self;
  PreviewForm.ShowModal;
  PreviewForm.Free;
end;


procedure TRSPrinter.Print;
begin
  if FPreview = ppYes then
    PreviewReal
  else if FPreview = ppNo then
    PrintAll
  else
    PreviewReal;
end;

procedure TRSPrinter.BuildPage(Number : integer; Page : TMetaFile;ToPrint:boolean;Mode : TPrinterMode; ShowDialog:boolean);
var
  Pagina : PPage;
  Grafico : PGraphic;
  Escritura : PWrite;
  EA : integer; // LINEA ACTUAL
  DA : integer;
  LineaHorizontal : PHorizLine;
  LineaVertical : PVertLine;
  Derecha, Abajo : integer;
  Texto : TMetafile;
  Ancho, Alto : Integer;
  MargenIzquierdo, MargenSuperior : integer;
  AltoDeLinea, AnchoDeColumna : double;
  CantColumnas,i : integer;
  WaitForm : TForm;
  WaitPanel : TPanel;
  BtnCancelar : TButton;
begin
  if ToPrint then
    begin
      if DobleWide in FastFont then
        CantColumnas := 42
      else if Compress in FastFont then
        CantColumnas := 136
      else
        CantColumnas := 81;
      With TMetaFileCanvas.Create(Page,0) do
        try
          Brush.Color := clWhite;
          AnchoDeColumna := 640/CantColumnas;//(Hoja.Width-(2*MargenIzquierdo))/CantColumnas;
          AltoDeLinea := 12;//Round(AnchoDeColumna/Ancho*Alto);
          Page.Width := Round(AnchoDeColumna*CantColumnas);//Round(GetDeviceCaps(Printer.Handle,PHYSICALWIDTH)/GetDeviceCaps(Printer.Handle,LOGPIXELSX)*80);
          Page.Height := Round(AltoDeLinea*(Trunc(PageHeightP*6)-2));
          Font.Name := FWinFont;
          Font.Size := 18;
          Font.Pitch := fpFixed;
          Ancho := TextWidth('X');
          Alto := TextHeight('X');
          FillRect(Rect(0,0,(Page.Width)-1,(Page.Height-1))); //80*(Ancho),66*Alto));
          Pagina := LasPaginas.Items[Number-1];

          if Mode = pmWindows then
            begin
              For DA := 0 to Pagina.Graphics.Count-1 do
                begin
                  Grafico := Pagina.Graphics[DA];
                  Draw(Round(AnchoDeColumna*(Grafico^.Col)),Round(FWinSupMargin+AltoDeLinea*(Grafico^.Line)),Grafico^.Picture.Graphic);
                end;
            end;
          For EA := 0 to Pagina.Writed.Count-1 do
            begin
              Escritura := Pagina.Writed.Items[EA];
              if Escritura^.Line <= Lines then
                begin
                  Texto := TMetaFile.Create;
                  Texto.Width := Ancho*Length(Escritura^.Text);
                  Texto.Height := Alto;
                  With TMetaFileCanvas.Create(Texto,0) do
                    try
                      Font.Name := FWinFont;
                      Font.Size := 18;
                      Font.Pitch := fpFixed;
                      if Bold in Escritura^.FastFont then
                        Font.Style := Font.Style + [fsBold]
                      else
                        Font.Style := Font.Style - [fsBold];
                      if Italic in Escritura^.FastFont then
                        Font.Style := Font.Style + [fsItalic]
                      else
                        Font.Style := Font.Style - [fsItalic];
                      if Underline in Escritura^.FastFont then
                        Font.Style := Font.Style + [fsUnderline]
                      else
                        Font.Style := Font.Style - [fsUnderline];
                      TextOut(0,0,Escritura^.Text);
                    finally
                      Free
                    end;
                  if (DobleWide in Escritura^.FastFont) and not (DobleWide in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*Page.Width/42),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto)
                  else if (Compress in Escritura^.FastFont) and not (Compress in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*Page.Width/140),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto)
                  else
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col+Length(Escritura^.Text))),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto);
                  Texto.Free;
                end;
              Application.ProcessMessages;
            end;
          for i := 0 to Pagina.VerticalLines.Count-1 do
            begin
              LineaVertical := Pagina.VerticalLines[i];
              if LineaVertical.Col <= CantColumnas then
                begin
                  If LineaVertical.Kind = ltSingle then
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                      if LineaVertical^.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                    end
                  else // Double;
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                      if LineaVertical^.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      MoveTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                      if LineaVertical^.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                    end;
                end;
              Application.ProcessMessages;
            end;
          for i := 0 to Pagina.HorizLines.Count-1 do
            begin
              LineaHorizontal := Pagina.HorizLines[i];
              if LineaHorizontal.Line <= Lines then
                begin
                  If LineaHorizontal.Kind = ltSingle then
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2));
                      if LineaHorizontal^.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2));
                    end
                  else // Double;
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1);
                      if LineaHorizontal^.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1)
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1);
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1);
                      if LineaHorizontal^.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1)
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1);
                    end;
                end;
              Application.ProcessMessages;
            end;
         finally
          Free;
        end;
    end
  else // ToPreview
    begin
      if ShowDialog then
        begin
          WaitForm := TForm.CreateNew(self);
          WaitForm.Caption := 'Gerando o preview';
          WaitForm.FormStyle := fsStayOnTop;
          WaitForm.BorderStyle := bsDialog;
          WaitForm.BorderIcons := [];
          WaitForm.Height := 100;
          WaitForm.Width := 160;
          WaitForm.Position := poScreenCenter;
          Cancelado := False;
          BtnCancelar := TButton.Create(WaitForm);
          BtnCancelar.Parent := WaitForm;
          BtnCancelar.Caption := '&Cancelar';
          BtnCancelar.Align := alBottom;
          BtnCancelar.OnClick := BtnCancelarClick;
          WaitPanel := TPanel.Create(WaitForm);
          WaitPanel.Parent := WaitForm;
          WaitPanel.Align := alClient;
          WaitPanel.Caption := 'Por favor espere...';
          WaitForm.Show;
          WaitForm.Update;
        end
      else
        WaitForm := nil;
      if DobleWide in FastFont then
        CantColumnas := 42
      else if Compress in FastFont then
        CantColumnas := 136
      else
        CantColumnas := 81;
      With TMetaFileCanvas.Create(Page,0) do
        try
          Brush.Color := clWhite;

         MargenIzquierdo:=9;
         MargenSuperior:=9;


          AnchoDeColumna := 640/CantColumnas;//(Hoja.Width-(2*MargenIzquierdo))/CantColumnas;
          AltoDeLinea := 12;//Round(AnchoDeColumna/Ancho*Alto);

          Page.Width := Round(AnchoDeColumna*CantColumnas)+2*MargenIzquierdo;//Round(GetDeviceCaps(Printer.Handle,PHYSICALWIDTH)/GetDeviceCaps(Printer.Handle,LOGPIXELSX)*80);
          Page.Height := Round(AltoDeLinea*(Trunc(PageHeightP*6)-2))+2*MargenSuperior;
          Font.Name := FWinFont;
          Font.Size := 18;
          Font.Pitch := fpFixed;
          Ancho := TextWidth('X');
          Alto := TextHeight('X');
          FillRect(Rect(0,0,(Page.Width)-1,(Page.Height)-1)); //80*(Ancho),66*Alto));
          Rectangle(0,0,Page.Width,Page.Height);
          MargenSuperior := MargenSuperior+FWinSupMargin;
          Pen.Style := psSolid;
          Pen.Color := clBlack;
          Pagina := LasPaginas.Items[Number-1];
          if Mode = pmWindows then
            begin
              For DA := 0 to Pagina.Graphics.Count-1 do
                if not Cancelado then
                  begin
                    Grafico := Pagina.Graphics[DA];
                    Draw(Round(MargenIzquierdo+AnchoDeColumna*(Grafico^.Col)),Round(MargenSuperior+AltoDeLinea*(Grafico^.Line))+1,Grafico^.Picture.Graphic);
                  end;
            end;
          For EA := 0 to Pagina.Writed.Count-1 do
            if not Cancelado then
              begin
                Escritura := Pagina.Writed.Items[EA];
                if Escritura^.Line <= Lines then
                  begin
                    Texto := TMetaFile.Create;
                    Texto.Width := Ancho*Length(Escritura^.Text);
                    Texto.Height := Alto;
                    With TMetaFileCanvas.Create(Texto,0) do
                      try
                        Font.Name := FWinFont;
                        Font.Size := 18;
                        Font.Pitch := fpFixed;
                        if Bold in Escritura^.FastFont then
                          Font.Style := Font.Style + [fsBold]
                        else
                          Font.Style := Font.Style - [fsBold];
                        if Italic in Escritura^.FastFont then
                          Font.Style := Font.Style + [fsItalic]
                        else
                          Font.Style := Font.Style - [fsItalic];
                        if Underline in Escritura^.FastFont then
                          Font.Style := Font.Style + [fsUnderline]
                        else
                          Font.Style := Font.Style - [fsUnderline];
                          TextOut(0,0,Escritura^.Text);
                      finally
                        Free
                      end;
                    if (DobleWide in Escritura^.FastFont) and not (DobleWide in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*Page.Width/42),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto)
                    else if (Compress in Escritura^.FastFont) and not (Compress in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*Page.Width/140),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto)
                    else
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col+Length(Escritura^.Text))),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto);
                    Texto.Free;
                  end;
                Application.ProcessMessages;
              end;
          for i := 0 to Pagina.VerticalLines.Count-1 do
            if not Cancelado then
              begin
                LineaVertical := Pagina.VerticalLines[i];
                if LineaVertical.Col <= CantColumnas then
                  begin

                  (*   *)
                    If LineaVertical.Kind = ltSingle then
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                        if LineaVertical^.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      end
                    else // Double;
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                        if LineaVertical^.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2-1),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line1-1)+AltoDeLinea/2));
                        if LineaVertical^.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical^.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical^.Col)+AnchoDeColumna/2+1),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      end;
                      (*   *)

                  end;
                Application.ProcessMessages;
              end;
          for i := 0 to Pagina.HorizLines.Count-1 do
            if not Cancelado then
              begin
                LineaHorizontal := Pagina.HorizLines[i];
                if LineaHorizontal.Line <= Lines then
                  begin
                    If LineaHorizontal.Kind = ltSingle then
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2));
                        if LineaHorizontal^.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2));
                      end
                    else // Double;
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1);
                        if LineaHorizontal^.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1)
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)-1);
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1);
                        if LineaHorizontal^.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal^.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1)
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal^.Line-1)+AltoDeLinea/2)+1);
                      end;
                  end;
                Application.ProcessMessages;
              end;
         finally
          Free;
        end;
      if ShowDialog then
        WaitForm.Free;
    end;
end;

function TRSPrinter.GetPrintingWidth : integer;
begin
  result := Printer.PageWidth;
end;

function TRSPrinter.GetPrintingHeight : integer;
begin
  result := Printer.PageHeight;
end;

function TRSPrinter.PrintPage(Number:integer) : boolean; // MANDA UNA PAGINA A LA IMPRESORA
var
  Pag : PPage;
  Esc : pWrite;
  Imagen : TMetaFile;
  Resultado : Boolean;
  LP : TList;
  j : integer;
  LV : pVertLine;
  LH : pHorizLine;
  Job : pPrintJob;
begin
  if FModo = pmWindows then
  begin
   try
     Printer.Title := Title + ' (Página Nº'+' '+IntToStr(Number)+')';
     Printer.Copies := fCopias;
     Printer.BeginDoc;
     Imagen := TMetaFile.Create;
     Imagen.Width := 640; // PrintingWidth{PageWidth} div 4; // 640;
     Imagen.Height := 12*Lines; //PrintingHeight{PageHeight} div 4; // 1056;
     try
       BuildPage(Number,Imagen,True,pmWindows,False);
       Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Imagen);
       Printer.EndDoc;
     except
     end;
     Imagen.Free;
     Resultado := True;
   except
     Resultado := False;
   end;
  end
  else
  begin
    LP := TList.Create;
    New(Pag);
    LP.Add(Pag);
    Pag^.PrintedLines := PPage(LasPaginas[Number-1])^.PrintedLines + 1;
    Pag^.Writed := TList.Create;
    Pag^.VerticalLines := TList.Create;
    Pag^.HorizLines := TList.Create;

    for j := 0 to PPage(LasPaginas[Number-1])^.Writed.Count-1 do
    begin
      new(Esc);
      Esc^ := pWrite(PPage(LasPaginas[Number-1])^.Writed[j])^;
      Pag.Writed.Add(Esc);
    end;

    for j := 0 to PPage(LasPaginas[Number-1])^.VerticalLines.Count-1 do
    begin
      new(LV);
      LV^ := pVertLine(PPage(LasPaginas[Number-1])^.VerticalLines[j])^;
      Pag.VerticalLines.Add(LV);
    end;

    for j := 0 to PPage(LasPaginas[Number-1])^.HorizLines.Count-1 do
    begin
      new(LH);
      LH^ := pHorizLine(PPage(LasPaginas[Number-1])^.HorizLines[j])^;
      Pag.HorizLines.Add(LH);
    end;

    New(Job);
    Job.Name := Title + 'Pág. '+IntToStr(Number-1);
    Job.PageSize := PageSize;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.FCopias := Copies;
    Job.FFuente := FastFont;
    Job.FPort := FastPort;
    Job.FLineas := Lines;
    Job.FTransliterate := Transliterate;
    Job.PRNNormal := PRNNormal;
    Job.PRNBold := PRNBold;
    Job.PRNWide := PRNWide;
    Job.PRNItalics := PRNItalics;
    Job.PRNULineON := PRNULineON;
    Job.PRNULineOFF := PRNULineOFF;
    Job.PRNCompON := PRNCompON;
    Job.PRNCompOFF := PRNCompOFF;
    Job.PRNSetup := PRNSetup;
    Job.PRNReset := PRNReset;
    Job.PRNSelLength := PRNSelLength;
    if not CurrentlyPrinting then
    begin
      CurrentlyPrinting := True;
      PrintThread := TPrintThread.Create(True);
      PrintThread.RSPrinter := self;
      PrintThread.Jobs := TThreadList.Create;
      TrayIcon := TRSPrintTray.Create(self);
      PrintingCanceled := False;
      PrintingPaused:= False;
      PrintingCancelAll := False;
    end;
    PrintThread.RSPrinter := self;
    PrintThread.Jobs.Add(Job);
    PrintThread.OnTerminate := PrintThreadTerminate;
    PrintThread.FreeOnTerminate := True;
    PrintThread.Resume;
    Resultado := True;
  end;

  PrintPage := Resultado;
end;

procedure TRSPrinter.PrintAll; // MANDA LA IMPRESION A LA IMPRESORA
var
  Pag : PPage;
  Esc : PWrite;
  Imagen : TMetaFile;
  i,j : integer;
  LH : PHorizLine;
  LV : PVertLine;
  LP : TList;
  Job : PPrintJob;
begin // IMPRIMIR
  if FModo = pmWindows then
  begin
    Printer.Title := Title;
    Printer.Copies := FCopias;
    Printer.BeginDoc;
    Imagen := TMetaFile.Create;
    Imagen.Width := 640; //PrintingWidth{PageWidth} div 4; // 640;
    Imagen.Height := 12 * Lines; //PrintingHeight{PageHeight} div 4; // 1056;
    BuildPage(1,Imagen,True,pmWindows,False);
    Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Imagen);
    if LasPaginas.Count > 1 then
      for i := 2 to LasPaginas.Count do
      begin
        Printer.NewPage;
        BuildPage(i,Imagen,True,pmWindows,False);
        Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Imagen);
      end;
    Printer.EndDoc;
    Imagen.Free;
  end
  else
  begin
    LP := TList.Create;

    for i := 0 to LasPaginas.Count-1 do
    begin
      New(Pag);
      LP.Add(Pag);
      Pag^.PrintedLines := PPage(LasPaginas[i])^.PrintedLines + 1;
      Pag^.Writed := TList.Create;
      Pag^.VerticalLines := TList.Create;
      Pag^.HorizLines := TList.Create;

      for j := 0 to PPage(LasPaginas[i])^.Writed.Count-1 do
      begin
        new(Esc);
        Esc^ := pWrite(PPage(LasPaginas[i])^.Writed[j])^;
        Pag.Writed.Add(Esc);
      end;

      for j := 0 to PPage(LasPaginas[i])^.VerticalLines.Count-1 do
      begin
        new(LV);
        LV^ := pVertLine(PPage(LasPaginas[i])^.VerticalLines[j])^;
        Pag.VerticalLines.Add(LV);
      end;

      for j := 0 to PPage(LasPaginas[i])^.HorizLines.Count-1 do
      begin
        new(LH);
        LH^ := pHorizLine(PPage(LasPaginas[i])^.HorizLines[j])^;
        Pag.HorizLines.Add(LH);
      end;
    end;

    New(Job);
    Job.Name := Title;
    Job.PageSize := PageSize;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.FCopias := Copies;
    Job.FFuente := FastFont;
    Job.FPort := FastPort;
    Job.FLineas := Lines;
    Job.FTransliterate := Transliterate;
    Job.PRNNormal := PRNNormal;
    Job.PRNBold := PRNBold;
    Job.PRNWide := PRNWide;
    Job.PRNItalics := PRNItalics;
    Job.PRNULineON := PRNULineON;
    Job.PRNULineOFF := PRNULineOFF;
    Job.PRNCompON := PRNCompON;
    Job.PRNCompOFF := PRNCompOFF;
    Job.PRNSetup := PRNSetup;
    Job.PRNReset := PRNReset;
    Job.PRNSelLength := PRNSelLength;
    if not CurrentlyPrinting then
      begin
        CurrentlyPrinting := True;
        PrintThread := TPrintThread.Create(True);
        PrintThread.RSPrinter := self;
        PrintThread.Jobs := TThreadList.Create;
        TrayIcon := TRSPrintTray.Create(Owner);
        PrintingCanceled := False;
        PrintingPaused := False;
        PrintingCancelAll := False;
      end;
    PrintThread.RSPrinter := self;
    PrintThread.Jobs.Add(Job);
    PrintThread.OnTerminate := PrintThreadTerminate;
    PrintThread.FreeOnTerminate := True;
    PrintThread.Resume;
  end;
end; // IMPRIMIR

procedure TRSPrinter.NewPage;                        // CREA UNA NUEVA PAGINA
var
  Pag : PPage;
begin
  New(Pag);
  LasPaginas.Add(Pag);
  Pag^.Writed := TList.Create;
  Pag^.VerticalLines := TList.Create;
  Pag^.HorizLines := TList.Create;
  Pag^.Graphics := TList.Create;
  Pag^.PrintedLines := 0;
  PaginaActual := LasPaginas.IndexOf(Pag);
  PageNo := PaginaActual+1;
end;

procedure TRSPrinter.SetModelName(Name : String);
begin
  if FModo = pmWindows then
  begin
    if Printer.Printers.IndexOf(Name) <> -1 then
      Printer.PrinterIndex := Printer.Printers.IndexOf(Name);
  end
  else
  begin
    if Name = 'Cannon F-60' then
      SetModel(Cannon_F60)
    else if Name = 'Cannon laser' then
      SetModel(Cannon_Laser)
    else if Name = 'Epson FX/LX/LQ' then
      SetModel(Epson_FX)
    else if Name = 'Epson Stylus' then
      SetModel(Epson_Stylus)
    else if Name = 'HP Deskjet' then
      SetModel(HP_Deskjet)
    else if Name = 'HP Laserjet' then
      SetModel(HP_Laserjet)
    else if Name = 'HP Thinkjet' then
      SetModel(HP_Thinkjet)
    else if Name = 'IBM Color Jet' then
      SetModel(IBM_Color_Jet)
    else if Name = 'IBM PC Graphics' then
      SetModel(IBM_PC_Graphics)
    else if Name = 'IBM Proprinter' then
      SetModel(IBM_Proprinter)
    else if Name = 'NEC 3500' then
      SetModel(NEC_3500)
    else if Name = 'NEC Pinwriter' then
      SetModel(NEC_Pinwriter)
    else if Name = 'Mp20Mi' then
      SetModel(Mp20Mi);
  end;
end;

function TRSPrinter.GetModelRealName(Model : TPrinterModel) : String;
begin
  case Model of
    Cannon_F60 : Result := 'Cannon F-60';
    Cannon_Laser : Result := 'Cannon Laser';
    Epson_FX : Result := 'Epson FX/LX/LQ';
    Epson_Stylus : Result := 'Epson Stylus';
    HP_Deskjet : Result := 'HP Deskjet';
    HP_Laserjet : Result := 'HP Laserjet';
    HP_Thinkjet : Result := 'HP Thinkjet';
    IBM_Color_Jet : Result := 'IBM Color Jet';
    IBM_PC_Graphics : Result := 'IBM PC Graphics';
    IBM_Proprinter : Result := 'IBM Proprinter';
    NEC_3500 : Result := 'NEC 3500';
    NEC_Pinwriter : Result := 'NEC Pinwriter';
    Mp20Mi : Result := 'Mp20Mi';
  end;
end;

function TRSPrinter.GetModelName : String;
begin
  Result := GetModelRealName(FModelo);
end;

procedure TRSPrinter.GetModels(Models : TStrings);
var
  i : integer;
begin
  Models.Clear;
  for i := Ord(Low(TPrinterModel)) to Ord(High(TPrinterModel)) do
    Models.Add(GetModelRealName(TPrinterModel(i)));
end;

{ TPrintThread }

procedure TPrintThread.Execute;
var
  Bien : boolean;
  Copias,i : integer;
  List : TList;
  Cant : integer;
begin
  List := Jobs.LockList;
  Cant := List.Count;
  if Cant > 0 then
    begin
      Job := pPrintJob(List[0]);
      PrintingJobName := Job^.Name;
      Synchronize(SetJobName);
    end
  else
    begin
      Job := nil;
      PrintingJobName := '';
      Synchronize(SetJobName);
    end;
  Jobs.UnlockList;
  While PrintingPaused and not PrintingCanceled and not PrintingCancelAll do
    Sleep(100);
  While (Job <> nil) do
    begin
      with Job^ do
        begin
          Bien := True;
          for Copias := 1 to FCopias do
            for i := 1 to LasPaginas.Count do
              if not PrintingCanceled and not PrintingCancelAll and Bien then
                Bien := Bien and PrintFast(i);
          While LasPaginas.Count > 0 do
            begin
              with PPage(LasPaginas[0])^ do
                begin
                  while Writed.Count > 0 do
                    begin
                      Dispose(pWrite(Writed[0]));
                      Writed.Delete(0);
                    end;
                  Writed.Free;
                  while VerticalLines.Count > 0 do
                    begin
                      Dispose(pVertLine(VerticalLines[0]));
                      VerticalLines.Delete(0);
                    end;
                  VerticalLines.Free;
                  while HorizLines.Count > 0 do
                    begin
                      Dispose(pHorizLine(HorizLines[0]));
                      HorizLines.Delete(0);
                    end;
                  HorizLines.Free;
                end;
              Dispose(PPage(LasPaginas[0]));
              LasPaginas.Delete(0);
            end;
          LasPaginas.Free;
        end;
      Dispose(Job);
      PrintingCanceled := False;
      List := Jobs.LockList;
      List.Delete(0);
      Cant := List.Count;
      if Cant > 0 then
        begin
          Job := pPrintJob(List[0]);
          PrintingJobName := Job^.Name;
          Synchronize(SetJobName);
        end
      else
        begin
          Job := nil;
          PrintingJobName := '';
          Synchronize(SetJobName);
        end;
    end;
  PrintingJobName := '';
end;

function TPrintThread.PrintFast(Number : integer): boolean;
var
(*Verificar*)
//  Impresora : TextFile; // Impresora
  LA : integer; // LINEA ACTUAL
  i,j,Contador : integer;
  LineaAImprimir : string;
  Pagina : PPage;
  Fuente : TFastFont;
  HLinea : PHorizLine;
  VLinea : PVertLine;
  UltimaEscritura : integer;
  Resultado : boolean;
  ListaImpressoras : tstringlist;
  Printerid: integer;

  procedure ImprimirCodigo(Codigo:string);
  var
    Sub : string;
    Cod : byte;
    P : byte;
  begin
    Sub := Codigo;
    While Length(Sub) > 0 do
      begin
        P := Pos(#32,Sub);
        if P = 0 then // ES EL ULTIMO CODIGO
          begin
           try
            Cod := StrToInt(Sub);
(*Verificar*)
            toprn(chr(Cod));
//            Write(Impresora,chr(Cod));
            Sub := '';
           except
           end;
          end
        else
          begin // HAY MAS CODIGOS
           try
            Cod := StrToInt(Copy(Sub,1,P-1));
(*Verificar*)
            toprn(chr(Cod));
//            Write(Impresora,chr(Cod));
            Sub := Copy(Sub,P+1,Length(Sub)-3);
           except
           end;
          end;
      end;
  end;

  procedure MaxX(var Linea:string; Col : byte);
  begin
    While Length(Linea)<Col do
      Linea := Linea + ' ';
  end;

  procedure ImprimirLinea;
  var
    Escritura : PWrite;
    i,Contador : integer;
    Columna : byte;
    ImprimioLinea : boolean;
    Txt : string;
  begin

    if Pagina.Writed.Count = UltimaEscritura then
      begin
        if Fuente <> Job.FFuente then
          begin // PONEMOS LA FUENTE POR DEFAULT
            Fuente := Job.FFuente;
            ImprimirCodigo(Job.PRNNormal);
            if Bold in Fuente then
              ImprimirCodigo(Job.PRNBold);
            if Italic in Fuente then
              ImprimirCodigo(Job.PRNItalics);
            if DobleWide in Fuente then
              ImprimirCodigo(Job.PRNWide)
            else if Compress in Fuente then
              ImprimirCodigo(Job.PRNCompON)
            else
              ImprimirCodigo(Job.PRNCompOFF);
            if Underline in Fuente then
              ImprimirCodigo(Job.PRNULineON)
            else
              ImprimirCodigo(Job.PRNULineOFF);
          end;
        if Job.FTransliterate and (LineaAImprimir<>'') then
          CharToOemBuff(PChar(@LineaAImprimir[1]), PansiChar(@LineaAImprimir[1]),Length(LineaAImprimir));
(*Verificar*)
        toprnln(LineaAImprimir);
//        Writeln(Impresora,LineaAImprimir);
      end
    else
      begin
        ImprimioLinea := False;
        Contador := UltimaEscritura;
        Columna := 1;
        Escritura := Pagina.Writed.Items[Contador];
        While (Contador < Pagina.Writed.Count) and (Escritura^.Line <= LA) do
          begin
            if Escritura^.Line = LA then
              begin
                ImprimioLinea := True;
                UltimaEscritura := Contador;
                MaxX(LineaAImprimir,Escritura^.Col+Length(Escritura^.Text));
                While Columna < Escritura^.Col do
                  begin
                    if (LineaAImprimir[Columna] <> #32) and
                       (Fuente <> Job.FFuente) then
                      begin // PONEMOS LA FUENTE POR DEFAULT
                        Fuente := Job.FFuente;
                        ImprimirCodigo(Job.PRNNormal);
                        if Bold in Fuente then
                          ImprimirCodigo(Job.PRNBold);
                        if Italic in Fuente then
                          ImprimirCodigo(Job.PRNItalics);
                        if DobleWide in Fuente then
                          ImprimirCodigo(Job.PRNWide)
                        else if Compress in Fuente then
                          ImprimirCodigo(Job.PRNCompON)
                        else
                          ImprimirCodigo(Job.PRNCompOFF);
                        if Underline in Fuente then
                          ImprimirCodigo(Job.PRNULineON)
                        else
                          ImprimirCodigo(Job.PRNULineOFF);
                      end;
(*Verificar*)
                    toprn(LineaAImprimir[Columna]);
                 //   Write(Impresora,LineaAImprimir[Columna]);
                    Inc(Columna);
                  end;
                 if Escritura^.FastFont <> Fuente then
                  begin // PONEMOS LA FUENTE DEL TEXTO
                    Fuente := Escritura^.FastFont;
                    ImprimirCodigo(Job.PRNNormal);
                    if Bold in Fuente then
                      ImprimirCodigo(Job.PRNBold);
                    if Italic in Fuente then
                      ImprimirCodigo(Job.PRNItalics);
                    if DobleWide in Fuente then
                      ImprimirCodigo(Job.PRNWide)
                    else if Compress in Fuente then
                      ImprimirCodigo(Job.PRNCompON)
                    else
                      ImprimirCodigo(Job.PRNCompOFF);
                    if Underline in Fuente then
                      ImprimirCodigo(Job.PRNULineON)
                    else
                      ImprimirCodigo(Job.PRNULineOFF);
                  end;
                Txt := Escritura^.Text;
                if Job.FTransliterate and (Txt<>'') then
                  CharToOemBuff(PChar(@Txt[1]), PansiChar(@Txt[1]),Length(Txt));
(*Verificar*)
                toprn(Txt);
                //Write(Impresora,Txt);
                if (Compress in Fuente) and not(Compress in Job.FFuente) then
                  begin
(*Verificar*)
                    for i := 1 to Length(Escritura^.Text) do
                      toprn(#8);
                      //Write(Impresora,#8);
                    if (Length(Escritura^.Text)*6) mod 10 = 0 then
                      i := Columna + (Length(Escritura^.Text) *6) div 10
                    else
                      i := Columna + (Length(Escritura^.Text) *6) div 10;
                    Fuente := Fuente - [Compress];
                    ImprimirCodigo(Job.PRNNormal);
                    if Bold in Fuente then
                      ImprimirCodigo(Job.PRNBold);
                    if Italic in Fuente then
                      ImprimirCodigo(Job.PRNItalics);
                    if DobleWide in Fuente then
                      ImprimirCodigo(Job.PRNWide)
                    else if Compress in Fuente then
                      ImprimirCodigo(Job.PRNCompON)
                    else
                      ImprimirCodigo(Job.PRNCompOFF);
                    if Underline in Fuente then
                      ImprimirCodigo(Job.PRNULineON)
                    else
                      ImprimirCodigo(Job.PRNULineOFF);
                    While Columna <= i do
                      begin
(*Verificar*)
                        toprn(#32);
                        //Write(Impresora,#32);
                        Inc(Columna);
                      end;
                  end
                else
                  Columna := Columna + Length(Escritura^.Text);
              end;
            Inc(Contador);
            if Contador < Pagina.Writed.Count then
              Escritura := Pagina.Writed.Items[Contador];
          end;
        if ImprimioLinea then
          begin
            if Fuente <> Job.FFuente then
              begin // PONEMOS LA FUENTE POR DEFAULT
                Fuente := Job.FFuente;
                ImprimirCodigo(Job.PRNNormal);
                if Bold in Fuente then
                  ImprimirCodigo(Job.PRNBold);
                if Italic in Fuente then
                  ImprimirCodigo(Job.PRNItalics);
                if DobleWide in Fuente then
                  ImprimirCodigo(Job.PRNWide)
                else if Compress in Fuente then
                  ImprimirCodigo(Job.PRNCompON)
                else
                  ImprimirCodigo(Job.PRNCompOFF);
                if Underline in Fuente then
                  ImprimirCodigo(Job.PRNULineON)
                else
                  ImprimirCodigo(Job.PRNULineOFF);
              end;
            While Columna <= Length(LineaAImprimir) do
              begin
(*Verificar*)
                toprn(LineaAImprimir[Columna]);
                //Write(Impresora,LineaAImprimir[Columna]);
                Inc(Columna);
              end;
(*Verificar*)
            toprnln('');
            //WriteLn(Impresora);
          end
        else
          begin
            if Fuente <> Job.FFuente then
              begin // PONEMOS LA FUENTE POR DEFAULT
                Fuente := Job.FFuente;
                ImprimirCodigo(Job.PRNNormal);
                if Bold in Fuente then
                  ImprimirCodigo(Job.PRNBold);
                if Italic in Fuente then
                  ImprimirCodigo(Job.PRNItalics);
                if DobleWide in Fuente then
                  ImprimirCodigo(Job.PRNWide)
                else if Compress in Fuente then
                  ImprimirCodigo(Job.PRNCompON)
                else
                  ImprimirCodigo(Job.PRNCompOFF);
                if Underline in Fuente then
                  ImprimirCodigo(Job.PRNULineON)
                else
                  ImprimirCodigo(Job.PRNULineOFF);
              end;
            if Job.FTransliterate and (LineaAImprimir<>'') then
              AnsiToOemBuff(PansiChar(LineaAImprimir[1]), PansiChar(LineaAImprimir[1]),Length(LineaAImprimir));
(*Verificar*)
            toprnln(pchar(LineaAImprimir));
            //Writeln(Impresora,LineaAImprimir);
          end;
      end;
  end;

begin
  ListaImpressoras:= tstringlist.Create;
  EnumPrt(ListaImpressoras, PrinterId);

  While PrintingPaused and not PrintingCanceled and not PrintingCancelAll do
    Sleep(100);
     try
       Resultado := True;
     //  AssignFile(Impresora,Job.FPort);

(*Verificar*)
//       ReWrite(Impresora);
       StartPrint(ListaImpressoras[printer.Printerindex],'TESTE DE IMPRESSAO','',1);

       ImprimirCodigo(Job.PRNReset);
       ImprimirCodigo(Job.PRNSetup);

           ImprimirCodigo(Job.PRNSelLength);
(*Verificar*)
//           Write(Impresora,#84);
       toprn(#84);


//         end;
       Fuente := Job.FFuente;
       ImprimirCodigo(Job.PRNNormal);
       if Bold in Fuente then
         ImprimirCodigo(Job.PRNBold);
       if Italic in Fuente then
         ImprimirCodigo(Job.PRNItalics);
       if DobleWide in Fuente then
         ImprimirCodigo(Job.PRNWide)
       else if Compress in Fuente then
         ImprimirCodigo(Job.PRNCompON)
       else
         ImprimirCodigo(Job.PRNCompOFF);
       if Underline in Fuente then
         ImprimirCodigo(Job.PRNULineON)
       else
         ImprimirCodigo(Job.PRNULineOFF);
       UltimaEscritura := 0;
       Pagina := Job.LasPaginas.Items[Number-1];
       LA := 1;
       while (LA<Min(Pagina.PrintedLines,Job.FLineas)) and (not PrintingCanceled) and not (PrintingCancelAll) do
         begin // SE IMPRIMEN TODAS LAS LINEAS
           LineaAImprimir := '';
           // ANALIZO PRIMERO LAS LINEAS HORIZONTALES
           For Contador := 0 to Pagina.HorizLines.Count-1 do
             begin
               HLinea := Pagina.HorizLines.Items[Contador];
               if HLinea^.Line = LA then // ES EN ESTA LINEA
                 begin
                   MaxX(LineaAImprimir,HLinea^.Col2);
                   if HLinea^.Kind = ltSingle then
                     begin
                       for i := HLinea^.Col1 to HLinea^.Col2 do
                         LineaAImprimir[i] := #196;
                     end
                   else // LINEA DOBLE
                     begin
                       for i := HLinea^.Col1 to HLinea^.Col2 do
                         LineaAImprimir[i] := #205;
                     end;
                 end;
             end;
           // AHORA SE ANALIZAN LAS LINEAS VERTICALES
           For Contador := 0 to Pagina.VerticalLines.Count-1 do
             begin
               VLinea := Pagina.VerticalLines.Items[Contador];
                 if (VLinea^.Line1<=LA) and (VLinea^.Line2>=LA) then
                   begin // LA LINEA PASA POR ESTA LINEA
                     MaxX(LineaAImprimir,VLinea^.Col);
                     if VLinea^.Line1=LA then
                       begin // ES LA PRIMER LINEA
                         if LineaAImprimir[VLinea^.Col] = #196 then // LINEA HORIZONTAL SIMPLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[192,193,194,195,196,197,199,208,210,211,214,215,218]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #194
                                     else
                                       LineaAImprimir[VLinea^.Col] := #210;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #191
                                     else
                                       LineaAImprimir[VLinea^.Col] := #183;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #218
                                     else
                                       LineaAImprimir[VLinea^.Col] := #214;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #194
                                     else
                                       LineaAImprimir[VLinea^.Col] := #210;
                                   end;
                               end;
                           end
                         else if LineaAImprimir[VLinea^.Col] = #205 then // LINEA HORIZONTAL DOBLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[198,200,201,202,203,204,205,206,207,209,212,213]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[181,184,185,187,188,190,202,203,205,206,207,209,216]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #209
                                     else
                                       LineaAImprimir[VLinea^.Col] := #203;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #184
                                     else
                                       LineaAImprimir[VLinea^.Col] := #187;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[181,184,185,187,188,190,202,203,205,206,207,209,216]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #213
                                      else
                                       LineaAImprimir[VLinea^.Col] := #201;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #209
                                     else
                                       LineaAImprimir[VLinea^.Col] := #203;
                                   end;
                               end;
                           end
                         else // HAY OTRO CODIGO
                           begin
                             if VLinea^.Kind = ltSingle then
                               LineaAImprimir[VLinea^.Col] := #179
                             else // Doble
                               LineaAImprimir[VLinea^.Col] := #186;
                           end;
                       end
                     else if VLinea^.Line2=LA then
                       begin // ES LA ULTIMA LINEA
                         if LineaAImprimir[VLinea^.Col] = #196 then // LINEA HORIZONTAL SIMPLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[192,193,194,195,196,197,199,208,210,211,214,215,218]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #193
                                     else
                                       LineaAImprimir[VLinea^.Col] := #207;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #217
                                     else
                                       LineaAImprimir[VLinea^.Col] := #189;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #192
                                     else
                                       LineaAImprimir[VLinea^.Col] := #211;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                     else
                                       LineaAImprimir[VLinea^.Col] := #207;
                                   end;
                               end;
                           end
                         else if LineaAImprimir[VLinea^.Col] = #205 then // LINEA HORIZONTAL DOBLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[181,184,185,187,188,190,202,203,205,206,207,209,216]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[198,200,201,202,203,204,205,206,207,209,212,213]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #207
                                     else
                                       LineaAImprimir[VLinea^.Col] := #202;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #190
                                     else
                                       LineaAImprimir[VLinea^.Col] := #188;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[198,200,201,202,203,204,205,206,207,209,212,213]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #212
                                     else
                                       LineaAImprimir[VLinea^.Col] := #200;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #208
                                     else
                                       LineaAImprimir[VLinea^.Col] := #202;
                                   end;
                               end;
                           end
                         else // HAY OTRO CODIGO
                           begin
                             if VLinea^.Kind = ltSingle then
                               LineaAImprimir[VLinea^.Col] := #179
                             else // Doble
                               LineaAImprimir[VLinea^.Col] := #186;
                           end;
                       end
                     else
                       begin // ES UNA LINEA DEL MEDIO
                         if LineaAImprimir[VLinea^.Col] = #196 then // LINEA HORIZONTAL SIMPLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[192,193,194,195,196,197,199,208,210,211,214,215,218]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #197
                                     else
                                       LineaAImprimir[VLinea^.Col] := #215;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #180
                                     else
                                       LineaAImprimir[VLinea^.Col] := #182;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[180,182,183,189,191,193,194,196,197,208,210,215,217]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #195
                                     else
                                       LineaAImprimir[VLinea^.Col] := #199;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #197
                                     else
                                       LineaAImprimir[VLinea^.Col] := #215;
                                   end;
                               end;
                           end
                         else if LineaAImprimir[VLinea^.Col] = #205 then // LINEA HORIZONTAL DOBLE
                           begin
                             if (VLinea^.Col > 0) and (ord(LineaAImprimir[VLinea^.Col-1])in[198,200,201,202,203,204,205,206,207,209,212,213]) then
                               begin // VIENE DE LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[181,184,185,187,188,190,202,203,205,206,207,209,216]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #216
                                     else
                                       LineaAImprimir[VLinea^.Col] := #206;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #181
                                     else
                                       LineaAImprimir[VLinea^.Col] := #185;
                                   end;
                               end
                             else
                               begin // NO VA PARA LA IZQUIERDA
                                 if (VLinea^.Col < Length(LineaAImprimir)) and (ord(LineaAImprimir[VLinea^.Col+1])in[181,184,185,187,188,190,202,203,205,206,207,209,216]) then
                                   begin // SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #198
                                     else
                                       LineaAImprimir[VLinea^.Col] := #204;
                                   end
                                 else
                                   begin // NO SIGUE A LA DERECHA
                                     if VLinea^.Kind = ltSingle then
                                       LineaAImprimir[VLinea^.Col] := #216
                                     else
                                       LineaAImprimir[VLinea^.Col] := #206;
                                   end;
                               end;
                           end
                         else
                           begin
                             if VLinea^.Kind = ltSingle then
                               LineaAImprimir[VLinea^.Col] := #179
                             else // Doble
                               LineaAImprimir[VLinea^.Col] := #186;
                           end;
                       end;
                   end;
               end;
             ImprimirLinea;
             inc(LA);
             While PrintingPaused and not PrintingCanceled and not PrintingCancelAll do
               Sleep(100);
           end;

(* Cria um salto de picote manual selecionavel nas propriedades da rsprint *)
         if (Job.PageSize=pzContinuous) and
            (Job.PageLength=0) then
           begin
             for j := 1 to RSPrinter.PageContinuousJump do
               toprnln('');
              // WriteLn(Impresora,'');
           end
         else
(*Verificar*)
//           Write(Impresora,#12);
         toprn(#12);
     except
       Resultado := False;
     end;

(*Verificar*)
endprint();
//      {$I-}
//      CloseFile(Impresora);
//      {$I+}
  PrintFast := Resultado;
end;

procedure TRSPrinter.PrintThreadTerminate(Sender : TObject);
begin
  CurrentlyPrinting := False;
  TrayIcon.Free;
  TrayIcon := nil;
end;

procedure TRSPrinter.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if CanClose and CurrentlyPrinting and (MessageDlg('Há páginas a serem impressas. Cancelar? ',mtConfirmation,[mbYes,mbNo],0)=mrNo) then
    CanClose := False;

  if Assigned(OldCloseQuery) then
    OldCloseQuery(Sender,CanClose);
end;

procedure TPrintThread.SetJobName;
begin
  if Job = nil then
    RSPrinter.TrayIcon.Tip :=  ''
  else
  begin
    if PrintingPaused then
      RSPrinter.TrayIcon.Tip := 'Em pausa '+PrintingJobName
    else
      RSPrinter.TrayIcon.Tip := 'Imprimindo '+PrintingJobName;
  end;
end;

{ TRSPrintTray }

procedure TRSPrintTray.CancelClick(Sender: TObject);
begin
  RSPrinter.CancelPrinting;
end;

procedure TRSPrintTray.CancelAllClick(Sender: TObject);
begin
  RSPrinter.CancelAllPrinting;
end;

constructor TRSPrintTray.Create(AOwner: TComponent);
var
  TNIData: TNOTIFYICONDATA;
  PauseItem : TMenuItem;
  CancelItem : TMenuItem;
  CancelAllItem : TMenuItem;
  DivItem : TMenuItem;
begin
  inherited Create(AOwner);

  RSPrinter := TRSPrinter(AOwner);
  FWnd := AllocateHWnd(WndProc);
  FIcon := TIcon.Create;
  FIcon.Handle := LoadIcon(hinstance,'PRINTER');
  FPopUpMenuL := TPopupMenu.Create(self);
  PauseItem := TMenuItem.Create(FPopupMenuL);
  PauseItem.OnClick := PauseClick;
  PauseItem.Caption := 'Em pausa ';
  DivItem := TMenuItem.Create(FPopupMenuL);
  DivItem.Caption := '-';
  CancelItem := TMenuITem.Create(FPopupMenuL);
  CancelItem.OnClick := CancelClick;
  CancelItem.Caption := 'Cancelar a impressão';
  CancelAllItem := TMenuItem.Create(FPopupMenuL);
  CancelAllItem.OnClick := CancelAllClick;
  CancelAllItem.Caption := 'Cancelar toda a impressão';
  FPopUpMenuL.Items.Add(CancelItem);
  FPopUpMenuL.Items.Add(CancelAllItem);
  FPopUpMenuL.Items.Add(DivItem);
  FPopUpMenuL.Items.Add(PauseItem);
  with TNIData do
  begin
    cbSize := TNOTIFYICONDATA.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_ADD, @TNIData);
  end;
end;

destructor TRSPrintTray.Destroy;
var
  TNIData: TNOTIFYICONDATA;
begin
  FPopUpMenuL.Free;
  with TNIData do
  begin
    cbSize := TNOTIFYICONDATA.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_DELETE, @TNIData);
  end;
  FIcon.Free;
  DeallocateHWnd(FWnd);

  inherited Destroy;
end;

procedure TRSPrintTray.DoOnLeftClick;
begin
  if Assigned(FPopupMenuL) then
    DoPopup(1);
end;

procedure TRSPrintTray.DoPopup(i: integer);
var
  Point : TPoint;
begin
  GetCursorPos(Point);
  SetForeGroundWindow(FWnd);
  case i of
    1: FPopupmenuL.Popup(Point.X, Point.Y);
  end;
  PostMessage(0, 0, 0, 0);
end;

procedure TRSPrintTray.PauseClick(Sender: TObject);
begin
  if not TMenuItem(Sender).Checked then
  begin
    TMenuItem(Sender).Checked := True;
    RSPrinter.PausePrinting;
    Tip := 'Em pausa '+PrintingJobName;
  end
  else
  begin
    TMenuItem(Sender).Checked := False;
    RSPrinter.RestorePrinting;
    Tip := 'Imprimindo '+PrintingJobName;
  end;
end;

procedure TRSPrintTray.SetTip(s: string);
begin
  if FTip <> s then
  begin
    FTip := s;
    if not (csDesigning in ComponentState) then
      TBChange;
  end;
end;

procedure TRSPrintTray.TBChange;
var
  TNIData: TNOTIFYICONDATA;
begin
  with TNIData do
  begin
   cbSize := TNOTIFYICONDATAW.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_MODIFY, @TNIData);
  end;
end;

procedure TRSPrintTray.WndProc(var Msg: TMessage);
begin
  with Msg do
  begin
    if Msg = WM_TASKICON then
      case LParamLo of
        WM_LBUTTONDOWN: DoOnLeftClick;
      end
    else
      Result := DefWindowProc(FWnd, Msg, wParam, lParam);
  end;
end;

procedure TRSPrinter.CancelAllPrinting;
begin
  PrintingCancelAll := True;
end;

procedure TRSPrinter.CancelPrinting;
begin
  PrintingCanceled := True;
end;

procedure TRSPrinter.PausePrinting;
begin
  PrintingPaused := True;
end;

procedure TRSPrinter.RestorePrinting;
begin
  PrintingPaused := False;
end;

procedure TRSPrinter.BtnCancelarClick(Sender: TObject);
begin
  Cancelado := True;
end;

end.

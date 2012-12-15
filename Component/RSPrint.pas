unit RSPrint;

interface

uses
  Windows, Messages, SysUtils, CommDlg, Classes, Graphics, Controls, ExtCtrls, StdCtrls, Consts, ShellAPI, Menus,
  Printers, ComCtrls, Forms, Dialogs, RSPrint.CommonTypes;

const
  RSPrintVersion = 'Versão 2.0';

type
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
    FPageSize: TPageSize;
    FPageLength: byte;
    FReportName: String;
    FSaveConfToRegistry: boolean;
    FContinuousJump: byte;
    Cancelado: boolean;
    FPrinterStatus: TPrinterStatus;
    FControlCodes: TControlCodes;

    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    function GetModelRealName(Model : TPrinterModel) : String;
    procedure SetModel(Name: TPrinterModel);
    procedure SetFastPuerto(Puerto: string);
    function GetPaginas: integer;
    procedure Clear;
    procedure BtnCancelarClick(Sender: TObject);
    function IsCurrentlyPrinting: Boolean;

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

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure BeginDoc;  // PARA COMENZAR UNA NUEVA IMPRESION
    procedure EndDoc; // ELIMINA LA INFORMACION DE LAS PAGINAS AL FINALIZAR

    procedure PreviewReal;                            // MUESTRA LA IMPRESION EN PANTALLA
    procedure Print;                           // MANDA LA IMPRESION A LA IMPRESORA
    procedure PrintAll;                           // MANDA LA IMPRESION A LA IMPRESORA
    procedure CancelAllPrinting;
    procedure PausePrinting;
    procedure RestorePrinting;
    procedure CancelPrinting;

    procedure Write(Line, Col: byte; Text: string); // ESCRIBE UN TEXTO EN LA PAGINA
    procedure WriteFont(Line, Col: byte; Text: string; Font: TFastFont); // ESCRIBE UN TEXTO EN LA PAGINA
    procedure DrawMetafile(Col, Line: single; Picture: TMetafile); // Hace un dibujo

    procedure Box(Line1,Col1,Line2,Col2 : byte; Kind : TLineType);                // ESCRIBE UN RECTANGULO
    procedure LineH(Line,Col1,Col2 : byte; Kind : TLineType);           // ESCRIBE UNA LINEA HORIZONTAL
    procedure LineV(Line1,Line2,Col : byte; Kind : TLineType);             // ESCRIBE UNA LINEA VERTICAL

    procedure NewPage;                        // CREA UNA NUEVA PAGINA

    procedure SetModelName(Name : String);
    function GetModelName : String;

    procedure GetModels(Models : TStrings);

    procedure BuildPage(Number : integer; Page : TMetaFile;ToPrint:boolean;Mode : TPrinterMode;ShowDialog : boolean);
    function PrintPage(Number:integer) : boolean;// MANDA UNA PAGINA A LA IMPRESORA
    function GetPrintingWidth : integer;
    function GetPrintingHeight : integer;

    property Lines : byte read FLineas write FLineas default 66; // CANTIDAD DE LINEAS POR PAGINA
    property WinPrinter : string read fWinPrinter write fWinPrinter;
    property WinPort : string read fWinPort write fWinPort;
    property CurrentlyPrinting: Boolean read IsCurrentlyPrinting;

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

  TTbProcedure = procedure of object;

  TTbGetTextEvent = procedure(var Text: string) of object;

procedure Register;

implementation

uses
  RSPrint.Preview, RSPrint.Utils, RSPrint.FastMode, ComObj;

procedure Register;
begin
  RegisterComponents('RS Componentes', [TRSPrinter]);
end;

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
      FControlCodes.Normal := '27 70 27 54 00';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 54 01';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    Cannon_Laser :
    begin
      FControlCodes.Normal := '27 38';
      FControlCodes.Bold := '27 79';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '';
      FControlCodes.UnderlineON := '27 69';
      FControlCodes.UnderlineOFF := '27 82';
      FControlCodes.CondensedON := '27 31 09';
      FControlCodes.CondensedOFF := '27 31 13';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    HP_Deskjet :
    begin
      FControlCodes.Normal := '27 40 115 48 66 27 40 115 48 83';
      FControlCodes.Bold := '27 40 115 51 66';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 40 115 49 83';
      FControlCodes.UnderlineON := '27 38 100 48 68';
      FControlCodes.UnderlineOFF := '27 38 100 64';
      FControlCodes.CondensedON := '27 40 115 49 54 46 54 72';
      FControlCodes.CondensedOFF := '27 40 115 49 48 72';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    HP_Laserjet :
    begin
      FControlCodes.Normal := '27 40 115 48 66 27 40 115 48 83';
      FControlCodes.Bold := '27 40 115 53 66';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 40 115 49 83';
      FControlCodes.UnderlineON := '27 38 100 68';
      FControlCodes.UnderlineOFF := '27 38 100 64';
      FControlCodes.CondensedON := '27 40 115 49 54 46 54 72';
      FControlCodes.CondensedOFF := '27 40 115 49 48 72';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '12';
      FControlCodes.SelLength := '27 67';
    end;

    HP_Thinkjet :
    begin
      FControlCodes.Normal := '27 70';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '';
      FControlCodes.UnderlineON := '27 45 49';
      FControlCodes.UnderlineOFF := '27 45 48';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    IBM_Color_Jet :
    begin
      FControlCodes.Normal := '27 72';
      FControlCodes.Bold := '27 71';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    IBM_PC_Graphics :
    begin
      FControlCodes.Normal := '27 70 27 55';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 54';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    IBM_Proprinter :
    begin
      FControlCodes.Normal := '27 70';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    NEC_3500 :
    begin
      FControlCodes.Normal := '27 72';
      FControlCodes.Bold := '27 71';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '';
      FControlCodes.UnderlineON := '27 45';
      FControlCodes.UnderlineOFF := '27 39';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    NEC_Pinwriter:
    begin
      FControlCodes.Normal := '27 70 27 53';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 52';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '';
      FControlCodes.SelLength := '27 67';
    end;

    Epson_Stylus:
    begin
      FControlCodes.Normal := '27 70 27 53 27 45 00';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 52';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '27 70 15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '27 64';
      FControlCodes.SelLength := '27 67';
    end;

    Mp20Mi:
    begin
      FControlCodes.Normal := '27 77';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 87 01';
      FControlCodes.Italic := '27 52';
      FControlCodes.UnderlineON := '27 45 01';
      FControlCodes.UnderlineOFF := '27 45 00';
      FControlCodes.CondensedON := '27 15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '27 64';
      FControlCodes.SelLength := '27 67';
    end;
    else
    begin
      FControlCodes.Normal := '27 70 27 53';
      FControlCodes.Bold := '27 69';
      FControlCodes.Wide := '27 14';
      FControlCodes.Italic := '27 52';
      FControlCodes.UnderlineON := '27 45 49';
      FControlCodes.UnderlineOFF := '27 45 48';
      FControlCodes.CondensedON := '15';
      FControlCodes.CondensedOFF := '18';
      FControlCodes.Setup := '';
      FControlCodes.Reset := '27 64';
      FControlCodes.SelLength := '27 67';
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
  PreviewForm := TFrmPreview.Create(self);
  TFrmPreview(PreviewForm).RSPrinter := self;
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

function TRSPrinter.IsCurrentlyPrinting: Boolean;
begin
  Result := FPrinterStatus.CurrentlyPrinting;
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
  FastMode: TFastMode;
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
    Job.PageContinuousJump := PageContinuousJump;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.FCopias := Copies;
    Job.FFuente := FastFont;
    Job.FPort := FastPort;
    Job.FLineas := Lines;
    Job.FTransliterate := Transliterate;
    Job.ControlCodes := FControlCodes;

    FastMode := TFastMode.Create(FPrinterStatus);
    try
      FastMode.Print(Job);
    finally
      FastMode.Free;
    end;

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
  FastMode: TFastMode;
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
    Job.PageContinuousJump := PageContinuousJump;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.FCopias := Copies;
    Job.FFuente := FastFont;
    Job.FPort := FastPort;
    Job.FLineas := Lines;
    Job.FTransliterate := Transliterate;
    Job.ControlCodes := FControlCodes;

    FastMode := TFastMode.Create(FPrinterStatus);
    try
      FastMode.Print(Job);
    finally
      FastMode.Free;
    end;
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

procedure TRSPrinter.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if CanClose and FPrinterStatus.CurrentlyPrinting and (MessageDlg('Há páginas a serem impressas. Cancelar? ',mtConfirmation,[mbYes,mbNo],0)=mrNo) then
    CanClose := False;

  if Assigned(OldCloseQuery) then
    OldCloseQuery(Sender,CanClose);
end;

procedure TRSPrinter.CancelAllPrinting;
begin
  FPrinterStatus.CancelAllPrinting;
end;

procedure TRSPrinter.CancelPrinting;
begin
  FPrinterStatus.CancelPrinting;
end;

procedure TRSPrinter.PausePrinting;
begin
  FPrinterStatus.PausePrinting;
end;

procedure TRSPrinter.RestorePrinting;
begin
  FPrinterStatus.PrintingPaused := False;
end;

procedure TRSPrinter.BtnCancelarClick(Sender: TObject);
begin
  Cancelado := True;
end;

end.

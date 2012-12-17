unit RSPrint;

interface

uses
  Windows, Messages, SysUtils, CommDlg, Classes, Graphics, Controls, ExtCtrls, StdCtrls, Consts, ShellAPI, Menus,
  Printers, ComCtrls, Forms, Dialogs, RSPrint.CommonTypes;

const
  RSPrintVersion = 'Versão 2.0';

type
  TRSPrinter = class(TComponent)
  private const
    PREVIEW_FONT_NAME = 'courier new';
    PREVIEW_FONT_SIZE = 18;

  private
    FOwner: TForm;
    FOnPrinterError: TNotifyEvent;
    FModelo: TPrinterModel;           // MODELO DE IMPRESORA
    FLineas: Byte;              // CANTIDAD DE LINEAS POR PAGINA
    FColumnas: Byte;            // CANTIDAD DE COLUMNAS EN LA PAGINA
    FDefaultFont: TFastFont;         // TIPO DE LETRA
    FMode: TPrinterMode;        // MODO DE IMPRESION (NORMAL/MEJORADO)
    LasPaginas: TList;             // Almacenamiento de las páginas
    FCopias: Integer;
    FWinSupMargin: Integer;     // MARGEN SUPERIOR PARA EL MODO WINDOWS
    FWinBotMargin: Integer;     // MARGEN INFERIOR PARA EL MODO WINDOWS
    PaginaActual: Integer;
    FShowPreview: TRSPrinterPreview;
    FZoom: TInitialZoom;
    FTitulo: string;
    FWinPrinter: string;        // NOMBRE DE LA IMPRESORA EN WINDOWS
    FTransliterate: Boolean;
    FPageSize: TPageSize;
    FPageLength: Byte;
    FReportName: string;
    FContinuousJump: Byte;
    Cancelado: Boolean;
    FPrinterStatus: TPrinterStatus;
    FControlCodes: TControlCodes;
    FOldCloseQuery: procedure (Sender: TObject; var CanClose: Boolean) of object;

    procedure FormCloseQuery(sender: TObject; var canClose: Boolean);

    function GetModelRealName(model: TPrinterModel): string;
    procedure SetModel(name: TPrinterModel);

    function GetPaginas: Integer;
    procedure Clear;
    procedure BtnCancelarClick(sender: TObject);
    function IsCurrentlyPrinting: Boolean;

  protected
    procedure Loaded; override;

  public
    PageNo: Integer;
    PageWidth: Integer;         // ANCHO DE PAGINA EN PIXELS
    PageHeight: Integer;  // ALTO DE PAGINA EN PIXELS
    PageWidthP: Double;        // ANCHO DE PAGINA EN PULGADAS
    PageHeightP: Double;  // ALTO DE PAGINA EN PULGADAS
    PageOrientation: TPrinterOrientation; // ORIENTACION DE LA PAGINA
    ReGenerate: Procedure of object;

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure BeginDoc;
    procedure EndDoc;

    procedure PreviewReal;
    procedure Print;
    procedure PrintAll;
    procedure CancelAllPrinting;
    procedure PausePrinting;
    procedure RestorePrinting;
    procedure CancelPrinting;

    procedure Write(line, col: Byte; text: string);
    procedure WriteFont(line, col: Byte; text: string; font: TFastFont);
    procedure NewPage;

    procedure DrawMetafile(col, line: Single; picture: TMetafile);

    procedure Box(line1, col1, line2, col2: Byte; kind: TLineType);
    procedure LineH(line, col1, col2: Byte; kind: TLineType);
    procedure LineV(line1, line2, col: Byte; kind: TLineType);

    procedure SetModelName(name: string);
    function GetModelName: string;
    procedure GetModels(models: TStrings);

    procedure BuildPage(number: Integer; page: TMetaFile; toPrint: Boolean; mode: TPrinterMode; showDialog: Boolean);
    function PrintPage(number: Integer): Boolean;
    function GetPrintingWidth: Integer;
    function GetPrintingHeight: Integer;

    property Lines: Byte read FLineas write FLineas default 66;
    property WinPrinter: string read FWinPrinter write FWinPrinter;
    property CurrentlyPrinting: Boolean read IsCurrentlyPrinting;

  published
    property PageContinuousJump: Byte read FContinuousJump write FContinuousJump default 5;
    property PageSize: TPageSize read FPageSize write FPageSize;
    property PageLength: Byte read FPageLength write FPageLength;
    property FastPrinter: TPrinterModel read FModelo Write SetModel;
    property FastFont: TFastFont read FDefaultFont write FDefaultFont;
    property Mode: TPrinterMode read FMode write FMode default pmFast;
    property Columnas: Byte read FColumnas default 80;
    property Paginas: Integer read GetPaginas default 0;
    property ReportName: string read FReportName write FReportName;
    property Zoom: TInitialZoom read FZoom write FZoom default zWidth;
    property Preview: TRSPrinterPreview read FShowPreview write FShowPreview;
    property Title: string read FTitulo write FTitulo;
    property Copies: Integer read fCopias write fCopias default 1;
    property OnPrinterError: TNotifyEvent read FOnPrinterError write FOnPrinterError;
    property PrintingWidth: Integer read GetPrintingWidth;
    property PrintingHeight: Integer read GetPrintingHeight;
    property Transliterate: Boolean read FTransliterate write FTransliterate default True;
    property WinMarginTop: Integer read FWinSupMargin write FWinSupMargin default 0;
    property WinMarginBottom: Integer read FWinBotMargin write FWinBotMargin default 0;
  end;

  TTbProcedure = procedure of object;

  TTbGetTextEvent = procedure(var text: string) of object;

procedure Register;

implementation

uses
  RSPrint.Preview, RSPrint.Utils, RSPrint.FastMode, ComObj;

procedure Register;
begin
  RegisterComponents('RS', [TRSPrinter]);
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
    FOldCloseQuery := TForm(AOwner).OnCloseQuery;
    TForm(AOwner).OnCloseQuery := FormCloseQuery;
  end;

  SetModel(EPSON_FX);
  FColumnas := 80;
  PaginaActual := 0;
  PageNo:= 1;
  fCopias := 1;
  FZoom := zWidth;
  FDefaultFont := [];
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
    WinPrinter := ADevice;
  end;

  ReGenerate := nil;
  FTransliterate := True;
end;

destructor TRSPrinter.Destroy;
begin
  Clear;

  if not (CsDesigning in ComponentState) then
    FOwner.OnCloseQuery := FOldCloseQuery;

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

procedure TRSPrinter.SetModel(name : TPrinterModel);
var
  Nombre: TPrinterModel;
begin
  Nombre := name;
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

function TRSPrinter.GetPaginas: Integer;
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

procedure TRSPrinter.WriteFont(line, col: Byte; text: string; font: TFastFont);
var
  Txt: PWrite;
  Pag: PPage;
  P: Integer;
begin
  if LasPaginas.Count > 0 then
    begin
      Pag := LasPaginas.Items[PaginaActual];
      P := 0;

      if Pag.Writed.Count > 0 then
      begin
        Txt := Pag.Writed.Items[0];
        while (P<Pag.Writed.Count) and
              ((Txt^.Line < line) or
              ((Txt^.Line = line)and(Txt^.Col<=col))) do
        begin
          Inc(P);
          if P<Pag.Writed.Count then
            Txt := Pag.Writed.Items[P];
        end;
      end;

      New(Txt);
      Pag.Writed.Insert(P,Txt);
      Txt^.Col := col;
      Txt^.Line := line;
      Txt^.Text := text;
      Txt^.Font := font;

      if line > Pag.PrintedLines then
        Pag.PrintedLines := line;
    end;
end;

procedure TRSPrinter.Write(line, col: Byte; text: string);
begin
  WriteFont(line, col, text, FDefaultFont);
end;

procedure TRSPrinter.DrawMetafile(col, line: Single; picture: TMetafile);
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
    Graf^.Col := col;
    Graf^.Line := line;
    Graf^.Picture := Dib;
    Dib.Metafile.Assign(picture);
    Pag^.Graphics.Add(Graf);
  end;
end;

procedure TRSPrinter.Box(line1, col1, line2, col2: Byte; kind: TLineType);
var
  Pag: PPage;
  VLine: PVertLine;
  HLine: PHorizLine;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];

    New(VLine);
    VLine^.Col := col1;
    VLine^.Line1 := line1;
    VLine^.Line2 := line2;
    VLine^.Kind := kind;
    Pag^.VerticalLines.Add(VLine);

    New(VLine);
    VLine^.Col := col2;
    VLine^.Line1 := line1;
    VLine^.Line2 := line2;
    VLine^.Kind := kind;
    Pag^.VerticalLines.Add(VLine);

    New(HLine);
    HLine^.Col1 := col1;
    HLine^.Col2 := col2;
    HLine^.Line := line1;
    HLine^.Kind := kind;
    Pag^.HorizLines.Add(HLine);

    New(HLine);
    HLine^.Col1 := col1;
    HLine^.Col2 := col2;
    HLine^.Line := line2;
    HLine^.Kind := kind;
    Pag^.HorizLines.Add(HLine);

    if line2 > Pag^.PrintedLines then
      Pag^.PrintedLines := line2;
  end;
end;

procedure TRSPrinter.LineH(line, col1, col2: Byte; Kind: TLineType);
var
  HLine: PHorizLine;
  Pag: PPage;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];

    New(HLine);
    Pag.HorizLines.Add(HLine);
    HLine^.Col1 := col1;
    HLine^.Col2 := col2;
    HLine^.Line := line;
    HLine^.Kind := kind;

    if line > Pag^.PrintedLines then
      Pag^.PrintedLines := line;
  end;
end;

procedure TRSPrinter.LineV(line1, line2, col: Byte; kind: TLineType);
var
  VLine: PVertLine;
  Pag: PPage;
begin
  if LasPaginas.Count > 0 then
  begin
    Pag := LasPaginas.Items[PaginaActual];

    New(VLine);
    Pag.VerticalLines.Add(VLine);
    VLine^.Col := col;
    VLine^.Line1 := line1;
    VLine^.Line2 := line2;
    VLine^.Kind := kind;

    if line2 > Pag^.PrintedLines then
      Pag^.PrintedLines := line2;
  end;
end;

procedure TRSPrinter.PreviewReal;
var
  PreviewForm: TForm;
begin
  PreviewForm := TFrmPreview.Create(self);
  try
    TFrmPreview(PreviewForm).RSPrinter := self;
    PreviewForm.ShowModal;
  finally
    PreviewForm.Release;
  end;
end;

procedure TRSPrinter.Print;
begin
  if FShowPreview = ppNo then
    PrintAll
  else
    PreviewReal;
end;

procedure TRSPrinter.BuildPage(number: Integer; page: TMetaFile; toPrint: Boolean; mode: TPrinterMode; showDialog: Boolean);
var
  Pagina: PPage;
  Grafico: PGraphic;
  Escritura: PWrite;
  EA: Integer; // LINEA ACTUAL
  DA: Integer;
  LineaHorizontal: PHorizLine;
  LineaVertical: PVertLine;
  Derecha: Integer;
  Abajo: Integer;
  Texto: TMetafile;
  Ancho: Integer;
  Alto: Integer;
  MargenIzquierdo: Integer;
  MargenSuperior: Integer;
  AltoDeLinea: Double;
  AnchoDeColumna: Double;
  CantColumnas: Integer;
  I: Integer;
  WaitForm: TForm;
  WaitPanel: TPanel;
  BtnCancelar: TButton;
begin
  if toPrint then
    begin
      if DobleWide in FastFont then
        CantColumnas := 42
      else if Compress in FastFont then
        CantColumnas := 136
      else
        CantColumnas := 81;

      With TMetaFileCanvas.Create(page,0) do
        try
          Brush.Color := clWhite;
          AnchoDeColumna := 640/CantColumnas;//(Hoja.Width-(2*MargenIzquierdo))/CantColumnas;
          AltoDeLinea := 12;//Round(AnchoDeColumna/Ancho*Alto);
          page.Width := Round(AnchoDeColumna*CantColumnas);//Round(GetDeviceCaps(Printer.Handle,PHYSICALWIDTH)/GetDeviceCaps(Printer.Handle,LOGPIXELSX)*80);
          page.Height := Round(AltoDeLinea*(Trunc(PageHeightP*6)-2));
          Font.Name := PREVIEW_FONT_NAME;
          Font.Size := PREVIEW_FONT_SIZE;
          Font.Pitch := fpFixed;
          Ancho := TextWidth('X');
          Alto := TextHeight('X');
          FillRect(Rect(0,0,(page.Width)-1,(page.Height-1))); //80*(Ancho),66*Alto));
          Pagina := LasPaginas.Items[number-1];

          if mode = pmWindows then
            begin
              For DA := 0 to Pagina.Graphics.Count-1 do
                begin
                  Grafico := Pagina.Graphics[DA];
                  Draw(Round(AnchoDeColumna*(Grafico^.Col)),Round(FWinSupMargin+AltoDeLinea*(Grafico^.Line)),Grafico^.Picture.Graphic);
                end;
            end;

          for EA := 0 to Pagina.Writed.Count-1 do
            begin
              Escritura := Pagina.Writed.Items[EA];
              if Escritura^.Line <= Lines then
                begin
                  Texto := TMetaFile.Create;
                  Texto.Width := Ancho*Length(Escritura^.Text);
                  Texto.Height := Alto;
                  With TMetaFileCanvas.Create(Texto,0) do
                    try
                      Font.Name := PREVIEW_FONT_NAME;
                      Font.Size := PREVIEW_FONT_SIZE;
                      Font.Pitch := fpFixed;
                      if Bold in Escritura^.Font then
                        Font.Style := Font.Style + [fsBold]
                      else
                        Font.Style := Font.Style - [fsBold];
                      if Italic in Escritura^.Font then
                        Font.Style := Font.Style + [fsItalic]
                      else
                        Font.Style := Font.Style - [fsItalic];
                      if Underline in Escritura^.Font then
                        Font.Style := Font.Style + [fsUnderline]
                      else
                        Font.Style := Font.Style - [fsUnderline];
                      TextOut(0,0,Escritura^.Text);
                    finally
                      Free
                    end;

                  if (DobleWide in Escritura^.Font) and not (DobleWide in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*page.Width/42),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto)
                  else if (Compress in Escritura^.Font) and not (Compress in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*page.Width/140),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto)
                  else
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura^.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line-1))-1,Round(AnchoDeColumna*(Escritura^.Col+Length(Escritura^.Text))),Round(FWinSupMargin+AltoDeLinea*(Escritura^.Line))+1),Texto);
                  Texto.Free;
                end;
              Application.ProcessMessages;
            end;
          for I := 0 to Pagina.VerticalLines.Count-1 do
            begin
              LineaVertical := Pagina.VerticalLines[I];
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
          for I := 0 to Pagina.HorizLines.Count-1 do
            begin
              LineaHorizontal := Pagina.HorizLines[I];
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
      if showDialog then
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
      With TMetaFileCanvas.Create(page,0) do
        try
          Brush.Color := clWhite;

         MargenIzquierdo:=9;
         MargenSuperior:=9;


          AnchoDeColumna := 640/CantColumnas;//(Hoja.Width-(2*MargenIzquierdo))/CantColumnas;
          AltoDeLinea := 12;//Round(AnchoDeColumna/Ancho*Alto);

          page.Width := Round(AnchoDeColumna*CantColumnas)+2*MargenIzquierdo;//Round(GetDeviceCaps(Printer.Handle,PHYSICALWIDTH)/GetDeviceCaps(Printer.Handle,LOGPIXELSX)*80);
          page.Height := Round(AltoDeLinea*(Trunc(PageHeightP*6)-2))+2*MargenSuperior;
          Font.Name := PREVIEW_FONT_NAME;
          Font.Size := PREVIEW_FONT_SIZE;
          Font.Pitch := fpFixed;
          Ancho := TextWidth('X');
          Alto := TextHeight('X');
          FillRect(Rect(0,0,(page.Width)-1,(page.Height)-1)); //80*(Ancho),66*Alto));
          Rectangle(0,0,page.Width,page.Height);
          MargenSuperior := MargenSuperior+FWinSupMargin;
          Pen.Style := psSolid;
          Pen.Color := clBlack;
          Pagina := LasPaginas.Items[number-1];
          if mode = pmWindows then
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
                        Font.Name := PREVIEW_FONT_NAME;
                        Font.Size := PREVIEW_FONT_SIZE;
                        Font.Pitch := fpFixed;
                        if Bold in Escritura^.Font then
                          Font.Style := Font.Style + [fsBold]
                        else
                          Font.Style := Font.Style - [fsBold];
                        if Italic in Escritura^.Font then
                          Font.Style := Font.Style + [fsItalic]
                        else
                          Font.Style := Font.Style - [fsItalic];
                        if Underline in Escritura^.Font then
                          Font.Style := Font.Style + [fsUnderline]
                        else
                          Font.Style := Font.Style - [fsUnderline];
                          TextOut(0,0,Escritura^.Text);
                      finally
                        Free
                      end;
                    if (DobleWide in Escritura^.Font) and not (DobleWide in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*page.Width/42),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto)
                    else if (Compress in Escritura^.Font) and not (Compress in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)+Length(Escritura^.Text)*page.Width/140),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto)
                    else
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura^.Col+Length(Escritura^.Text))),Round(MargenSuperior+AltoDeLinea*(Escritura^.Line))+1),Texto);
                    Texto.Free;
                  end;
                Application.ProcessMessages;
              end;
          for I := 0 to Pagina.VerticalLines.Count-1 do
            if not Cancelado then
              begin
                LineaVertical := Pagina.VerticalLines[I];
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
          for I := 0 to Pagina.HorizLines.Count-1 do
            if not Cancelado then
              begin
                LineaHorizontal := Pagina.HorizLines[I];
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
      if showDialog then
        WaitForm.Free;
    end;
end;

function TRSPrinter.GetPrintingWidth: Integer;
begin
  Result := Printer.PageWidth;
end;

function TRSPrinter.IsCurrentlyPrinting: Boolean;
begin
  Result := FPrinterStatus.CurrentlyPrinting;
end;

function TRSPrinter.GetPrintingHeight: Integer;
begin
  Result := Printer.PageHeight;
end;

function TRSPrinter.PrintPage(number: Integer): Boolean;
var
  Pag: PPage;
  Esc: pWrite;
  Img: TMetaFile;
  Resultado: Boolean;
  LP: TList;
  j: integer;
  LV: pVertLine;
  LH: pHorizLine;
  Job: pPrintJob;
  FastMode: TFastMode;
begin
  if FMode = pmWindows then
  begin
   try
     Printer.Title := Title + ' (Página Nº'+' '+IntToStr(number)+')';
     Printer.Copies := fCopias;
     Printer.BeginDoc;
     Img := TMetaFile.Create;
     Img.Width := 640; // PrintingWidth{PageWidth} div 4; // 640;
     Img.Height := 12*Lines; //PrintingHeight{PageHeight} div 4; // 1056;
     try
       BuildPage(number,Img,True,pmWindows,False);
       Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Img);
       Printer.EndDoc;
     except
     end;
     Img.Free;
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
    Pag^.PrintedLines := PPage(LasPaginas[number-1])^.PrintedLines + 1;
    Pag^.Writed := TList.Create;
    Pag^.VerticalLines := TList.Create;
    Pag^.HorizLines := TList.Create;

    for j := 0 to PPage(LasPaginas[number-1])^.Writed.Count-1 do
    begin
      new(Esc);
      Esc^ := pWrite(PPage(LasPaginas[number-1])^.Writed[j])^;
      Pag.Writed.Add(Esc);
    end;

    for j := 0 to PPage(LasPaginas[number-1])^.VerticalLines.Count-1 do
    begin
      new(LV);
      LV^ := pVertLine(PPage(LasPaginas[number-1])^.VerticalLines[j])^;
      Pag.VerticalLines.Add(LV);
    end;

    for j := 0 to PPage(LasPaginas[number-1])^.HorizLines.Count-1 do
    begin
      new(LH);
      LH^ := pHorizLine(PPage(LasPaginas[number-1])^.HorizLines[j])^;
      Pag.HorizLines.Add(LH);
    end;

    New(Job);
    Job.Name := Title + 'Pág. '+IntToStr(number-1);
    Job.PageSize := PageSize;
    Job.PageContinuousJump := PageContinuousJump;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.Copias := Copies;
    Job.DefaultFont := FastFont;
    Job.Lineas := Lines;
    Job.Transliterate := Transliterate;
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

procedure TRSPrinter.PrintAll;
var
  Pag: PPage;
  Esc: PWrite;
  Imagen: TMetaFile;
  I: Integer;
  J: Integer;
  LH: PHorizLine;
  LV: PVertLine;
  LP: TList;
  Job: PPrintJob;
  FastMode: TFastMode;
begin
  if FMode = pmWindows then
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
      for I := 2 to LasPaginas.Count do
      begin
        Printer.NewPage;
        BuildPage(I,Imagen,True,pmWindows,False);
        Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Imagen);
      end;
    Printer.EndDoc;
    Imagen.Free;
  end
  else
  begin
    LP := TList.Create;

    for I := 0 to LasPaginas.Count-1 do
    begin
      New(Pag);
      LP.Add(Pag);
      Pag^.PrintedLines := PPage(LasPaginas[I])^.PrintedLines + 1;
      Pag^.Writed := TList.Create;
      Pag^.VerticalLines := TList.Create;
      Pag^.HorizLines := TList.Create;

      for J := 0 to PPage(LasPaginas[I])^.Writed.Count-1 do
      begin
        new(Esc);
        Esc^ := pWrite(PPage(LasPaginas[I])^.Writed[J])^;
        Pag.Writed.Add(Esc);
      end;

      for J := 0 to PPage(LasPaginas[I])^.VerticalLines.Count-1 do
      begin
        new(LV);
        LV^ := pVertLine(PPage(LasPaginas[I])^.VerticalLines[J])^;
        Pag.VerticalLines.Add(LV);
      end;

      for J := 0 to PPage(LasPaginas[I])^.HorizLines.Count-1 do
      begin
        new(LH);
        LH^ := pHorizLine(PPage(LasPaginas[I])^.HorizLines[J])^;
        Pag.HorizLines.Add(LH);
      end;
    end;

    New(Job);
    Job.Name := Title;
    Job.PageSize := PageSize;
    Job.PageContinuousJump := PageContinuousJump;
    Job.PageLength := PageLength;
    Job.LasPaginas := LP;
    Job.Copias := Copies;
    Job.DefaultFont := FastFont;
    Job.Lineas := Lines;
    Job.Transliterate := Transliterate;
    Job.ControlCodes := FControlCodes;

    FastMode := TFastMode.Create(FPrinterStatus);
    try
      FastMode.Print(Job);
    finally
      FastMode.Free;
    end;
  end;
end;

procedure TRSPrinter.NewPage;
var
  Pag: PPage;
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

procedure TRSPrinter.SetModelName(name: string);
begin
  if FMode = pmWindows then
  begin
    if Printer.Printers.IndexOf(name) <> -1 then
      Printer.PrinterIndex := Printer.Printers.IndexOf(name);
  end
  else
  begin
    if name = 'Cannon F-60' then
      SetModel(Cannon_F60)
    else if name = 'Cannon laser' then
      SetModel(Cannon_Laser)
    else if name = 'Epson FX/LX/LQ' then
      SetModel(Epson_FX)
    else if name = 'Epson Stylus' then
      SetModel(Epson_Stylus)
    else if name = 'HP Deskjet' then
      SetModel(HP_Deskjet)
    else if name = 'HP Laserjet' then
      SetModel(HP_Laserjet)
    else if name = 'HP Thinkjet' then
      SetModel(HP_Thinkjet)
    else if name = 'IBM Color Jet' then
      SetModel(IBM_Color_Jet)
    else if name = 'IBM PC Graphics' then
      SetModel(IBM_PC_Graphics)
    else if name = 'IBM Proprinter' then
      SetModel(IBM_Proprinter)
    else if name = 'NEC 3500' then
      SetModel(NEC_3500)
    else if name = 'NEC Pinwriter' then
      SetModel(NEC_Pinwriter)
    else if name = 'Mp20Mi' then
      SetModel(Mp20Mi);
  end;
end;

function TRSPrinter.GetModelRealName(model: TPrinterModel): string;
begin
  case model of
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

function TRSPrinter.GetModelName: string;
begin
  Result := GetModelRealName(FModelo);
end;

procedure TRSPrinter.GetModels(models : TStrings);
var
  I: Integer;
begin
  models.Clear;
  for I := Ord(Low(TPrinterModel)) to Ord(High(TPrinterModel)) do
    models.Add(GetModelRealName(TPrinterModel(I)));
end;

procedure TRSPrinter.FormCloseQuery(sender: TObject; var canClose: Boolean);
begin
  if canClose and FPrinterStatus.CurrentlyPrinting and (MessageDlg('Há páginas a serem impressas. Cancelar? ',mtConfirmation,[mbYes,mbNo],0)=mrNo) then
    canClose := False;

  if Assigned(FOldCloseQuery) then
    FOldCloseQuery(sender,canClose);
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

procedure TRSPrinter.BtnCancelarClick(sender: TObject);
begin
  Cancelado := True;
end;

end.

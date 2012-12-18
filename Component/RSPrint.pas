unit RSPrint;

interface

uses
  Windows, Messages, SysUtils, CommDlg, Classes, Graphics, Controls, ExtCtrls, StdCtrls, Consts, ShellAPI, Menus,
  Printers, ComCtrls, Forms, Dialogs, RSPrint.CommonTypes, RSPrint.Types.Document, Generics.Collections;

type
  TRSPrinter = class(TComponent)
  private const
    PREVIEW_FONT_NAME = 'courier new';
    PREVIEW_FONT_SIZE = 18;

  private
    FDocument: TDocument;

    FOwner: TForm;
    FOnPrinterError: TNotifyEvent;
    FModelo: TPrinterModel;           // MODELO DE IMPRESORA
    FColumnas: Byte;            // CANTIDAD DE COLUMNAS EN LA PAGINA
    FMode: TPrinterMode;        // MODO DE IMPRESION (NORMAL/MEJORADO)
    FWinSupMargin: Integer;     // MARGEN SUPERIOR PARA EL MODO WINDOWS
    FWinBotMargin: Integer;     // MARGEN INFERIOR PARA EL MODO WINDOWS
    FShowPreview: TRSPrinterPreview;
    FZoom: TInitialZoom;
    FWinPrinter: string;        // NOMBRE DE LA IMPRESORA EN WINDOWS
    Cancelado: Boolean;

    function GetModelRealName(model: TPrinterModel): string;
    procedure SetModel(name: TPrinterModel);

    function GetPaginas: Integer;
    procedure BtnCancelarClick(sender: TObject);

    function PrintPageInFastMode(pageNumber: Integer): Boolean;
    function PrintPageInWindowsMode(number: Integer): Boolean;

    procedure PrintAllInFastMode;
    procedure PrintAllInWindowsMode;

    function GetTitle: string;
    procedure SetTitle(const value: string);
    function GetPageSize: TPageSize;
    procedure SetPageSize(const value: TPageSize);
    function GetPageContinuousJump: Byte;
    procedure SetPageContinuousJump(const value: Byte);
    function GetPageLength: Byte;
    procedure SetPageLength(const value: Byte);
    function GetCopies: Integer;
    procedure SetCopies(const value: Integer);
    function GetDefaultFont: TFastFont;
    procedure SetDefaultFont(const value: TFastFont);
    function GetLines: Byte;
    procedure SetLines(const value: Byte);
    function GetTransliterate: Boolean;
    procedure SetTransliterate(const value: Boolean);
    function GetCurrentPageNumber: Integer;

  protected
    procedure Loaded; override;

  public
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

    property Lines: Byte read GetLines write SetLines default 66;
    property WinPrinter: string read FWinPrinter write FWinPrinter;
    property PageNo: Integer read GetCurrentPageNumber;

  published
    property PageContinuousJump: Byte read GetPageContinuousJump write SetPageContinuousJump default 5;
    property PageSize: TPageSize read GetPageSize write SetPageSize;
    property PageLength: Byte read GetPageLength write SetPageLength;
    property FastPrinter: TPrinterModel read FModelo Write SetModel;
    property FastFont: TFastFont read GetDefaultFont write SetDefaultFont;
    property Mode: TPrinterMode read FMode write FMode default pmFast;
    property Columnas: Byte read FColumnas default 80;
    property Paginas: Integer read GetPaginas default 0;
    property Zoom: TInitialZoom read FZoom write FZoom default zWidth;
    property Preview: TRSPrinterPreview read FShowPreview write FShowPreview;
    property Title: string read GetTitle write SetTitle;
    property Copies: Integer read GetCopies write SetCopies default 1;
    property OnPrinterError: TNotifyEvent read FOnPrinterError write FOnPrinterError;
    property PrintingWidth: Integer read GetPrintingWidth;
    property PrintingHeight: Integer read GetPrintingHeight;
    property Transliterate: Boolean read GetTransliterate write SetTransliterate default True;
    property WinMarginTop: Integer read FWinSupMargin write FWinSupMargin default 0;
    property WinMarginBottom: Integer read FWinBotMargin write FWinBotMargin default 0;
  end;

procedure Register;

implementation

uses
  RSPrint.Preview, RSPrint.Utils, RSPrint.FastMode, RSPrint.Types.Img, RSPrint.Types.Page, ComObj;

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

  if (AOwner is TForm) and not (csDesigning in ComponentState) then
    FOwner := TForm(AOwner);

  FColumnas := 80;
  FZoom := zWidth;

  FDocument := TDocument.Create;
  FDocument.Copies := 1;
  FDocument.PageSize := pzLegal;
  FDocument.PageContinuousJump := 5;
  FDocument.DefaultFont := [];
  FDocument.Transliterate := True;

  SetModel(EPSON_FX);

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
end;

destructor TRSPrinter.Destroy;
begin
  FDocument.Free;

  inherited Destroy;
end;

procedure TRSPrinter.SetCopies(const value: Integer);
begin
  FDocument.Copies := value;
end;

procedure TRSPrinter.SetDefaultFont(const value: TFastFont);
begin
  FDocument.DefaultFont := value;
end;

procedure TRSPrinter.SetLines(const value: Byte);
begin
  FDocument.LinesPerPage := value;
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
      FDocument.ControlCodes.Normal := '27 70 27 54 00';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 54 01';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    Cannon_Laser :
    begin
      FDocument.ControlCodes.Normal := '27 38';
      FDocument.ControlCodes.Bold := '27 79';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '';
      FDocument.ControlCodes.UnderlineON := '27 69';
      FDocument.ControlCodes.UnderlineOFF := '27 82';
      FDocument.ControlCodes.CondensedON := '27 31 09';
      FDocument.ControlCodes.CondensedOFF := '27 31 13';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    HP_Deskjet :
    begin
      FDocument.ControlCodes.Normal := '27 40 115 48 66 27 40 115 48 83';
      FDocument.ControlCodes.Bold := '27 40 115 51 66';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 40 115 49 83';
      FDocument.ControlCodes.UnderlineON := '27 38 100 48 68';
      FDocument.ControlCodes.UnderlineOFF := '27 38 100 64';
      FDocument.ControlCodes.CondensedON := '27 40 115 49 54 46 54 72';
      FDocument.ControlCodes.CondensedOFF := '27 40 115 49 48 72';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    HP_Laserjet :
    begin
      FDocument.ControlCodes.Normal := '27 40 115 48 66 27 40 115 48 83';
      FDocument.ControlCodes.Bold := '27 40 115 53 66';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 40 115 49 83';
      FDocument.ControlCodes.UnderlineON := '27 38 100 68';
      FDocument.ControlCodes.UnderlineOFF := '27 38 100 64';
      FDocument.ControlCodes.CondensedON := '27 40 115 49 54 46 54 72';
      FDocument.ControlCodes.CondensedOFF := '27 40 115 49 48 72';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '12';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    HP_Thinkjet :
    begin
      FDocument.ControlCodes.Normal := '27 70';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '';
      FDocument.ControlCodes.UnderlineON := '27 45 49';
      FDocument.ControlCodes.UnderlineOFF := '27 45 48';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    IBM_Color_Jet :
    begin
      FDocument.ControlCodes.Normal := '27 72';
      FDocument.ControlCodes.Bold := '27 71';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    IBM_PC_Graphics :
    begin
      FDocument.ControlCodes.Normal := '27 70 27 55';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 54';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    IBM_Proprinter :
    begin
      FDocument.ControlCodes.Normal := '27 70';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    NEC_3500 :
    begin
      FDocument.ControlCodes.Normal := '27 72';
      FDocument.ControlCodes.Bold := '27 71';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '';
      FDocument.ControlCodes.UnderlineON := '27 45';
      FDocument.ControlCodes.UnderlineOFF := '27 39';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    NEC_Pinwriter:
    begin
      FDocument.ControlCodes.Normal := '27 70 27 53';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 52';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    Epson_Stylus:
    begin
      FDocument.ControlCodes.Normal := '27 70 27 53 27 45 00';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 52';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '27 70 15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '27 64';
      FDocument.ControlCodes.SelLength := '27 67';
    end;

    Mp20Mi:
    begin
      FDocument.ControlCodes.Normal := '27 77';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 87 01';
      FDocument.ControlCodes.Italic := '27 52';
      FDocument.ControlCodes.UnderlineON := '27 45 01';
      FDocument.ControlCodes.UnderlineOFF := '27 45 00';
      FDocument.ControlCodes.CondensedON := '27 15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '27 64';
      FDocument.ControlCodes.SelLength := '27 67';
    end;
    else
    begin
      FDocument.ControlCodes.Normal := '27 70 27 53';
      FDocument.ControlCodes.Bold := '27 69';
      FDocument.ControlCodes.Wide := '27 14';
      FDocument.ControlCodes.Italic := '27 52';
      FDocument.ControlCodes.UnderlineON := '27 45 49';
      FDocument.ControlCodes.UnderlineOFF := '27 45 48';
      FDocument.ControlCodes.CondensedON := '15';
      FDocument.ControlCodes.CondensedOFF := '18';
      FDocument.ControlCodes.Setup := '';
      FDocument.ControlCodes.Reset := '27 64';
      FDocument.ControlCodes.SelLength := '27 67';
    end;
  end;
end;

function TRSPrinter.GetPageContinuousJump: Byte;
begin
  Result := FDocument.PageContinuousJump;
end;

function TRSPrinter.GetPageLength: Byte;
begin
  Result := FDocument.PageLength;
end;

function TRSPrinter.GetPageSize: TPageSize;
begin
  Result := FDocument.PageSize;
end;

function TRSPrinter.GetPaginas: Integer;
begin
  Result := FDocument.Pages.Count;
end;

procedure TRSPrinter.BeginDoc;  // PARA COMENZAR UNA NUEVA IMPRESION
begin
  FDocument.ClearPages;
  NewPage;
end;

procedure TRSPrinter.EndDoc; // ELIMINA LA INFORMACION DE LAS PAGINAS AL FINALIZAR
begin
  FDocument.ClearPages;
end;

procedure TRSPrinter.WriteFont(line, col: Byte; text: string; font: TFastFont);
var
  Escritura: TWrittenText;
  P: Integer;
begin
  if FDocument.Pages.Count > 0 then
    begin
      P := 0;

      if FDocument.CurrentPage.WrittenText.Count > 0 then
      begin
        Escritura := FDocument.CurrentPage.WrittenText.Items[0];
        while (P<FDocument.CurrentPage.WrittenText.Count) and
              ((Escritura.Line < line) or
              ((Escritura.Line = line)and(Escritura.Col<=col))) do
        begin
          Inc(P);
          if P<FDocument.CurrentPage.WrittenText.Count then
            Escritura := FDocument.CurrentPage.WrittenText.Items[P];
        end;
      end;

      Escritura.Col := col;
      Escritura.Line := line;
      Escritura.Text := text;
      Escritura.Font := font;
      FDocument.CurrentPage.WrittenText.Insert(P, Escritura);

      if line > FDocument.CurrentPage.PrintedLines then
        FDocument.CurrentPage.PrintedLines := line;
    end;
end;

procedure TRSPrinter.Write(line, col: Byte; text: string);
begin
  WriteFont(line, col, text, FDocument.DefaultFont);
end;

procedure TRSPrinter.DrawMetafile(col, line: Single; picture: TMetafile);
var
  Graf: TImg;
begin
  if FDocument.Pages.Count > 0 then
  begin
    Graf := TImg.Create;
    Graf.Col := col;
    Graf.Line := line;
    Graf.Picture.Metafile.Assign(picture);

    FDocument.CurrentPage.Images.Add(Graf);
  end;
end;

procedure TRSPrinter.Box(line1, col1, line2, col2: Byte; kind: TLineType);
var
  VLine: TVerticalLine;
  HLine: THorizontalLine;
begin
  if FDocument.Pages.Count > 0 then
  begin

    VLine.Col := col1;
    VLine.Line1 := line1;
    VLine.Line2 := line2;
    VLine.Kind := kind;
    FDocument.CurrentPage.VerticalLines.Add(VLine);

    VLine.Col := col2;
    VLine.Line1 := line1;
    VLine.Line2 := line2;
    VLine.Kind := kind;
    FDocument.CurrentPage.VerticalLines.Add(VLine);

    HLine.Col1 := col1;
    HLine.Col2 := col2;
    HLine.Line := line1;
    HLine.Kind := kind;
    FDocument.CurrentPage.HorizontalLines.Add(HLine);

    HLine.Col1 := col1;
    HLine.Col2 := col2;
    HLine.Line := line2;
    HLine.Kind := kind;
    FDocument.CurrentPage.HorizontalLines.Add(HLine);

    if line2 > FDocument.CurrentPage.PrintedLines then
      FDocument.CurrentPage.PrintedLines := line2;
  end;
end;

procedure TRSPrinter.LineH(line, col1, col2: Byte; Kind: TLineType);
var
  HLine: THorizontalLine;
begin
  if FDocument.Pages.Count > 0 then
  begin
    HLine.Col1 := col1;
    HLine.Col2 := col2;
    HLine.Line := line;
    HLine.Kind := kind;
    FDocument.CurrentPage.HorizontalLines.Add(HLine);

    if line > FDocument.CurrentPage.PrintedLines then
      FDocument.CurrentPage.PrintedLines := line;
  end;
end;

procedure TRSPrinter.LineV(line1, line2, col: Byte; kind: TLineType);
var
  VLine: TVerticalLine;
begin
  if FDocument.Pages.Count > 0 then
  begin
    VLine.Col := col;
    VLine.Line1 := line1;
    VLine.Line2 := line2;
    VLine.Kind := kind;
    FDocument.CurrentPage.VerticalLines.Add(VLine);

    if line2 > FDocument.CurrentPage.PrintedLines then
      FDocument.CurrentPage.PrintedLines := line2;
  end;
end;

procedure TRSPrinter.Loaded;
begin
  inherited;

  //What is the meaning of the line above?
  FDocument.LinesPerPage := FDocument.LinesPerPage - (WinMarginTop div 12) - (WinMarginBottom div 12);
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
  Pagina: TPage;
  Grafico: TImg;
  Escritura: TWrittenText;
  EA: Integer; // LINEA ACTUAL
  DA: Integer;
  LineaHorizontal: THorizontalLine;
  LineaVertical: TVerticalLine;
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
          Pagina := FDocument.Pages.Items[number-1];

          if mode = pmWindows then
            begin
              For DA := 0 to Pagina.Images.Count-1 do
                begin
                  Grafico := Pagina.Images[DA];
                  Draw(Round(AnchoDeColumna*(Grafico.Col)),Round(FWinSupMargin+AltoDeLinea*(Grafico.Line)),Grafico.Picture.Graphic);
                end;
            end;

          for EA := 0 to Pagina.WrittenText.Count-1 do
            begin
              Escritura := Pagina.WrittenText.Items[EA];
              if Escritura.Line <= Lines then
                begin
                  Texto := TMetaFile.Create;
                  Texto.Width := Ancho*Length(Escritura.Text);
                  Texto.Height := Alto;
                  With TMetaFileCanvas.Create(Texto,0) do
                    try
                      Font.Name := PREVIEW_FONT_NAME;
                      Font.Size := PREVIEW_FONT_SIZE;
                      Font.Pitch := fpFixed;
                      if Bold in Escritura.Font then
                        Font.Style := Font.Style + [fsBold]
                      else
                        Font.Style := Font.Style - [fsBold];
                      if Italic in Escritura.Font then
                        Font.Style := Font.Style + [fsItalic]
                      else
                        Font.Style := Font.Style - [fsItalic];
                      if Underline in Escritura.Font then
                        Font.Style := Font.Style + [fsUnderline]
                      else
                        Font.Style := Font.Style - [fsUnderline];
                      TextOut(0, 0, Escritura.Text);
                    finally
                      Free
                    end;

                  if (DobleWide in Escritura.Font) and not (DobleWide in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line-1))-1,Round(AnchoDeColumna*(Escritura.Col)+Length(Escritura.Text)*page.Width/42),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line))+1),Texto)
                  else if (Compress in Escritura.Font) and not (Compress in FastFont) then
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line-1))-1,Round(AnchoDeColumna*(Escritura.Col)+Length(Escritura.Text)*page.Width/140),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line))+1),Texto)
                  else
                    StretchDraw(rect(Round(AnchoDeColumna*(Escritura.Col)),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line-1))-1,Round(AnchoDeColumna*(Escritura.Col+Length(Escritura.Text))),Round(FWinSupMargin+AltoDeLinea*(Escritura.Line))+1),Texto);
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
                      MoveTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                      if LineaVertical.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                    end
                  else // Double;
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                      if LineaVertical.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      MoveTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                             Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                      if LineaVertical.Line2 <= Lines then
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                               Round(FWinSupMargin+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                               Round(FWinSupMargin+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                    end;
                end;
              Application.ProcessMessages;
            end;
          for I := 0 to Pagina.HorizontalLines.Count-1 do
            begin
              LineaHorizontal := Pagina.HorizontalLines[I];
              if LineaHorizontal.Line <= Lines then
                begin
                  If LineaHorizontal.Kind = ltSingle then
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2));
                      if LineaHorizontal.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2))
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2));
                    end
                  else // Double;
                    begin
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1);
                      if LineaHorizontal.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1)
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1);
                      MoveTo(Round(AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                             Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1);
                      if LineaHorizontal.Col2 <= CantColumnas then
                        LineTo(Round(AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1)
                      else
                        LineTo(Round(AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                               Round(FWinSupMargin+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1);
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
          Pagina := FDocument.Pages.Items[number-1];
          if mode = pmWindows then
            begin
              For DA := 0 to Pagina.Images.Count-1 do
                if not Cancelado then
                  begin
                    Grafico := Pagina.Images[DA];
                    Draw(Round(MargenIzquierdo+AnchoDeColumna*(Grafico.Col)),Round(MargenSuperior+AltoDeLinea*(Grafico.Line))+1,Grafico.Picture.Graphic);
                  end;
            end;
          For EA := 0 to Pagina.WrittenText.Count-1 do
            if not Cancelado then
              begin
                Escritura := Pagina.WrittenText.Items[EA];
                if Escritura.Line <= Lines then
                  begin
                    Texto := TMetaFile.Create;
                    Texto.Width := Ancho*Length(Escritura.Text);
                    Texto.Height := Alto;
                    With TMetaFileCanvas.Create(Texto,0) do
                      try
                        Font.Name := PREVIEW_FONT_NAME;
                        Font.Size := PREVIEW_FONT_SIZE;
                        Font.Pitch := fpFixed;
                        if Bold in Escritura.Font then
                          Font.Style := Font.Style + [fsBold]
                        else
                          Font.Style := Font.Style - [fsBold];
                        if Italic in Escritura.Font then
                          Font.Style := Font.Style + [fsItalic]
                        else
                          Font.Style := Font.Style - [fsItalic];
                        if Underline in Escritura.Font then
                          Font.Style := Font.Style + [fsUnderline]
                        else
                          Font.Style := Font.Style - [fsUnderline];
                          TextOut(0,0,Escritura.Text);
                      finally
                        Free
                      end;
                    if (DobleWide in Escritura.Font) and not (DobleWide in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col)+Length(Escritura.Text)*page.Width/42),Round(MargenSuperior+AltoDeLinea*(Escritura.Line))+1),Texto)
                    else if (Compress in Escritura.Font) and not (Compress in FastFont) then
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col)+Length(Escritura.Text)*page.Width/140),Round(MargenSuperior+AltoDeLinea*(Escritura.Line))+1),Texto)
                    else
                      StretchDraw(rect(Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col)),Round(MargenSuperior+AltoDeLinea*(Escritura.Line-1))-1,Round(MargenIzquierdo+AnchoDeColumna*(Escritura.Col+Length(Escritura.Text))),Round(MargenSuperior+AltoDeLinea*(Escritura.Line))+1),Texto);
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
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                        if LineaVertical.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      end
                    else // Double;
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                        if LineaVertical.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2-1),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                               Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line1-1)+AltoDeLinea/2));
                        if LineaVertical.Line2 <= Lines then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                                 Round(MargenSuperior+AltoDeLinea*(LineaVertical.Line2-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaVertical.Col)+AnchoDeColumna/2+1),
                                 Round(MargenSuperior+AltoDeLinea*(Lines-1)+AltoDeLinea/2));
                      end;
                      (*   *)

                  end;
                Application.ProcessMessages;
              end;
          for I := 0 to Pagina.HorizontalLines.Count-1 do
            if not Cancelado then
              begin
                LineaHorizontal := Pagina.HorizontalLines[I];
                if LineaHorizontal.Line <= Lines then
                  begin
                    If LineaHorizontal.Kind = ltSingle then
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2));
                        if LineaHorizontal.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2))
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2));
                      end
                    else // Double;
                      begin
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1);
                        if LineaHorizontal.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1)
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)-1);
                        MoveTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col1)+AnchoDeColumna/2),
                               Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1);
                        if LineaHorizontal.Col2 <= CantColumnas then
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(LineaHorizontal.Col2)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1)
                        else
                          LineTo(Round(MargenIzquierdo+AnchoDeColumna*(CantColumnas)+AnchoDeColumna/2),
                                 Round(MargenSuperior+AltoDeLinea*(LineaHorizontal.Line-1)+AltoDeLinea/2)+1);
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

function TRSPrinter.GetTitle: string;
begin
  Result := FDocument.Title;
end;

function TRSPrinter.GetTransliterate: Boolean;
begin
  Result := FDocument.Transliterate;
end;

function TRSPrinter.GetPrintingHeight: Integer;
begin
  Result := Printer.PageHeight;
end;

function TRSPrinter.PrintPage(number: Integer): Boolean;
begin
  if FMode = pmWindows then
    Result := PrintPageInWindowsMode(number)
  else
    Result := PrintPageInFastMode(number);
end;

function TRSPrinter.PrintPageInWindowsMode(number: Integer): Boolean;
var
  Img: TMetaFile;
begin
  try
    Printer.Title := FDocument.Title + ' (Página Nº'+' '+IntToStr(number)+')';
    Printer.Copies := FDocument.Copies;
    Printer.BeginDoc;
    Img := TMetaFile.Create;
    Img.Width := 640; // PrintingWidth{PageWidth} div 4; // 640;
    Img.Height := 12*Lines; //PrintingHeight{PageHeight} div 4; // 1056;
    try
      BuildPage(number, Img, True, pmWindows, False);
      Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Img);
      Printer.EndDoc;
    except
    end;
    Img.Free;
    Result := True;
  except
    Result := False;
  end;
end;

function TRSPrinter.PrintPageInFastMode(pageNumber: Integer): Boolean;
var
  FastMode: TFastMode;
begin
  FastMode := TFastMode.Create;
  try
    if pageNumber = TFastMode.ALL_PAGES then
      FastMode.Print(FDocument, TFastMode.ALL_PAGES)
    else
      FastMode.Print(FDocument, pageNumber);
  finally
    FastMode.Free;
  end;

  Result := True;
end;

procedure TRSPrinter.PrintAll;
begin
  if FMode = pmWindows then
    PrintAllInWindowsMode
  else
    PrintAllInFastMode;
end;

procedure TRSPrinter.PrintAllInWindowsMode;
var
  Img: TMetaFile;
  I: Integer;
begin
  Printer.Title := Title;
  Printer.Copies := FDocument.Copies;
  Printer.BeginDoc;
  Img := TMetaFile.Create;
  Img.Width := 640; //PrintingWidth{PageWidth} div 4; // 640;
  Img.Height := 12 * Lines; //PrintingHeight{PageHeight} div 4; // 1056;
  BuildPage(1,Img,True,pmWindows,False);
  Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Img);
  if FDocument.Pages.Count > 1 then
    for I := 2 to FDocument.Pages.Count do
    begin
      Printer.NewPage;
      BuildPage(I,Img,True,pmWindows,False);
      Printer.Canvas.StretchDraw(Rect(0,0,Printer.PageWidth,Printer.PageHeight),Img);
    end;
  Printer.EndDoc;
  Img.Free;
end;

procedure TRSPrinter.PrintAllInFastMode;
begin
  PrintPageInFastMode(TFastMode.ALL_PAGES);
end;

procedure TRSPrinter.NewPage;
begin
  FDocument.NewPage;
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

procedure TRSPrinter.SetPageContinuousJump(const value: Byte);
begin
  FDocument.PageContinuousJump := value;
end;

procedure TRSPrinter.SetPageLength(const value: Byte);
begin
  FDocument.PageLength := value;
end;

procedure TRSPrinter.SetPageSize(const value: TPageSize);
begin
  FDocument.PageSize := value;
end;

procedure TRSPrinter.SetTitle(const value: string);
begin
  FDocument.Title := value;
end;

procedure TRSPrinter.SetTransliterate(const value: Boolean);
begin
  FDocument.Transliterate := value;
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

function TRSPrinter.GetCopies: Integer;
begin
  Result := FDocument.Copies;
end;

function TRSPrinter.GetCurrentPageNumber: Integer;
begin
  Result := FDocument.CurrentPageNumber;
end;

function TRSPrinter.GetDefaultFont: TFastFont;
begin
  Result := FDocument.DefaultFont;
end;

function TRSPrinter.GetLines: Byte;
begin
  Result := FDocument.LinesPerPage;
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

procedure TRSPrinter.BtnCancelarClick(sender: TObject);
begin
  Cancelado := True;
end;

end.

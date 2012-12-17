unit RSPrint.Preview;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, ExtCtrls, Spin, Printers, Zlib,
  Buttons, ActnList, ComCtrls;

type
  TFrmPreview = class(TForm)
    Scroller: TScrollBox;
    Shower: TPaintBox;
    PrinterPanelH: TPanel;
    LCopiasV: TLabel;
    CopiasV: TSpinEdit;
    BtnPagPriV: TSpeedButton;
    BtnPagAntV: TSpeedButton;
    BtnPageDirect: TSpeedButton;
    BtnPagSigV: TSpeedButton;
    BtnPagUltV: TSpeedButton;
    ZoomShowerV: TSpinEdit;
    LZoomV: TLabel;
    BtnAnchoV: TSpeedButton;
    BtnCompletaV: TSpeedButton;
    BtnImpActualV: TSpeedButton;
    BtnImpTodoV: TSpeedButton;
    BtnPropiedadesV: TSpeedButton;
    Bevel1: TBevel;
    SpeedButton2: TSpeedButton;
    Bevel2: TBevel;
    Bevel3: TBevel;
    ActionList1: TActionList;
    ActUp: TAction;
    ActDown: TAction;
    ActFechar: TAction;
    ActRight: TAction;
    ActLeft: TAction;
    ActHome: TAction;
    ActEnd: TAction;
    ActF1: TAction;
    ActPgUp: TAction;
    ActPgDn: TAction;
    ActEnter: TAction;
    StatusBar1: TStatusBar;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CopiasVChange(Sender: TObject);
    procedure ZoomShowerVChange(Sender: TObject);
    procedure ShowerPaint(Sender: TObject);
    procedure BtnAnchoVClick(Sender: TObject);
    procedure BtnCompletaVClick(Sender: TObject);
    procedure BtnImpActualVClick(Sender: TObject);
    procedure BtnImpTodoVClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ShowerMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BtnPropiedadesVClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure LTotPagHClick(Sender: TObject);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure ActUpExecute(Sender: TObject);
    procedure ActDownExecute(Sender: TObject);
    procedure ActFecharExecute(Sender: TObject);
    procedure ActRightExecute(Sender: TObject);
    procedure ActLeftExecute(Sender: TObject);
    procedure ActHomeExecute(Sender: TObject);
    procedure ActEndExecute(Sender: TObject);
    procedure ActF1Execute(Sender: TObject);
    procedure ActPgUpExecute(Sender: TObject);
    procedure ActPgDnExecute(Sender: TObject);
    procedure ActEnterExecute(Sender: TObject);

  private
    FPagina : integer;
    FZoom : Double;
    DontDraw : Boolean;
    HojasDib,
    HojasNum : TList;
    procedure GetHoja(Numero : integer; Hoja : TMetaFile; ViewDialog: boolean);
    procedure UpdateStatus(Pag_Atual,Total_paginas:integer);
    procedure PageChange;

  public
    RSPrinter : TComponent;
    Hoja : TMetaFile;

  end;

  function WOpenPrinter(pPrinterName: PChar; var phPrinter: THandle; pDefault: pointer):BOOL; stdcall;
  function WClosePrinter(hPrinter: THandle): BOOL; stdcall;
  function WGetPrinter(hPrinter: THandle; Level: DWORD; pPrinter: Pointer; cbBuf: DWORD; pcbNeeded: PDWORD): BOOL; stdcall;
  function WSetPrinter(hPrinter: THandle; Level: DWORD; pPrinter: Pointer; Command: DWORD): BOOL; stdcall;
  function WDocumentProperties(hWnd: HWND; hPrinter: THandle; pDeviceName: PChar;
    const pDevModeOutput: TDeviceMode; var pDevModeInput: TDeviceMode;
    fMode: DWORD): Longint; stdcall;

var
  FrmPreview: TFrmPreview;

implementation

{$R *.DFM}

uses
  RSPrint, RSPrint.Preview.PrinterSetup, RSPrint.CommonTypes, ZLibConst;

const
  winspl  = 'winspool.drv';

function WOpenPrinter; external winspl name 'OpenPrinterA';
function WClosePrinter; external winspl name 'ClosePrinter';
function WGetPrinter; external winspl name 'GetPrinterA';
function WSetPrinter; external winspl name 'SetPrinterA';
function WDocumentProperties; external winspl name 'DocumentPropertiesA';


procedure tFrmPreview.UpdateStatus(Pag_Atual,Total_paginas:integer);
var
  Modo : string;
begin
  if TRSPrinter(RSPrinter).Mode = pmfast then
    Modo:= 'Matricial'
  else
    Modo := 'Gráfico';

  StatusBar1.Panels[0].Text := 'Página: '+IntToStr(Pag_atual)+' de '+IntToStr(total_paginas);
  StatusBar1.Panels[1].Text := 'Zoom: '+FloatToStr(FZoom);
  StatusBar1.Panels[2].Text := 'Modo: '+ Modo;
  StatusBar1.Panels[3].Text := '  Impressora: '+Printer.Printers[Printer.PrinterIndex];

  Caption:='Preview de Impressão      Página: '+IntToStr(Pag_atual)+
  ' de '+IntToStr(total_paginas);

  if FPagina = 1 then
  begin
    BtnPagPriV.Enabled := False;
    BtnPagAntV.Enabled := False;
  end
  else
  begin
    BtnPagPriV.Enabled := True;
    BtnPagAntV.Enabled := True;
  end;

  if TRSPrinter(RSPrinter).Paginas > FPagina then
  begin
    BtnPagSigV.Enabled := True;
    BtnPagUltV.Enabled := True;
  end
  else
  begin
    BtnPagSigV.Enabled := False;
    BtnPagUltV.Enabled := False;
  end;

  if TRSPrinter(RSPrinter).Paginas = 1 then
  begin
    BtnPagSigV.Enabled := False;
    BtnPagUltV.Enabled := False;
    BtnPageDirect.Enabled := False;
    BtnImpActualV.Enabled := False;
  end;
end;

procedure TFrmPreview.FormShow(Sender: TObject);
var
  ZoomAncho: Double;
  ZoomAlto: Double;
begin
  FPagina := 1;
  Hoja := TMetaFile.Create;
  Hoja.Width := Round(TRSPrinter(RSPrinter).PageWidthP * 80);
  Hoja.Height := Round(TRSPrinter(RSPrinter).PageHeightP * 80);
  Shower.Width := Shower.Canvas.TextWidth('X')*80+20;
  Shower.Height := Shower.Canvas.TextHeight('X')*66+20;

  FPagina := 1;
  UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);
  GetHoja(FPagina,Hoja,True);

  if TRSPrinter(RSPrinter).Zoom = zWidth then
    ZoomShowerV.Value := Round((Screen.Width-30)*100/(Hoja.Width))
  else // a lo alto
  begin
      ZoomAncho := (Screen.Width-30)*100/(Hoja.Width);
      ZoomAlto := (Scroller.ClientHeight)*111/(Hoja.Height);
      if ZoomAncho < ZoomAlto then
        ZoomShowerV.Value := Round(ZoomAncho)
      else
        ZoomShowerV.Value := Round(ZoomAlto);
  end;

  CopiasV.Value := TRSPrinter(RSPrinter).Copies;
end;

procedure TFrmPreview.FormCreate(Sender: TObject);
begin
  DontDraw := True;
  Width := Screen.Width - 30;
  Height := Screen.Height - 30;
  Left := 0;
  Top := 0;
  Screen.Cursors[1] := LoadCursor(hInstance,'TBZOOM');
  WindowState := wsMaximized;
  HojasDib := TList.Create;
  HojasNum := TList.Create;
end;

procedure TFrmPreview.FormDestroy(Sender: TObject);
begin
  Hoja.Free;
  HojasNum.Free;

  while HojasDib.Count > 0 do
  begin
    TMemoryStream(HojasDib[0]).Free;
    HojasDib.Delete(0);
  end;
end;

procedure TFrmPreview.CopiasVChange(Sender: TObject);
begin
  TRSPrinter(RSPrinter).Copies := CopiasV.Value;
end;

procedure TFrmPreview.ZoomShowerVChange(Sender: TObject);
begin
  FZoom := ZoomShowerV.Value;
  DontDraw := true;
  ShowerPaint(self);
  UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);
end;

procedure TFrmPreview.ShowerPaint(Sender: TObject);
var
  Rect: TRect;
begin
  if DontDraw then
  begin
    Shower.Width := Round(Hoja.Width*FZoom/100)+10;
    Shower.Height := Round(Hoja.Height*FZoom/100)+10;
    DontDraw := False;
    Exit;
  end;

  Rect.Left := 10;
  Rect.Top := 10;
  Rect.Right := Round(Hoja.Width*FZoom/100);
  Rect.Bottom := Round(Hoja.Height*FZoom/100);
  Shower.Width := Round(Hoja.Width*FZoom/100)+10;
  Shower.Height := Round(Hoja.Height*FZoom/100)+10;
  Shower.Canvas.StretchDraw(Rect,Hoja);
end;

procedure TFrmPreview.BtnAnchoVClick(Sender: TObject);
begin
  ZoomShowerV.Value := Round((Scroller.Width-30)*100/(Hoja.Width));
end;

procedure TFrmPreview.BtnCompletaVClick(Sender: TObject);
var
  ZoomAncho: Double;
  ZoomAlto: Double;
begin
  ZoomAncho := (Scroller.ClientWidth-30)*100/(Hoja.Width);
  ZoomAlto := (Scroller.ClientHeight)*111/(Hoja.Height);

  if ZoomAncho < ZoomAlto then
    ZoomShowerV.Value := Round(ZoomAncho)
  else
    ZoomShowerV.Value := Round(ZoomAlto);
end;

procedure TFrmPreview.PageChange;
begin
  Enabled := False;
  GetHoja(FPagina,Hoja,False);
  UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);
  ShowerPaint(self);
  Enabled := True;
end;

procedure TFrmPreview.BtnImpActualVClick(Sender: TObject);
begin
  Enabled := False;
  TRSPrinter(RSPrinter).PrintPage(FPagina);
  Enabled := True;
  BringToFront;
end;

procedure TFrmPreview.BtnImpTodoVClick(Sender: TObject);
begin
  Enabled := False;
  TRSPrinter(RSPrinter).PrintAll;
  Enabled := True;
  Close;
end;

procedure TFrmPreview.ShowerMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  RealPositionH : integer;
  RealPositionV : integer;
  Muestra : integer;
begin
  RealPositionH := {Scroller.HorzScrollBar.ScrollPos +} X;
  RealPositionV := {Scroller.VertScrollBar.Position +} Y;

  if Scroller.HorzScrollBar.Visible then
  begin // HAY MOVIMIENTO HORIZONTAL
    Muestra := Scroller.Width;

    if Shower.Width < Muestra then
      Muestra := Shower.Width;

    if RealPositionH < (Muestra * 0.5) then
      Scroller.HorzScrollBar.Position := 0
    else
      Scroller.HorzScrollBar.Position := RealPositionH - (Muestra div 2);
  end;

  if Scroller.VertScrollBar.Visible then
  begin // HAY MOVIMIENTO VERTICAL
    Muestra := Scroller.Height;

    if Shower.Height < Muestra then
      Muestra := Shower.Height;

    if RealPositionV < (Muestra * 0.5) then
      Scroller.VertScrollBar.Position := 0
    else
      Scroller.VertScrollBar.Position := RealPositionV - (Muestra div 2);
  end;
end;


procedure TFrmPreview.BtnPropiedadesVClick(Sender: TObject);
var
  ADevice: array [0..255] of Char;
  ADriver: array [0..255] of Char;
  APort: array [0..255] of Char;
  DeviceMode: THandle;
  f: TFrmPrinterSetup;
begin
  f:= TFrmPrinterSetup.Create(application);

  // Passa para o form de configuracao o modo e impressora atual
  TRSPrinter(RSPrinter).GetModels(f.ImpressorasTexto);
  f.ImpTxtSel:=TRSPrinter(RSPrinter).GetModelName;

  if f.ShowModal = mrOK then
    begin
      if (printer.Canvas.TextHeight('X') = 1) or (printer.Canvas.TextHeight('X') = 83) then
      begin
        TRSPrinter(RSPrinter).Mode:= pmfast;
        TRSPrinter(RSPrinter).SetModelName(f.ImpTxtSel);
      end
      else
        TRSPrinter(RSPrinter).Mode:= pmwindows;

      // Refaz a pagina con as novas configuracoes
      try
        TRSPrinter(RSPrinter).PageOrientation := Printer.Orientation;
        TRSPrinter(RSPrinter).PageWidth := GetDeviceCaps(Printer.Handle, PHYSICALWIDTH);
        TRSPrinter(RSPrinter).PageHeight := GetDeviceCaps(Printer.Handle, PHYSICALHEIGHT);
        TRSPrinter(RSPrinter).PageWidthP := TRSPrinter(RSPrinter).PageWidth/GetDeviceCaps(Printer.Handle, LOGPIXELSX);
        TRSPrinter(RSPrinter).PageHeightP := TRSPrinter(RSPrinter).PageHeight/GetDeviceCaps(Printer.Handle, LOGPIXELSY);
        Printer.GetPrinter(ADevice, ADriver, APort, DeviceMode);
        TRSPrinter(RSPrinter).WinPrinter := ADevice;

        if TRSPrinter(RSPrinter).Lines <> Trunc(TRSPrinter(RSPrinter).PageHeightP*6)-4 then
        begin
          TRSPrinter(RSPrinter).Lines := Trunc(TRSPrinter(RSPrinter).PageHeightP*6)-4-TRSPrinter(RSPrinter).WinMarginTop div 12-TRSPrinter(RSPrinter).WinMarginBottom div 12;
          if assigned(TRSPrinter(RSPrinter).ReGenerate)then
          begin
            TRSPrinter(RSPrinter).ReGenerate;
            UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);

            if FPagina > TRSPrinter(RSPrinter).Paginas then
            begin
              FPagina := TRSPrinter(RSPrinter).Paginas;
              UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);
            end;

            if FPagina < TRSPrinter(RSPrinter).Paginas then
            begin
              BtnPagSigV.Enabled := True;
              BtnPagUltV.Enabled := True;
            end
            else
            begin
              BtnPagSigV.Enabled := False;
              BtnPagUltV.Enabled := False;
            end;
          end;
        end;
      except
      end;

      Hoja.Free;
      Hoja := TMetaFile.Create;
      Hoja.Width := Round(TRSPrinter(RSPrinter).PageWidthP * 80); // div 4; // 640;
      Hoja.Height := Round(TRSPrinter(RSPrinter).PageHeightP * 80); // div 4; // 1056;
      HojasNum.Clear;
      While HojasDib.Count > 0 do
      begin
        TMemoryStream(HojasDib[0]).Free;
        HojasDib.Delete(0);
      end;

      GetHoja(FPagina,Hoja,False);
      Shower.Invalidate;
      UpdateStatus(fpagina,TRSPrinter(RSPrinter).Paginas);
    end;

  f.Free;
end;

procedure TFrmPreview.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = '+' then
  begin
    ZoomShowerV.Value := ZoomShowerV.Value + 20;
    Key := #0;
  end
  else if Key = '-' then
  begin
    ZoomShowerV.Value := ZoomShowerV.Value - 20;
    Key := #0;
  end;
end;

procedure TFrmPreview.GetHoja(Numero: integer; Hoja: TMetaFile; ViewDialog : boolean);
var
  formcaption: string;
  Data: TMemoryStream;
  M: TMemoryStream;
  C: TCompressionStream;
  buf: Pointer;
  size: Integer;
begin
  if HojasNum.IndexOf(Pointer(Numero)) >= 0 then
  begin
    Data := TMemoryStream(HojasDib[HojasNum.IndexOf(Pointer(Numero))]);

    try
      ZDecompress(Data.Memory, Data.Size * 3, buf, size);
    except
      on E: Exception do
      begin
        E.Message := Format('Error Decompressing Buffer (Len = %d):'#13#10'%s', [Data.Size, e.Message]);
        raise;
      end;
    end;

    M := TMemoryStream.Create;
    M.Write(buf^, size);
    FreeMem(buf);
    M.Position := 0;
    Hoja.LoadFromStream(M);
    M.Free;
  end
  else
  begin
    formcaption := caption;

    if not ViewDialog then
      Caption := formcaption + ' - ' + 'Preparando Página';

    TRSPrinter(RSPrinter).BuildPage(Numero,Hoja,False,TRSPrinter(RSPrinter).Mode,ViewDialog);

    if not ViewDialog then
      Caption := formcaption;

    M := TMemoryStream.Create;
    C := TCompressionStream.Create(clFastest, M);
    Hoja.SaveToStream(C);
    C.Free;
    HojasNum.Add(Pointer(Numero));
    HojasDib.Add(M);
  end;
end;

procedure TFrmPreview.LTotPagHClick(Sender: TObject);
var
  Txt: string;
begin
  Txt := '';

  if InputQuery('Acesso Rápido','Digite o número da página',Txt) and (StrToIntDef(Txt,0)>0) and (StrToIntDef(Txt,0)<=TRSPrinter(RSPrinter).Paginas) and (StrToIntDef(Txt,0)<>FPagina) then
  begin
    FPagina := StrToIntDef(Txt,0);
    PageChange;
  end;
end;

procedure TFrmPreview.FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if (Scroller.VertScrollBar.Position < Scroller.VertScrollBar.Range) then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position + 100;
    if Scroller.VertScrollBar.Position > Scroller.VertScrollBar.Range then
      Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Range;
  end;
end;

procedure TFrmPreview.FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if (Scroller.VertScrollBar.Position > 0) then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position - 100;
    if Scroller.VertScrollBar.Position < 0 then
      Scroller.VertScrollBar.Position := 0;
  end;
end;

procedure TFrmPreview.ActUpExecute(Sender: TObject);
begin
  if (Scroller.VertScrollBar.Position > 0) then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position - 100;
    if Scroller.VertScrollBar.Position < 0 then
      Scroller.VertScrollBar.Position := 0;
  end;
end;

procedure TFrmPreview.ActDownExecute(Sender: TObject);
begin
  if (Scroller.VertScrollBar.Position < Scroller.VertScrollBar.Range) then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position + 100;
    if Scroller.VertScrollBar.Position > Scroller.VertScrollBar.Range then
      Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Range;
  end;
end;

procedure TFrmPreview.ActFecharExecute(Sender: TObject);
begin
  Close;
end;

procedure TFrmPreview.ActRightExecute(Sender: TObject);
begin
  if (Scroller.HorzScrollBar.Position < Scroller.HorzScrollBar.Range) then
  begin
    Scroller.HorzScrollBar.Position := Scroller.HorzScrollBar.Position + 20;
    if Scroller.HorzScrollBar.Position > Scroller.HorzScrollBar.Range then
      Scroller.HorzScrollBar.Position := Scroller.HorzScrollBar.Range;
  end;
end;

procedure TFrmPreview.ActLeftExecute(Sender: TObject);
begin
  if (Scroller.HorzScrollBar.Position > 0) then
  begin
    Scroller.HorzScrollBar.Position := Scroller.HorzScrollBar.Position - 20;
    if Scroller.HorzScrollBar.Position < 0 then
      Scroller.HorzScrollBar.Position := 0;
  end;
end;

procedure TFrmPreview.ActHomeExecute(Sender: TObject);
begin
  if FPagina > 1 then
  begin
    FPagina := 1;
    PageChange;
  end;
end;

procedure TFrmPreview.ActEndExecute(Sender: TObject);
begin
  if FPagina < TRSPrinter(RSPrinter).Paginas then
  begin
    FPagina := TRSPrinter(RSPrinter).Paginas;
    PageChange;
  end;
end;

procedure TFrmPreview.ActF1Execute(Sender: TObject);
begin
  ShowMessage('Ajuda'+#13+#13+
              'Inicio do relatorio    - Home'+#13+
              'Final do relatorio     - End'+#13+
              'Pagina anterior       - PgUp'+#13+
              'Pagina Sequinte      - PgDn'+#13+
              'Imprimir                   - Enter'+#13+
              'Fechar                     - Esc' );
end;

procedure TFrmPreview.ActPgUpExecute(Sender: TObject);
begin
  if FPagina > 1 then
  begin
    Dec(FPagina);
    PageChange;
  end;
end;

procedure TFrmPreview.ActPgDnExecute(Sender: TObject);
begin
  if FPagina < TRSPrinter(RSPrinter).Paginas then
  begin
    Inc(FPagina);
    PageChange;
  end;
end;

procedure TFrmPreview.ActEnterExecute(Sender: TObject);
begin
  ShowMessage('imprimir');
end;

end.




unit RSPrint.Preview;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, ExtCtrls, Spin, Printers, Zlib,
  Buttons, ActnList, ComCtrls, RSPrint;

type
  TFrmPreview = class(TForm)
    Scroller: TScrollBox;
    Display: TPaintBox;
    PrinterPanelH: TPanel;
    labCopies: TLabel;
    edtCopies: TSpinEdit;
    btnFirstPage: TSpeedButton;
    btnPriorPage: TSpeedButton;
    btnGoToPageNumber: TSpeedButton;
    btnNextPage: TSpeedButton;
    btnLastPage: TSpeedButton;
    edtZoom: TSpinEdit;
    labZoom: TLabel;
    btnFitOnWindow: TSpeedButton;
    btnShowEntirePage: TSpeedButton;
    btnPrintCurrentPage: TSpeedButton;
    btnPrintAllPages: TSpeedButton;
    btnPrintSetup: TSpeedButton;
    Bevel1: TBevel;
    btnClose: TSpeedButton;
    Bevel2: TBevel;
    Bevel3: TBevel;
    ActionList: TActionList;
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
    procedure edtCopiesChange(Sender: TObject);
    procedure edtZoomChange(Sender: TObject);
    procedure DisplayPaint(Sender: TObject);
    procedure btnFitOnWindowClick(Sender: TObject);
    procedure btnShowEntirePageClick(Sender: TObject);
    procedure btnPrintCurrentPageClick(Sender: TObject);
    procedure btnPrintAllPagesClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DisplayMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnPrintSetupClick(Sender: TObject);
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

  private
    FPagina: integer;
    FZoom: Double;
    DontDraw: Boolean;
    HojasDib: TList;
    HojasNum: TList;

    procedure GetHoja(Numero: Integer; Hoja: TMetaFile; ViewDialog: boolean);
    procedure UpdateStatus(Pag_Atual, Total_paginas: Integer);
    procedure PageChange;

  public
    FRSPrinter: TRSPrinter;
    Hoja: TMetaFile;

    class procedure Execute(AOwner: TRSPrinter);

  end;

implementation

{$R *.DFM}

uses
  RSPrint.Preview.PrinterSetup, RSPrint.Types.CommonTypes, ZLibConst;

class procedure TFrmPreview.Execute(AOwner: TRSPrinter);
var
  FrmPreview: TFrmPreview;
begin
  FrmPreview := TFrmPreview.Create(AOwner);
  try
    FrmPreview.FRSPrinter := AOwner;
    FrmPreview.ShowModal;
  finally
    FrmPreview.Release;
  end;
end;

procedure TFrmPreview.UpdateStatus(Pag_Atual, Total_paginas: Integer);
var
  Modo : string;
begin
  if FRSPrinter.Mode = pmFast then
    Modo:= 'Matricial'
  else
    Modo := 'Gráfico';

  StatusBar1.Panels[0].Text := 'Página: '+IntToStr(Pag_atual) + ' de ' + IntToStr(total_paginas);
  StatusBar1.Panels[1].Text := 'Zoom: ' + FloatToStr(FZoom);
  StatusBar1.Panels[2].Text := 'Modo: ' + Modo;
  StatusBar1.Panels[3].Text := '  Impressora: ' + Printer.Printers[Printer.PrinterIndex];

  Caption := 'Preview de Impressão      Página: ' + IntToStr(Pag_atual) + ' de ' + IntToStr(total_paginas);

  btnFirstPage.Enabled := (FPagina > 1);
  btnPriorPage.Enabled := (FPagina > 1);

  btnNextPage.Enabled := (FRSPrinter.Paginas > FPagina) and (FRSPrinter.Paginas > 1);
  btnLastPage.Enabled := (FRSPrinter.Paginas > FPagina) and (FRSPrinter.Paginas > 1);

  btnGoToPageNumber.Enabled := (FRSPrinter.Paginas > 1);
  btnPrintCurrentPage.Enabled := (FRSPrinter.Paginas > 1);

end;

procedure TFrmPreview.FormShow(Sender: TObject);
var
  ZoomAncho: Double;
  ZoomAlto: Double;
begin
  FPagina := 1;
  Hoja := TMetaFile.Create;
  Hoja.Width := Round(FRSPrinter.PageWidthP * 80);
  Hoja.Height := Round(FRSPrinter.PageHeightP * 80);
  Display.Width := Display.Canvas.TextWidth('X')*80+20;
  Display.Height := Display.Canvas.TextHeight('X')*66+20;

  FPagina := 1;
  UpdateStatus(FPagina, FRSPrinter.Paginas);
  GetHoja(FPagina,Hoja,True);

  if FRSPrinter.Zoom = zWidth then
    edtZoom.Value := Round((Screen.Width-30)*100/(Hoja.Width))
  else // a lo alto
  begin
      ZoomAncho := (Screen.Width-30)*100/(Hoja.Width);
      ZoomAlto := (Scroller.ClientHeight)*111/(Hoja.Height);
      if ZoomAncho < ZoomAlto then
        edtZoom.Value := Round(ZoomAncho)
      else
        edtZoom.Value := Round(ZoomAlto);
  end;

  edtCopies.Value := FRSPrinter.Copies;
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

procedure TFrmPreview.edtCopiesChange(Sender: TObject);
begin
  FRSPrinter.Copies := edtCopies.Value;
end;

procedure TFrmPreview.edtZoomChange(Sender: TObject);
begin
  FZoom := edtZoom.Value;
  DontDraw := true;
  DisplayPaint(self);
  UpdateStatus(FPagina, FRSPrinter.Paginas);
end;

procedure TFrmPreview.DisplayPaint(Sender: TObject);
var
  Rect: TRect;
begin
  if DontDraw then
  begin
    Display.Width := Round(Hoja.Width*FZoom/100)+10;
    Display.Height := Round(Hoja.Height*FZoom/100)+10;
    DontDraw := False;
    Exit;
  end;

  Rect.Left := 10;
  Rect.Top := 10;
  Rect.Right := Round(Hoja.Width*FZoom/100);
  Rect.Bottom := Round(Hoja.Height*FZoom/100);
  Display.Width := Round(Hoja.Width*FZoom/100)+10;
  Display.Height := Round(Hoja.Height*FZoom/100)+10;
  Display.Canvas.StretchDraw(Rect,Hoja);
end;

procedure TFrmPreview.btnFitOnWindowClick(Sender: TObject);
begin
  edtZoom.Value := Round((Scroller.Width-30)*100/(Hoja.Width));
end;

procedure TFrmPreview.btnShowEntirePageClick(Sender: TObject);
var
  ZoomAncho: Double;
  ZoomAlto: Double;
begin
  ZoomAncho := (Scroller.ClientWidth-30)*100/(Hoja.Width);
  ZoomAlto := (Scroller.ClientHeight)*111/(Hoja.Height);

  if ZoomAncho < ZoomAlto then
    edtZoom.Value := Round(ZoomAncho)
  else
    edtZoom.Value := Round(ZoomAlto);
end;

procedure TFrmPreview.PageChange;
begin
  Enabled := False;
  GetHoja(FPagina,Hoja,False);
  UpdateStatus(FPagina, FRSPrinter.Paginas);
  DisplayPaint(self);
  Enabled := True;
end;

procedure TFrmPreview.btnPrintCurrentPageClick(Sender: TObject);
begin
  Enabled := False;
  FRSPrinter.PrintPage(FPagina);
  Enabled := True;
  BringToFront;
end;

procedure TFrmPreview.btnPrintAllPagesClick(Sender: TObject);
begin
  Enabled := False;
  FRSPrinter.PrintAll;
  Enabled := True;
  Close;
end;

procedure TFrmPreview.DisplayMouseDown(Sender: TObject;
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

    if Display.Width < Muestra then
      Muestra := Display.Width;

    if RealPositionH < (Muestra * 0.5) then
      Scroller.HorzScrollBar.Position := 0
    else
      Scroller.HorzScrollBar.Position := RealPositionH - (Muestra div 2);
  end;

  if Scroller.VertScrollBar.Visible then
  begin // HAY MOVIMIENTO VERTICAL
    Muestra := Scroller.Height;

    if Display.Height < Muestra then
      Muestra := Display.Height;

    if RealPositionV < (Muestra * 0.5) then
      Scroller.VertScrollBar.Position := 0
    else
      Scroller.VertScrollBar.Position := RealPositionV - (Muestra div 2);
  end;
end;


procedure TFrmPreview.btnPrintSetupClick(Sender: TObject);
var
  ADevice: array [0..255] of Char;
  ADriver: array [0..255] of Char;
  APort: array [0..255] of Char;
  DeviceMode: THandle;
  f: TFrmPrinterSetup;
begin
  f:= TFrmPrinterSetup.Create(application);

  // Passa para o form de configuracao o modo e impressora atual
  FRSPrinter.GetModels(f.ImpressorasTexto);
  f.ImpTxtSel := FRSPrinter.GetModelName;

  if f.ShowModal = mrOK then
    begin
      if (printer.Canvas.TextHeight('X') = 1) or (printer.Canvas.TextHeight('X') = 83) then
      begin
        FRSPrinter.Mode:= pmfast;
        FRSPrinter.SetModelName(f.ImpTxtSel);
      end
      else
        FRSPrinter.Mode:= pmwindows;

      // Refaz a pagina con as novas configuracoes
      try
        FRSPrinter.PageOrientation := Printer.Orientation;
        FRSPrinter.PageWidth := GetDeviceCaps(Printer.Handle, PHYSICALWIDTH);
        FRSPrinter.PageHeight := GetDeviceCaps(Printer.Handle, PHYSICALHEIGHT);
        FRSPrinter.PageWidthP := FRSPrinter.PageWidth/GetDeviceCaps(Printer.Handle, LOGPIXELSX);
        FRSPrinter.PageHeightP := FRSPrinter.PageHeight/GetDeviceCaps(Printer.Handle, LOGPIXELSY);
        Printer.GetPrinter(ADevice, ADriver, APort, DeviceMode);
        FRSPrinter.WinPrinter := ADevice;

        if FRSPrinter.Lines <> Trunc(FRSPrinter.PageHeightP*6)-4 then
        begin
          FRSPrinter.Lines := Trunc(FRSPrinter.PageHeightP*6)-4-FRSPrinter.WinMarginTop div 12-FRSPrinter.WinMarginBottom div 12;
          if Assigned(FRSPrinter.ReGenerate)then
          begin
            FRSPrinter.ReGenerate;
            UpdateStatus(FPagina, FRSPrinter.Paginas);

            if FPagina > FRSPrinter.Paginas then
            begin
              FPagina := FRSPrinter.Paginas;
              UpdateStatus(FPagina, FRSPrinter.Paginas);
            end;

            if FPagina < FRSPrinter.Paginas then
            begin
              btnNextPage.Enabled := True;
              btnLastPage.Enabled := True;
            end
            else
            begin
              btnNextPage.Enabled := False;
              btnLastPage.Enabled := False;
            end;
          end;
        end;
      except
      end;

      Hoja.Free;
      Hoja := TMetaFile.Create;
      Hoja.Width := Round(FRSPrinter.PageWidthP * 80); // div 4; // 640;
      Hoja.Height := Round(FRSPrinter.PageHeightP * 80); // div 4; // 1056;
      HojasNum.Clear;
      While HojasDib.Count > 0 do
      begin
        TMemoryStream(HojasDib[0]).Free;
        HojasDib.Delete(0);
      end;

      GetHoja(FPagina,Hoja,False);
      Display.Invalidate;
      UpdateStatus(FPagina, FRSPrinter.Paginas);
    end;

  f.Free;
end;

procedure TFrmPreview.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = '+' then
  begin
    edtZoom.Value := edtZoom.Value + 20;
    Key := #0;
  end
  else if Key = '-' then
  begin
    edtZoom.Value := edtZoom.Value - 20;
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

    FRSPrinter.BuildPageToPreview(Numero, Hoja, FRSPrinter.Mode, ViewDialog);

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

  if InputQuery('Acesso Rápido','Digite o número da página',Txt) and (StrToIntDef(Txt,0)>0) and (StrToIntDef(Txt,0)<=FRSPrinter.Paginas) and (StrToIntDef(Txt,0)<>FPagina) then
  begin
    FPagina := StrToIntDef(Txt,0);
    PageChange;
  end;
end;

procedure TFrmPreview.FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Scroller.VertScrollBar.Position < Scroller.VertScrollBar.Range then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position + 100;

    if Scroller.VertScrollBar.Position > Scroller.VertScrollBar.Range then
      Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Range;
  end;
end;

procedure TFrmPreview.FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Scroller.VertScrollBar.Position > 0 then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position - 100;

    if Scroller.VertScrollBar.Position < 0 then
      Scroller.VertScrollBar.Position := 0;
  end;
end;

procedure TFrmPreview.ActUpExecute(Sender: TObject);
begin
  if Scroller.VertScrollBar.Position > 0 then
  begin
    Scroller.VertScrollBar.Position := Scroller.VertScrollBar.Position - 100;

    if Scroller.VertScrollBar.Position < 0 then
      Scroller.VertScrollBar.Position := 0;
  end;
end;

procedure TFrmPreview.ActDownExecute(Sender: TObject);
begin
  if Scroller.VertScrollBar.Position < Scroller.VertScrollBar.Range then
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
  if Scroller.HorzScrollBar.Position < Scroller.HorzScrollBar.Range then
  begin
    Scroller.HorzScrollBar.Position := Scroller.HorzScrollBar.Position + 20;

    if Scroller.HorzScrollBar.Position > Scroller.HorzScrollBar.Range then
      Scroller.HorzScrollBar.Position := Scroller.HorzScrollBar.Range;
  end;
end;

procedure TFrmPreview.ActLeftExecute(Sender: TObject);
begin
  if Scroller.HorzScrollBar.Position > 0 then
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
  if FPagina < FRSPrinter.Paginas then
  begin
    FPagina := FRSPrinter.Paginas;
    PageChange;
  end;
end;

procedure TFrmPreview.ActF1Execute(Sender: TObject);
begin
  ShowMessage('Ajuda'+#13+#13+
              'Inicio do relatorio    - Home'+#13+
              'Final do relatorio     - End'+#13+
              'Pagina anterior       - PgUp'+#13+
              'Pagina seguinte      - PgDn'+#13+
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
  if FPagina < FRSPrinter.Paginas then
  begin
    Inc(FPagina);
    PageChange;
  end;
end;

end.




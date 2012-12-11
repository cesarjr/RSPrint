unit RSPrint.PrintThread;

interface

uses
  Classes, RSPrint.RSPrintTray, RSPrint.CommonTypes, RSPrint.PrintThread.FastPrintDevice;

type
  TPrintThread = class(TThread)
  private const
    SINGLE_LINE = #196;
    DOUBLE_LINE = #205;

  private
    FJob: PPrintJob;
    FJobs: TThreadList;
    FTrayIcon: TRSPrintTray;
    FPrinterStatus: TPrinterStatus;
    FFastPrintDevice: IFastPrintDevice;

    procedure SetJobName;
    procedure InflateLineWithSpaces(var line: string; maxCol: Byte);
    procedure PageContinuosJump;
    procedure ProcessPauseIfNecessary;

    procedure PrintControlCodes(codeSequence: string);
    procedure PrintHorizontalLines(currentLine: Integer; page: PPage; var lineToPrint: string);
    procedure PrintVerticalLines(page: PPage; var lineToPrint: string; currentLine: Integer);
    procedure PrintCurrentLine(page: PPage; var ultimaEscritura: Integer; font: TFastFont; var lineToPrint: string; const currentLine: Integer);

  protected
    procedure Execute; override;
    function PrintFast(number: integer) : boolean;

  public
    constructor Create(printerStatus: TPrinterStatus); reintroduce;
    destructor Destroy; override;

    procedure AddJob(job: PPrintJob);
  end;

implementation

uses
  SysUtils, Windows, Printers, RSPrint.Utils, RSPrint.PrintThread.FastPrintSpool, Dialogs;

procedure TPrintThread.AddJob(job: PPrintJob);
begin
  FJobs.Add(job);
end;

procedure TPrintThread.PrintVerticalLines(page: PPage; var lineToPrint: string; currentLine: Integer);
var
  VerticalLine: PVertLine;
  Contador: Integer;
begin
  // AHORA SE ANALIZAN LAS LINEAS VERTICALES
  for Contador := 0 to page.VerticalLines.Count - 1 do
  begin
    VerticalLine := page.VerticalLines.Items[Contador];
    if (VerticalLine^.Line1 <= currentLine) and (VerticalLine^.Line2 >= currentLine) then
    begin
      // LA LINEA PASA POR ESTA LINEA
      InflateLineWithSpaces(lineToPrint, VerticalLine^.Col);
      if VerticalLine^.Line1 = currentLine then
      begin
        // ES LA PRIMER LINEA
        if lineToPrint[VerticalLine^.Col] = SINGLE_LINE then
        // LINEA HORIZONTAL SIMPLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [192, 193, 194, 195, 196, 197, 199, 208, 210, 211, 214, 215, 218]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Â'
              else
                lineToPrint[VerticalLine^.Col] := 'Ò';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := '¿'
              else
                lineToPrint[VerticalLine^.Col] := '·';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ú'
              else
                lineToPrint[VerticalLine^.Col] := 'Ö';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Â'
              else
                lineToPrint[VerticalLine^.Col] := 'Ò';
            end;
          end;
        end
        else if lineToPrint[VerticalLine^.Col] = 'Í' then
        // LINEA HORIZONTAL DOBLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [198, 200, 201, 202, 203, 204, 205, 206, 207, 209, 212, 213]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [181, 184, 185, 187, 188, 190, 202, 203, 205, 206, 207, 209, 216]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ñ'
              else
                lineToPrint[VerticalLine^.Col] := 'Ë';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := '¸'
              else
                lineToPrint[VerticalLine^.Col] := '»';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [181, 184, 185, 187, 188, 190, 202, 203, 205, 206, 207, 209, 216]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Õ'
              else
                lineToPrint[VerticalLine^.Col] := 'É';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ñ'
              else
                lineToPrint[VerticalLine^.Col] := 'Ë';
            end;
          end;
        end
        else
        // HAY OTRO CODIGO
        begin
          if VerticalLine^.Kind = ltSingle then
            lineToPrint[VerticalLine^.Col] := '³'
          else
            // Doble
            lineToPrint[VerticalLine^.Col] := 'º';
        end;
      end
      else if VerticalLine^.Line2 = currentLine then
      begin
        // ES LA ULTIMA LINEA
        if lineToPrint[VerticalLine^.Col] = SINGLE_LINE then
        // LINEA HORIZONTAL SIMPLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [192, 193, 194, 195, 196, 197, 199, 208, 210, 211, 214, 215, 218]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Á'
              else
                lineToPrint[VerticalLine^.Col] := 'Ï';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ù'
              else
                lineToPrint[VerticalLine^.Col] := '½';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'À'
              else
                lineToPrint[VerticalLine^.Col] := 'Ó';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
              else
                lineToPrint[VerticalLine^.Col] := 'Ï';
            end;
          end;
        end
        else if lineToPrint[VerticalLine^.Col] = 'Í' then
        // LINEA HORIZONTAL DOBLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [181, 184, 185, 187, 188, 190, 202, 203, 205, 206, 207, 209, 216]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [198, 200, 201, 202, 203, 204, 205, 206, 207, 209, 212, 213]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ï'
              else
                lineToPrint[VerticalLine^.Col] := 'Ê';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := '¾'
              else
                lineToPrint[VerticalLine^.Col] := '¼';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [198, 200, 201, 202, 203, 204, 205, 206, 207, 209, 212, 213]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ô'
              else
                lineToPrint[VerticalLine^.Col] := 'È';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ð'
              else
                lineToPrint[VerticalLine^.Col] := 'Ê';
            end;
          end;
        end
        else
        // HAY OTRO CODIGO
        begin
          if VerticalLine^.Kind = ltSingle then
            lineToPrint[VerticalLine^.Col] := '³'
          else
            // Doble
            lineToPrint[VerticalLine^.Col] := 'º';
        end;
      end
      else
      begin
        // ES UNA LINEA DEL MEDIO
        if lineToPrint[VerticalLine^.Col] = 'Ä' then
        // LINEA HORIZONTAL SIMPLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [192, 193, 194, 195, 196, 197, 199, 208, 210, 211, 214, 215, 218]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Å'
              else
                lineToPrint[VerticalLine^.Col] := '×';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := '´'
              else
                lineToPrint[VerticalLine^.Col] := '¶';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [180, 182, 183, 189, 191, 193, 194, 196, 197, 208, 210, 215, 217]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ã'
              else
                lineToPrint[VerticalLine^.Col] := 'Ç';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Å'
              else
                lineToPrint[VerticalLine^.Col] := '×';
            end;
          end;
        end
        else if lineToPrint[VerticalLine^.Col] = 'Í' then
        // LINEA HORIZONTAL DOBLE
        begin
          if (VerticalLine^.Col > 0) and (ord(lineToPrint[VerticalLine^.Col - 1]) in [198, 200, 201, 202, 203, 204, 205, 206, 207, 209, 212, 213]) then
          begin
            // VIENE DE LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [181, 184, 185, 187, 188, 190, 202, 203, 205, 206, 207, 209, 216]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ø'
              else
                lineToPrint[VerticalLine^.Col] := 'Î';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'µ'
              else
                lineToPrint[VerticalLine^.Col] := '¹';
            end;
          end
          else
          begin
            // NO VA PARA LA IZQUIERDA
            if (VerticalLine^.Col < Length(lineToPrint)) and (ord(lineToPrint[VerticalLine^.Col + 1]) in [181, 184, 185, 187, 188, 190, 202, 203, 205, 206, 207, 209, 216]) then
            begin
              // SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Æ'
              else
                lineToPrint[VerticalLine^.Col] := 'Ì';
            end
            else
            begin
              // NO SIGUE A LA DERECHA
              if VerticalLine^.Kind = ltSingle then
                lineToPrint[VerticalLine^.Col] := 'Ø'
              else
                lineToPrint[VerticalLine^.Col] := 'Î';
            end;
          end;
        end
        else
        begin
          if VerticalLine^.Kind = ltSingle then
            lineToPrint[VerticalLine^.Col] := '³'
          else
            // Doble
            lineToPrint[VerticalLine^.Col] := 'º';
        end;
      end;
    end;
  end;
end;

procedure TPrintThread.PrintHorizontalLines(currentLine: Integer; page: PPage; var lineToPrint: string);
var
  HorizontalLine: PHorizLine;
  I: Integer;
  Contador: Integer;
begin
  // ANALIZO PRIMERO LAS LINEAS HORIZONTALES
  for Contador := 0 to page.HorizLines.Count - 1 do
  begin
    HorizontalLine := page.HorizLines.Items[Contador];
    // ES EN ESTA LINEA
    if HorizontalLine^.Line = currentLine then
    begin
      InflateLineWithSpaces(lineToPrint, HorizontalLine^.Col2);
      if HorizontalLine^.Kind = ltSingle then
      begin
        for I := HorizontalLine^.Col1 to HorizontalLine^.Col2 do
          lineToPrint[I] := SINGLE_LINE;
      end
      else
      begin
        for I := HorizontalLine^.Col1 to HorizontalLine^.Col2 do
          lineToPrint[I] := DOUBLE_LINE;
      end;
    end;
  end;
end;

constructor TPrintThread.Create(printerStatus: TPrinterStatus);
begin
  inherited Create(True);

  FJobs := TThreadList.Create;
  FTrayIcon := TRSPrintTray.Create(nil, FPrinterStatus);
  FPrinterStatus := printerStatus;

  FreeOnTerminate := True;
  FPrinterStatus.StartPrinting;

  FFastPrintDevice := TFastPrintSpool.Create;
end;

destructor TPrintThread.Destroy;
begin
  FJobs.Free;
  FPrinterStatus.CurrentlyPrinting := False;

  inherited;
end;

procedure TPrintThread.InflateLineWithSpaces(var line: string; maxCol: Byte);
begin
  while Length(line) < maxCol do
    line := line + ' ';
end;

procedure TPrintThread.PrintControlCodes(codeSequence: string);
const
  LAST_CODE_OF_SEQUENCE = 0;
  SEPARATOR = #32;
var
  CodesToPrint: string;
  CodeToPrint: byte;
  NextCodePosition: byte;
begin
  CodesToPrint := codeSequence;

  while Length(CodesToPrint) > 0 do
  begin
    NextCodePosition := Pos(SEPARATOR, CodesToPrint);

    if NextCodePosition = LAST_CODE_OF_SEQUENCE then
    begin
      try
        CodeToPrint := StrToInt(CodesToPrint);
        FFastPrintDevice.Write(Chr(CodeToPrint));
        CodesToPrint := '';
      except
      end;
    end
    else
    begin // HAY MAS CODIGOS
      try
        CodeToPrint := StrToInt(Copy(CodesToPrint, 1, NextCodePosition-1));
        FFastPrintDevice.Write(Chr(CodeToPrint));
        CodesToPrint := Copy(CodesToPrint, NextCodePosition+1, Length(CodesToPrint)-3);
      except
      end;
    end;
  end;
end;

procedure TPrintThread.Execute;
var
  Bien: Boolean;
  Copias: Integer;
  I: Integer;
  List: TList;
begin
  List := FJobs.LockList;

  if List.Count > 0 then
  begin
    FJob := pPrintJob(List[0]);
    FPrinterStatus.PrintingJobName := FJob^.Name;
  end
  else
  begin
    FJob := nil;
    FPrinterStatus.PrintingJobName := '';
  end;
  Synchronize(SetJobName);

  FJobs.UnlockList;

  ProcessPauseIfNecessary;

  while (FJob <> nil) do
  begin
    Bien := True;
    for Copias := 1 to FJob^.FCopias do
      for I := 1 to FJob^.LasPaginas.Count do
        if not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll and Bien then
          Bien := Bien and PrintFast(I);

    while FJob^.LasPaginas.Count > 0 do
    begin
      while PPage(FJob^.LasPaginas[0])^.Writed.Count > 0 do
      begin
        Dispose(pWrite(PPage(FJob^.LasPaginas[0])^.Writed[0]));
        PPage(FJob^.LasPaginas[0])^.Writed.Delete(0);
      end;
      PPage(FJob^.LasPaginas[0])^.Writed.Free;

      while PPage(FJob^.LasPaginas[0])^.VerticalLines.Count > 0 do
      begin
        Dispose(pVertLine(PPage(FJob^.LasPaginas[0])^.VerticalLines[0]));
        PPage(FJob^.LasPaginas[0])^.VerticalLines.Delete(0);
      end;
      PPage(FJob^.LasPaginas[0])^.VerticalLines.Free;

      while PPage(FJob^.LasPaginas[0])^.HorizLines.Count > 0 do
      begin
        Dispose(pHorizLine(PPage(FJob^.LasPaginas[0])^.HorizLines[0]));
        PPage(FJob^.LasPaginas[0])^.HorizLines.Delete(0);
      end;
      PPage(FJob^.LasPaginas[0])^.HorizLines.Free;

      Dispose(PPage(FJob^.LasPaginas[0]));
      FJob^.LasPaginas.Delete(0);
    end;
    FJob^.LasPaginas.Free;

    Dispose(FJob);
    FPrinterStatus.PrintingCanceled := False;
    List := FJobs.LockList;
    List.Delete(0);

    if List.Count > 0 then
    begin
      FJob := pPrintJob(List[0]);
      FPrinterStatus.PrintingJobName := FJob^.Name;
    end
    else
    begin
      FJob := nil;
      FPrinterStatus.PrintingJobName := '';
    end;
  end;
  Synchronize(SetJobName);

  FPrinterStatus.PrintingJobName := '';
end;

procedure TPrintThread.ProcessPauseIfNecessary;
begin
  while FPrinterStatus.PrintingPaused and not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll do
    Sleep(100);
end;

function TPrintThread.PrintFast(number: integer): boolean;
var
  CurrentLine: Integer;
  LineToPrint: string;
  Page: PPage;
  Font: TFastFont;
  UltimaEscritura: Integer;
  Resultado: Boolean;
  ListaImpressoras: TstringList;
  PrinterId: Integer;
begin
  ListaImpressoras:= TStringList.Create;
  TUtils.EnumPrt(ListaImpressoras, PrinterId);

  ProcessPauseIfNecessary;

  try
    Resultado := True;
    FFastPrintDevice.BeginDoc;

    PrintControlCodes(FJob.ControlCodes.Reset);
    PrintControlCodes(FJob.ControlCodes.Setup);

    PrintControlCodes(FJob.ControlCodes.SelLength);
    FFastPrintDevice.Write(#84);


    Font := FJob.FFuente;
    PrintControlCodes(FJob.ControlCodes.Normal);
    if Bold in Font then
      PrintControlCodes(FJob.ControlCodes.Bold);
    if Italic in Font then
      PrintControlCodes(FJob.ControlCodes.Italic);
    if DobleWide in Font then
      PrintControlCodes(FJob.ControlCodes.Wide)
    else if Compress in Font then
      PrintControlCodes(FJob.ControlCodes.CondensedON)
    else
      PrintControlCodes(FJob.ControlCodes.CondensedOFF);
    if Underline in Font then
      PrintControlCodes(FJob.ControlCodes.UnderlineON)
    else
      PrintControlCodes(FJob.ControlCodes.UnderlineOFF);

    UltimaEscritura := 0;
    Page := FJob.LasPaginas.Items[number-1];
    CurrentLine := 1;
    while (CurrentLine < TUtils.Min(Page.PrintedLines, FJob.FLineas)) and (not FPrinterStatus.PrintingCanceled) and not (FPrinterStatus.PrintingCancelAll) do
    begin // SE IMPRIMEN TODAS LAS LINEAS
      LineToPrint := '';

      PrintHorizontalLines(CurrentLine, Page, LineToPrint);
      PrintVerticalLines(Page, LineToPrint, CurrentLine);

      PrintCurrentLine(Page, UltimaEscritura, Font, LineToPrint, CurrentLine);
      Inc(CurrentLine);

      ProcessPauseIfNecessary;
    end;

    if (FJob.PageSize = pzContinuous) and (FJob.PageLength = 0) then
      PageContinuosJump
    else
      FFastPrintDevice.Write(#12);
  except
    //Resultado := False;
    on e: exception do
      ShowMessage(e.Message);
  end;

  FFastPrintDevice.EndDoc;
  PrintFast := Resultado;
end;

procedure TPrintThread.PageContinuosJump;
var
  LinesToJump: Integer;
begin
  for LinesToJump := 1 to FJob.PageContinuousJump do
    FFastPrintDevice.WriteLn('');
end;

procedure TPrintThread.PrintCurrentLine(page: PPage; var ultimaEscritura: Integer; font: TFastFont; var lineToPrint: string; const currentLine: Integer);
var
  Escritura: PWrite;
  i: Integer;
  Contador: Integer;
  Columna: Byte;
  LineWasPrinted: Boolean;
  Txt : string;
begin
  if page.Writed.Count = UltimaEscritura then
  begin
    if font <> FJob.FFuente then
    begin // PONEMOS LA FUENTE POR DEFAULT
      font := FJob.FFuente;
      PrintControlCodes(FJob.ControlCodes.Normal);
      if Bold in font then
        PrintControlCodes(FJob.ControlCodes.Bold);
      if Italic in font then
        PrintControlCodes(FJob.ControlCodes.Italic);
      if DobleWide in font then
        PrintControlCodes(FJob.ControlCodes.Wide)
      else if Compress in font then
        PrintControlCodes(FJob.ControlCodes.CondensedON)
      else
        PrintControlCodes(FJob.ControlCodes.CondensedOFF);
      if Underline in font then
        PrintControlCodes(FJob.ControlCodes.UnderlineON)
      else
        PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
    end;
    if FJob.FTransliterate and (lineToPrint <> '') then
      CharToOemBuff(PChar(@lineToPrint[1]), PansiChar(@lineToPrint[1]),Length(lineToPrint));
    FFastPrintDevice.WriteLn(lineToPrint);
  end
  else
  begin
    LineWasPrinted := False;
    Contador := UltimaEscritura;
    Columna := 1;
    Escritura := page.Writed.Items[Contador];
    while (Contador < page.Writed.Count) and (Escritura^.Line <= currentLine) do
    begin
      if Escritura^.Line = currentLine then
      begin
        LineWasPrinted := True;
        UltimaEscritura := Contador;
        InflateLineWithSpaces(lineToPrint, Escritura^.Col+Length(Escritura^.Text));
        while Columna < Escritura^.Col do
        begin
          if (lineToPrint[Columna] <> #32) and (font <> FJob.FFuente) then
          begin // PONEMOS LA FUENTE POR DEFAULT
            font := FJob.FFuente;
            PrintControlCodes(FJob.ControlCodes.Normal);
            if Bold in font then
              PrintControlCodes(FJob.ControlCodes.Bold);
            if Italic in font then
              PrintControlCodes(FJob.ControlCodes.Italic);
            if DobleWide in font then
              PrintControlCodes(FJob.ControlCodes.Wide)
            else if Compress in font then
              PrintControlCodes(FJob.ControlCodes.CondensedON)
            else
              PrintControlCodes(FJob.ControlCodes.CondensedOFF);
            if Underline in font then
              PrintControlCodes(FJob.ControlCodes.UnderlineON)
            else
              PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
          end;
          FFastPrintDevice.Write(lineToPrint[Columna]);
          Inc(Columna);
        end;
        if Escritura^.FastFont <> font then
        begin // PONEMOS LA FUENTE DEL TEXTO
          font := Escritura^.FastFont;
          PrintControlCodes(FJob.ControlCodes.Normal);
          if Bold in font then
            PrintControlCodes(FJob.ControlCodes.Bold);
          if Italic in font then
            PrintControlCodes(FJob.ControlCodes.Italic);
          if DobleWide in font then
            PrintControlCodes(FJob.ControlCodes.Wide)
          else if Compress in font then
            PrintControlCodes(FJob.ControlCodes.CondensedON)
          else
            PrintControlCodes(FJob.ControlCodes.CondensedOFF);
          if Underline in font then
            PrintControlCodes(FJob.ControlCodes.UnderlineON)
          else
            PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
        end;
        Txt := Escritura^.Text;
        if FJob.FTransliterate and (Txt<>'') then
          CharToOemBuff(PChar(@Txt[1]), PansiChar(@Txt[1]),Length(Txt));
        FFastPrintDevice.Write(Txt);
        if (Compress in font) and not(Compress in FJob.FFuente) then
        begin
          for i := 1 to Length(Escritura^.Text) do
            FFastPrintDevice.Write(#8);
          if (Length(Escritura^.Text)*6) mod 10 = 0 then
            i := Columna + (Length(Escritura^.Text) *6) div 10
          else
            i := Columna + (Length(Escritura^.Text) *6) div 10;
          font := font - [Compress];
          PrintControlCodes(FJob.ControlCodes.Normal);
          if Bold in font then
            PrintControlCodes(FJob.ControlCodes.Bold);
          if Italic in font then
            PrintControlCodes(FJob.ControlCodes.Italic);
          if DobleWide in font then
            PrintControlCodes(FJob.ControlCodes.Wide)
          else if Compress in font then
            PrintControlCodes(FJob.ControlCodes.CondensedON)
          else
            PrintControlCodes(FJob.ControlCodes.CondensedOFF);
          if Underline in font then
            PrintControlCodes(FJob.ControlCodes.UnderlineON)
          else
            PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
          while Columna <= i do
          begin
            FFastPrintDevice.Write(#32);
            Inc(Columna);
          end;
        end
        else
          Columna := Columna + Length(Escritura^.Text);
      end;
      Inc(Contador);
      if Contador < page.Writed.Count then
        Escritura := page.Writed.Items[Contador];
    end;

    if LineWasPrinted then
    begin
      if font <> FJob.FFuente then
      begin // PONEMOS LA FUENTE POR DEFAULT
        font := FJob.FFuente;
        PrintControlCodes(FJob.ControlCodes.Normal);
        if Bold in font then
          PrintControlCodes(FJob.ControlCodes.Bold);
        if Italic in font then
          PrintControlCodes(FJob.ControlCodes.Italic);
        if DobleWide in font then
          PrintControlCodes(FJob.ControlCodes.Wide)
        else if Compress in font then
          PrintControlCodes(FJob.ControlCodes.CondensedON)
        else
          PrintControlCodes(FJob.ControlCodes.CondensedOFF);
        if Underline in font then
          PrintControlCodes(FJob.ControlCodes.UnderlineON)
        else
          PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
      end;
      while Columna <= Length(lineToPrint) do
      begin
        FFastPrintDevice.Write(lineToPrint[Columna]);
        Inc(Columna);
      end;
      FFastPrintDevice.WriteLn('');
    end
    else
    begin
      if font <> FJob.FFuente then
      begin // PONEMOS LA FUENTE POR DEFAULT
        font := FJob.FFuente;
        PrintControlCodes(FJob.ControlCodes.Normal);
        if Bold in font then
          PrintControlCodes(FJob.ControlCodes.Bold);
        if Italic in font then
          PrintControlCodes(FJob.ControlCodes.Italic);
        if DobleWide in font then
          PrintControlCodes(FJob.ControlCodes.Wide)
        else if Compress in font then
          PrintControlCodes(FJob.ControlCodes.CondensedON)
        else
          PrintControlCodes(FJob.ControlCodes.CondensedOFF);
        if Underline in font then
          PrintControlCodes(FJob.ControlCodes.UnderlineON)
        else
          PrintControlCodes(FJob.ControlCodes.UnderlineOFF);
      end;
      if FJob.FTransliterate and (lineToPrint <> '') then
        AnsiToOemBuff(PansiChar(lineToPrint[1]), PansiChar(lineToPrint[1]), Length(lineToPrint));
      FFastPrintDevice.WriteLn(PChar(lineToPrint));
    end;
  end;
end;

procedure TPrintThread.SetJobName;
begin
  if FJob = nil then
    FTrayIcon.Tip :=  ''
  else
  begin
    if FPrinterStatus.PrintingPaused then
      FTrayIcon.Tip := 'Em pausa ' + FPrinterStatus.PrintingJobName
    else
      FTrayIcon.Tip := 'Imprimindo ' + FPrinterStatus.PrintingJobName;
  end;
end;

end.

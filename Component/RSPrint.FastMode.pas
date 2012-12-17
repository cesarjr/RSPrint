unit RSPrint.FastMode;

interface

uses
  Classes, RSPrint.CommonTypes, RSPrint.FastMode.FastDevice;

type
  TFastMode = class
  private const
    SINGLE_LINE = #196;
    DOUBLE_LINE = #205;

  private
    FJob: PPrintJob;
    FPrinterStatus: TPrinterStatus;
    FFastDevice: IFastDevice;

    procedure PageContinuosJump;
    procedure ProcessPauseIfNecessary;

    procedure PrintControlCodes(codeSequence: string);
    procedure PrintHorizontalLines(currentLine: Integer; page: PPage; var lineToPrint: string);
    procedure PrintVerticalLines(page: PPage; var lineToPrint: string; currentLine: Integer);
    procedure PrintCurrentLine(page: PPage; var ultimaEscritura: Integer; font: TFastFont; var lineToPrint: string; const currentLine: Integer);

    procedure PrintJob;
    function PrintPage(number: integer): Boolean;

    procedure DisposeJob;
    procedure PrintFontCodes(font: TFastFont);

  public
    constructor Create(printerStatus: TPrinterStatus);
    destructor Destroy; override;

    procedure Print(job: PPrintJob);
  end;

implementation

uses
  SysUtils, Windows, Printers, RSPrint.Utils, RSPrint.FastMode.FastDeviceFile, RSPrint.FastMode.FastDeviceSpool,
  Dialogs;

constructor TFastMode.Create(printerStatus: TPrinterStatus);
begin
  FPrinterStatus := printerStatus;

  FPrinterStatus.StartPrinting;

  FFastDevice := TFastDeviceSpool.Create;
  //FFastDevice := TFastDeviceFile.Create;
end;

destructor TFastMode.Destroy;
begin
  FPrinterStatus.CurrentlyPrinting := False;

  inherited;
end;

procedure TFastMode.Print(job: PPrintJob);
begin
  FJob := job;

  FPrinterStatus.PrintingJobName := FJob^.Name;

  ProcessPauseIfNecessary;

  PrintJob;
  DisposeJob;

  FPrinterStatus.PrintingCanceled := False;
  FPrinterStatus.PrintingJobName := '';
end;

procedure TFastMode.ProcessPauseIfNecessary;
begin
  while FPrinterStatus.PrintingPaused and not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll do
    Sleep(100);
end;

procedure TFastMode.PrintJob;
var
  Bien: Boolean;
  Copias: Integer;
  I: Integer;
begin
  Bien := True;
  for Copias := 1 to FJob^.Copias do
  begin
    FFastDevice.BeginDoc(FJob^.Name);

    for I := 1 to FJob^.LasPaginas.Count do
      if not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll and Bien then
        Bien := Bien and PrintPage(I);

    FFastDevice.EndDoc;
  end;
end;

function TFastMode.PrintPage(number: integer): boolean;
var
  CurrentLine: Integer;
  LineToPrint: string;
  Page: PPage;
  Font: TFastFont;
  UltimaEscritura: Integer;
begin
  ProcessPauseIfNecessary;

  try
    Result := True;
    FFastDevice.BeginPage;

    PrintControlCodes(FJob.ControlCodes.Reset);
    PrintControlCodes(FJob.ControlCodes.Setup);

    PrintControlCodes(FJob.ControlCodes.SelLength);
    FFastDevice.Write(#84);

    Font := FJob.DefaultFont;
    PrintFontCodes(Font);

    UltimaEscritura := 0;
    Page := FJob.LasPaginas.Items[number-1];
    CurrentLine := 1;
    while (CurrentLine < TUtils.Min(Page.PrintedLines, FJob.Lineas)) and (not FPrinterStatus.PrintingCanceled) and not (FPrinterStatus.PrintingCancelAll) do
    begin
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
      FFastDevice.Write(#12);

    FFastDevice.EndPage;
  except
    on e: exception do
    begin
      ShowMessage(e.Message);
      Result := False;
    end;
  end;
end;

procedure TFastMode.PrintFontCodes(font: TFastFont);
begin
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
end;

procedure TFastMode.PrintVerticalLines(page: PPage; var lineToPrint: string; currentLine: Integer);
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
      TUtils.InflateLineWithSpaces(lineToPrint, VerticalLine^.Col);
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

procedure TFastMode.PrintHorizontalLines(currentLine: Integer; page: PPage; var lineToPrint: string);
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
      TUtils.InflateLineWithSpaces(lineToPrint, HorizontalLine^.Col2);
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

procedure TFastMode.PrintControlCodes(codeSequence: string);
var
  CodesToPrint: TStringList;
  CodeToPrint: Byte;
  NextCode: Byte;
begin
  if Trim(codeSequence) = '' then
    exit;

  CodesToPrint := TStringList.Create;
  try
    CodesToPrint.Delimiter := ' ';
    CodesToPrint.DelimitedText := Trim(codeSequence);

    for NextCode := 0 to CodesToPrint.Count-1 do
    begin
      CodeToPrint := StrToInt(CodesToPrint[NextCode]);
      FFastDevice.Write(Chr(CodeToPrint));
    end;
  finally
    CodesToPrint.Free;
  end;
end;

procedure TFastMode.DisposeJob;
begin
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
end;

procedure TFastMode.PageContinuosJump;
var
  LinesToJump: Integer;
begin
  for LinesToJump := 1 to FJob.PageContinuousJump do
    FFastDevice.WriteLn('');
end;

procedure TFastMode.PrintCurrentLine(page: PPage; var ultimaEscritura: Integer; font: TFastFont; var lineToPrint: string; const currentLine: Integer);
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
    if font <> FJob.DefaultFont then
    begin // PONEMOS LA FUENTE POR DEFAULT
      font := FJob.DefaultFont;
      PrintFontCodes(font);
    end;

    if FJob.Transliterate and (lineToPrint <> '') then
      CharToOemBuff(PChar(@lineToPrint[1]), PansiChar(@lineToPrint[1]),Length(lineToPrint));
    FFastDevice.WriteLn(lineToPrint);
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
        TUtils.InflateLineWithSpaces(lineToPrint, Escritura^.Col+Length(Escritura^.Text));
        while Columna < Escritura^.Col do
        begin
          if (lineToPrint[Columna] <> #32) and (font <> FJob.DefaultFont) then
          begin // PONEMOS LA FUENTE POR DEFAULT
            font := FJob.DefaultFont;
            PrintFontCodes(font);
          end;

          FFastDevice.Write(lineToPrint[Columna]);
          Inc(Columna);
        end;

        if Escritura^.Font <> font then
        begin // PONEMOS LA FUENTE DEL TEXTO
          font := Escritura^.Font;
          PrintFontCodes(font);
        end;

        Txt := Escritura^.Text;
        if FJob.Transliterate and (Txt<>'') then
          CharToOemBuff(PChar(@Txt[1]), PansiChar(@Txt[1]),Length(Txt));
        FFastDevice.Write(Txt);
        if (Compress in font) and not(Compress in FJob.DefaultFont) then
        begin
          for i := 1 to Length(Escritura^.Text) do
            FFastDevice.Write(#8);
          if (Length(Escritura^.Text)*6) mod 10 = 0 then
            i := Columna + (Length(Escritura^.Text) *6) div 10
          else
            i := Columna + (Length(Escritura^.Text) *6) div 10;

          font := font - [Compress];
          PrintFontCodes(font);

          while Columna <= i do
          begin
            FFastDevice.Write(#32);
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
      if font <> FJob.DefaultFont then
      begin // PONEMOS LA FUENTE POR DEFAULT
        font := FJob.DefaultFont;
        PrintFontCodes(font);
      end;

      while Columna <= Length(lineToPrint) do
      begin
        FFastDevice.Write(lineToPrint[Columna]);
        Inc(Columna);
      end;

      FFastDevice.WriteLn('');
    end
    else
    begin
      if font <> FJob.DefaultFont then
      begin // PONEMOS LA FUENTE POR DEFAULT
        font := FJob.DefaultFont;
        PrintFontCodes(font);
      end;

      if FJob.Transliterate and (lineToPrint <> '') then
        AnsiToOemBuff(PansiChar(lineToPrint[1]), PansiChar(lineToPrint[1]), Length(lineToPrint));

      FFastDevice.WriteLn(lineToPrint);
    end;
  end;
end;

end.

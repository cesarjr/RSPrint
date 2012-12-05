unit RSPrint.PrintThread;

interface

uses
  Classes, RSPrint.RSPrintTray, RSPrint.CommonTypes, RSPrint.Utils;

type
  TPrintThread = class(TThread)
  strict private
    FJob: PPrintJob;
    FJobs: TThreadList;
    FTrayIcon: TRSPrintTray;
    FPrinterStatus: TPrinterStatus;

    procedure SetJobName;

  protected
    procedure Execute; override;
    function PrintFast(Number : integer) : boolean;

  public
    constructor Create(printerStatus: TPrinterStatus); reintroduce;
    destructor Destroy; override;

    procedure AddJob(job: PPrintJob);
  end;

implementation

uses
  SysUtils, Windows, Printers;

procedure TPrintThread.AddJob(job: PPrintJob);
begin
  FJobs.Add(job);
end;

constructor TPrintThread.Create(printerStatus: TPrinterStatus);
begin
  inherited Create(True);

  FJobs := TThreadList.Create;
  FTrayIcon := TRSPrintTray.Create(nil, FPrinterStatus);
  FPrinterStatus := printerStatus;

  FreeOnTerminate := True;
  FPrinterStatus.StartPrinting;
end;

destructor TPrintThread.Destroy;
begin
  FJobs.Free;
  FPrinterStatus.CurrentlyPrinting := False;

  inherited;
end;

procedure TPrintThread.Execute;
var
  Bien : boolean;
  Copias,i : integer;
  List : TList;
  Cant : integer;
begin
  List := FJobs.LockList;
  Cant := List.Count;
  if Cant > 0 then
    begin
      FJob := pPrintJob(List[0]);
      FPrinterStatus.PrintingJobName := FJob^.Name;
      Synchronize(SetJobName);
    end
  else
    begin
      FJob := nil;
      FPrinterStatus.PrintingJobName := '';
      Synchronize(SetJobName);
    end;
  FJobs.UnlockList;
  While FPrinterStatus.PrintingPaused and not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll do
    Sleep(100);
  While (FJob <> nil) do
    begin
      with FJob^ do
        begin
          Bien := True;
          for Copias := 1 to FCopias do
            for i := 1 to LasPaginas.Count do
              if not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll and Bien then
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
      Dispose(FJob);
      FPrinterStatus.PrintingCanceled := False;
      List := FJobs.LockList;
      List.Delete(0);
      Cant := List.Count;
      if Cant > 0 then
        begin
          FJob := pPrintJob(List[0]);
          FPrinterStatus.PrintingJobName := FJob^.Name;
          Synchronize(SetJobName);
        end
      else
        begin
          FJob := nil;
          FPrinterStatus.PrintingJobName := '';
          Synchronize(SetJobName);
        end;
    end;
  FPrinterStatus.PrintingJobName := '';
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
            TUtils.ToPrn(chr(Cod));
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
            TUtils.ToPrn(chr(Cod));
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
        if Fuente <> FJob.FFuente then
          begin // PONEMOS LA FUENTE POR DEFAULT
            Fuente := FJob.FFuente;
            ImprimirCodigo(FJob.PRNNormal);
            if Bold in Fuente then
              ImprimirCodigo(FJob.PRNBold);
            if Italic in Fuente then
              ImprimirCodigo(FJob.PRNItalics);
            if DobleWide in Fuente then
              ImprimirCodigo(FJob.PRNWide)
            else if Compress in Fuente then
              ImprimirCodigo(FJob.PRNCompON)
            else
              ImprimirCodigo(FJob.PRNCompOFF);
            if Underline in Fuente then
              ImprimirCodigo(FJob.PRNULineON)
            else
              ImprimirCodigo(FJob.PRNULineOFF);
          end;
        if FJob.FTransliterate and (LineaAImprimir<>'') then
          CharToOemBuff(PChar(@LineaAImprimir[1]), PansiChar(@LineaAImprimir[1]),Length(LineaAImprimir));
(*Verificar*)
        TUtils.ToPrnLn(LineaAImprimir);
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
                       (Fuente <> FJob.FFuente) then
                      begin // PONEMOS LA FUENTE POR DEFAULT
                        Fuente := FJob.FFuente;
                        ImprimirCodigo(FJob.PRNNormal);
                        if Bold in Fuente then
                          ImprimirCodigo(FJob.PRNBold);
                        if Italic in Fuente then
                          ImprimirCodigo(FJob.PRNItalics);
                        if DobleWide in Fuente then
                          ImprimirCodigo(FJob.PRNWide)
                        else if Compress in Fuente then
                          ImprimirCodigo(FJob.PRNCompON)
                        else
                          ImprimirCodigo(FJob.PRNCompOFF);
                        if Underline in Fuente then
                          ImprimirCodigo(FJob.PRNULineON)
                        else
                          ImprimirCodigo(FJob.PRNULineOFF);
                      end;
(*Verificar*)
                    TUtils.ToPrn(LineaAImprimir[Columna]);
                 //   Write(Impresora,LineaAImprimir[Columna]);
                    Inc(Columna);
                  end;
                 if Escritura^.FastFont <> Fuente then
                  begin // PONEMOS LA FUENTE DEL TEXTO
                    Fuente := Escritura^.FastFont;
                    ImprimirCodigo(FJob.PRNNormal);
                    if Bold in Fuente then
                      ImprimirCodigo(FJob.PRNBold);
                    if Italic in Fuente then
                      ImprimirCodigo(FJob.PRNItalics);
                    if DobleWide in Fuente then
                      ImprimirCodigo(FJob.PRNWide)
                    else if Compress in Fuente then
                      ImprimirCodigo(FJob.PRNCompON)
                    else
                      ImprimirCodigo(FJob.PRNCompOFF);
                    if Underline in Fuente then
                      ImprimirCodigo(FJob.PRNULineON)
                    else
                      ImprimirCodigo(FJob.PRNULineOFF);
                  end;
                Txt := Escritura^.Text;
                if FJob.FTransliterate and (Txt<>'') then
                  CharToOemBuff(PChar(@Txt[1]), PansiChar(@Txt[1]),Length(Txt));
(*Verificar*)
                TUtils.ToPrn(Txt);
                //Write(Impresora,Txt);
                if (Compress in Fuente) and not(Compress in FJob.FFuente) then
                  begin
(*Verificar*)
                    for i := 1 to Length(Escritura^.Text) do
                      TUtils.ToPrn(#8);
                      //Write(Impresora,#8);
                    if (Length(Escritura^.Text)*6) mod 10 = 0 then
                      i := Columna + (Length(Escritura^.Text) *6) div 10
                    else
                      i := Columna + (Length(Escritura^.Text) *6) div 10;
                    Fuente := Fuente - [Compress];
                    ImprimirCodigo(FJob.PRNNormal);
                    if Bold in Fuente then
                      ImprimirCodigo(FJob.PRNBold);
                    if Italic in Fuente then
                      ImprimirCodigo(FJob.PRNItalics);
                    if DobleWide in Fuente then
                      ImprimirCodigo(FJob.PRNWide)
                    else if Compress in Fuente then
                      ImprimirCodigo(FJob.PRNCompON)
                    else
                      ImprimirCodigo(FJob.PRNCompOFF);
                    if Underline in Fuente then
                      ImprimirCodigo(FJob.PRNULineON)
                    else
                      ImprimirCodigo(FJob.PRNULineOFF);
                    While Columna <= i do
                      begin
(*Verificar*)
                        TUtils.ToPrn(#32);
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
            if Fuente <> FJob.FFuente then
              begin // PONEMOS LA FUENTE POR DEFAULT
                Fuente := FJob.FFuente;
                ImprimirCodigo(FJob.PRNNormal);
                if Bold in Fuente then
                  ImprimirCodigo(FJob.PRNBold);
                if Italic in Fuente then
                  ImprimirCodigo(FJob.PRNItalics);
                if DobleWide in Fuente then
                  ImprimirCodigo(FJob.PRNWide)
                else if Compress in Fuente then
                  ImprimirCodigo(FJob.PRNCompON)
                else
                  ImprimirCodigo(FJob.PRNCompOFF);
                if Underline in Fuente then
                  ImprimirCodigo(FJob.PRNULineON)
                else
                  ImprimirCodigo(FJob.PRNULineOFF);
              end;
            While Columna <= Length(LineaAImprimir) do
              begin
(*Verificar*)
                TUtils.ToPrn(LineaAImprimir[Columna]);
                //Write(Impresora,LineaAImprimir[Columna]);
                Inc(Columna);
              end;
(*Verificar*)
            TUtils.ToPrnLn('');
            //WriteLn(Impresora);
          end
        else
          begin
            if Fuente <> FJob.FFuente then
              begin // PONEMOS LA FUENTE POR DEFAULT
                Fuente := FJob.FFuente;
                ImprimirCodigo(FJob.PRNNormal);
                if Bold in Fuente then
                  ImprimirCodigo(FJob.PRNBold);
                if Italic in Fuente then
                  ImprimirCodigo(FJob.PRNItalics);
                if DobleWide in Fuente then
                  ImprimirCodigo(FJob.PRNWide)
                else if Compress in Fuente then
                  ImprimirCodigo(FJob.PRNCompON)
                else
                  ImprimirCodigo(FJob.PRNCompOFF);
                if Underline in Fuente then
                  ImprimirCodigo(FJob.PRNULineON)
                else
                  ImprimirCodigo(FJob.PRNULineOFF);
              end;
            if FJob.FTransliterate and (LineaAImprimir<>'') then
              AnsiToOemBuff(PansiChar(LineaAImprimir[1]), PansiChar(LineaAImprimir[1]),Length(LineaAImprimir));
(*Verificar*)
            TUtils.ToPrnLn(pchar(LineaAImprimir));
            //Writeln(Impresora,LineaAImprimir);
          end;
      end;
  end;

begin
  ListaImpressoras:= tstringlist.Create;
  TUtils.EnumPrt(ListaImpressoras, PrinterId);

  While FPrinterStatus.PrintingPaused and not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll do
    Sleep(100);
     try
       Resultado := True;
     //  AssignFile(Impresora,Job.FPort);

(*Verificar*)
//       ReWrite(Impresora);
       TUtils.StartPrint(ListaImpressoras[printer.Printerindex],'TESTE DE IMPRESSAO','',1);

       ImprimirCodigo(FJob.PRNReset);
       ImprimirCodigo(FJob.PRNSetup);

           ImprimirCodigo(FJob.PRNSelLength);
(*Verificar*)
//           Write(Impresora,#84);
       TUtils.ToPrn(#84);


//         end;
       Fuente := FJob.FFuente;
       ImprimirCodigo(FJob.PRNNormal);
       if Bold in Fuente then
         ImprimirCodigo(FJob.PRNBold);
       if Italic in Fuente then
         ImprimirCodigo(FJob.PRNItalics);
       if DobleWide in Fuente then
         ImprimirCodigo(FJob.PRNWide)
       else if Compress in Fuente then
         ImprimirCodigo(FJob.PRNCompON)
       else
         ImprimirCodigo(FJob.PRNCompOFF);
       if Underline in Fuente then
         ImprimirCodigo(FJob.PRNULineON)
       else
         ImprimirCodigo(FJob.PRNULineOFF);
       UltimaEscritura := 0;
       Pagina := FJob.LasPaginas.Items[Number-1];
       LA := 1;
       while (LA < TUtils.Min(Pagina.PrintedLines, FJob.FLineas)) and (not FPrinterStatus.PrintingCanceled) and not (FPrinterStatus.PrintingCancelAll) do
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
             While FPrinterStatus.PrintingPaused and not FPrinterStatus.PrintingCanceled and not FPrinterStatus.PrintingCancelAll do
               Sleep(100);
           end;

(* Cria um salto de picote manual selecionavel nas propriedades da rsprint *)
         if (FJob.PageSize = pzContinuous) and (FJob.PageLength = 0) then
           begin
             for j := 1 to FJob.PageContinuousJump do
               TUtils.ToPrnLn('');
              // WriteLn(Impresora,'');
           end
         else
(*Verificar*)
//           Write(Impresora,#12);
         TUtils.ToPrn(#12);
     except
       Resultado := False;
     end;

(*Verificar*)
  TUtils.EndPrint;
//      {$I-}
//      CloseFile(Impresora);
//      {$I+}
  PrintFast := Resultado;
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

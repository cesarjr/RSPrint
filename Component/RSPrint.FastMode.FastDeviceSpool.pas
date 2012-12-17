unit RSPrint.FastMode.FastDeviceSpool;

interface

uses
  RSPrint.FastMode.FastDevice, Windows, Classes;

type
  TFastDeviceSpool = class(TInterfacedObject, IFastDevice)

  private
    FPrinterHandle: THandle;
    FDevMode: TDeviceModeA;
    FPrinterJob: Dword;

    procedure EnumPrt(st: TStrings; var def: Integer);
    procedure StartPrint(prtName, docName: string; copies: Integer);

  public
    procedure BeginDoc;
    procedure BeginPage;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndPage;
    procedure EndDoc;
  end;


implementation

uses
  Printers, WinSpool;

procedure TFastDeviceSpool.BeginDoc;
var
  ListaImpressoras: TStringList;
  PrinterId: Integer;
begin
  ListaImpressoras := TStringList.Create;
  try
    EnumPrt(ListaImpressoras, PrinterId);
    StartPrint(ListaImpressoras[Printer.PrinterIndex], 'TESTE DE IMPRESSAO', 1);
  finally
    ListaImpressoras.Free;
  end;
end;

procedure TFastDeviceSpool.EnumPrt(st: TStrings; var def: Integer);
type
  PPrInfoArr = ^TPrInfoArr;
  TPrInfoArr = array [0..0] of TPRINTERINFO2;
var
  i: Integer;
  Indx: Integer;
  Level: Integer;
  buf: Pointer;
  Need: Dword;
  Returned: Dword;
  PrInfoArr: PPrInfoArr;
begin
  st.Clear;
  Def := 0;
  Level := 2;
  EnumPrinters(PRINTER_ENUM_LOCAL, nil, Level, nil, 0, Need, Returned);
  GetMem(buf, Need);

  try
    EnumPrinters(PRINTER_ENUM_LOCAL, nil, Level, PByte(buf), Need, Need, Returned);
    PrInfoArr := buf;

    {$RANGECHECKS OFF}
    for i:=0 to Returned-1 do
    begin
      Indx := st.Add(PrInfoArr[i].pPrinterName);
      if (PrInfoArr[i].Attributes AND PRINTER_ATTRIBUTE_DEFAULT) > 0 then
        Def := Indx;
    end;
    {$RANGECHECKS ON}
  finally
    FreeMem(buf);
  end;
end;

procedure TFastDeviceSpool.StartPrint(prtName, docName: string; copies: Integer);
var
  DocInfo: PDocInfo1;
begin
  FDevMode.dmCopies := Copies;
  FDevMode.dmFields := DM_COPIES;

  if OpenPrinter(PChar(PrtName), FPrinterHandle, nil) then
  begin
    New(DocInfo);
    DocInfo^.pDocName := PChar(DocName);
    DocInfo^.pDatatype := 'RAW';
    DocInfo^.pOutputFile := nil;

    FPrinterJob := StartDocPrinter(FPrinterHandle, 1, DocInfo);
  end;
end;

procedure TFastDeviceSpool.BeginPage;
begin
  StartPagePrinter(FPrinterHandle);
end;

procedure TFastDeviceSpool.WriteLn(value: string);
begin
  Write(value + #13#10);
end;

procedure TFastDeviceSpool.Write(value: string);
var
  BytesWritten: DWORD;
  DataToPrint: AnsiString;
begin
  DataToPrint := AnsiString(value);

  WritePrinter(FPrinterHandle, PAnsiChar(DataToPrint), Length(DataToPrint), BytesWritten);
end;

procedure TFastDeviceSpool.EndPage;
begin
  EndPagePrinter(FPrinterHandle);
end;

procedure TFastDeviceSpool.EndDoc;
begin
  EndDocPrinter(FPrinterHandle);
end;

end.

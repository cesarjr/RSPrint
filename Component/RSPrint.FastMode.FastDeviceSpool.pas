unit RSPrint.FastMode.FastDeviceSpool;

interface

uses
  RSPrint.FastMode.FastDevice, Windows, Classes;

type
  TFastDeviceSpool = class(TInterfacedObject, IFastDevice)

  private
    FPrinterHandle: THandle;

    function GetDefaultPrinterName: string;

  public
    procedure BeginDoc(documentName: string);
    procedure BeginPage;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndPage;
    procedure EndDoc;
  end;


implementation

uses
  Printers, WinSpool;

procedure TFastDeviceSpool.BeginDoc(documentName: string);
var
  Document: TDocInfo1;
  PrinterName: string;
begin
  PrinterName := GetDefaultPrinterName;

  if OpenPrinter(PChar(PrinterName), FPrinterHandle, nil) then
  begin
    Document.pDocName := PChar(documentName);
    Document.pDatatype := 'RAW';
    Document.pOutputFile := nil;

    StartDocPrinter(FPrinterHandle, 1, @Document);
  end;
end;

function TFastDeviceSpool.GetDefaultPrinterName: string;
begin
  if Printer.PrinterIndex > -1 then
    Result := Printer.Printers[Printer.PrinterIndex]
  else
    Result := '';
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

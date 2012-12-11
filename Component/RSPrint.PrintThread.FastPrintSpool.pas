unit RSPrint.PrintThread.FastPrintSpool;

interface

uses
  RSPrint.PrintThread.FastPrintDevice;

type
  TFastPrintSpool = class(TInterfacedObject, IFastPrintDevice)
  public
    procedure BeginDoc;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndDoc;
  end;


implementation

uses
  RSPrint.Utils, Classes, Printers;

procedure TFastPrintSpool.BeginDoc;
var
  ListaImpressoras: TStringList;
  PrinterId: Integer;
begin
  ListaImpressoras := TStringList.Create;
  try
    TUtils.EnumPrt(ListaImpressoras, PrinterId);
    TUtils.StartPrint(ListaImpressoras[Printer.PrinterIndex],'TESTE DE IMPRESSAO','',1);
  finally
    ListaImpressoras.Free;
  end;
end;

procedure TFastPrintSpool.EndDoc;
begin
  TUtils.EndPrint;
end;

procedure TFastPrintSpool.Write(value: string);
begin
  TUtils.ToPrn(value);
end;

procedure TFastPrintSpool.WriteLn(value: string);
begin
  TUtils.ToPrnLn(value);
end;

end.

unit RSPrint.PrintThread.FastPrintDevice;

interface

type
  IFastPrintDevice = interface
    procedure BeginDoc;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndDoc;
  end;

implementation

end.

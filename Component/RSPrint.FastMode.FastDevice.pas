unit RSPrint.FastMode.FastDevice;

interface

type
  IFastDevice = interface
    procedure BeginDoc;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndDoc;
  end;

implementation

end.

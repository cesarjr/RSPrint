unit RSPrint.FastMode.FastDevice;

interface

type
  IFastDevice = interface
    procedure BeginDoc(documentName: string);
    procedure BeginPage;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndPage;
    procedure EndDoc;
  end;

implementation

end.

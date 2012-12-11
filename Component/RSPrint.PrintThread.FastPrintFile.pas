unit RSPrint.PrintThread.FastPrintFile;

interface

uses
  RSPrint.PrintThread.FastPrintDevice;

type

  TFastPrintFile = class(TInterfacedObject, IFastPrintDevice)
  private
    FFile: TextFile;

  public
    procedure BeginDoc;
    procedure Write(value: string);
    procedure WriteLn(value: string);
    procedure EndDoc;
  end;

implementation

uses
  SysUtils;

procedure TFastPrintFile.BeginDoc;
var
  FileName: string;
begin
  FileName := 'rsprint.debug.' + FormatDateTime('yyyy.mm.dd.hh.nn.ss.zzz', now) + '.txt';

  AssignFile(FFile, FileName);
  ReWrite(FFile);
end;

procedure TFastPrintFile.EndDoc;
begin
  CloseFile(FFile);
end;

procedure TFastPrintFile.Write(value: string);
begin
  System.Write(FFile, value);
end;

procedure TFastPrintFile.WriteLn(value: string);
begin
  System.WriteLn(FFile, value);
end;

end.

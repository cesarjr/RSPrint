unit RSPrint.FastMode.FastDeviceFile;

interface

uses
  RSPrint.FastMode.FastDevice;

type

  TFastDeviceFile = class(TInterfacedObject, IFastDevice)
  private
    FFile: TextFile;

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
  SysUtils;

procedure TFastDeviceFile.BeginDoc;
var
  FileName: string;
begin
  FileName := 'rsprint.debug.' + FormatDateTime('yyyy.mm.dd.hh.nn.ss.zzz', now) + '.txt';

  AssignFile(FFile, FileName);
  ReWrite(FFile);
end;

procedure TFastDeviceFile.BeginPage;
begin
  {empty by purpose}
end;

procedure TFastDeviceFile.EndDoc;
begin
  CloseFile(FFile);
end;

procedure TFastDeviceFile.EndPage;
begin
  {empty by purpose}
end;

procedure TFastDeviceFile.Write(value: string);
begin
  System.Write(FFile, value);
end;

procedure TFastDeviceFile.WriteLn(value: string);
begin
  System.WriteLn(FFile, value);
end;

end.

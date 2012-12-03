unit Utils;
{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, Classes, WinSpool, SysUtils;

type

  TUtils = class

  private
    class var
      FPrinterHandle: THandle;
      FDevMode: TDeviceModeA;
      FPrinterJob: Dword;

  public
    class procedure EnumPrt(st: TStrings; var def: Integer);
    class procedure StartPrint(prtName, docName, toFile: string; copies: Integer);
    class procedure EndPrint;
    class function ToPrn(s: string): Boolean;
    class function ToPrnLn(s: string): Boolean;
    class function Min(Val1, Val2: Integer): Integer;
  end;

implementation

uses
  StrUtils;

class function TUtils.Min(Val1, Val2: Integer): Integer;
begin
  if Val1<Val2 then
    Min := Val1
  else
    Min := Val2;
end;

class procedure TUtils.EnumPrt(st: TStrings; var def: Integer);
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
  Def:=0;
  Level:=2;
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

class procedure TUtils.StartPrint(prtName, docName, toFile: string; copies: Integer);
var
  pdi: PDocInfo1;
  pd: TPrinterDefaults;
begin
  FDevMode.dmCopies := Copies;
  FDevMode.dmFields:=DM_COPIES;
  pd.pDatatype:='RAW';
  pd.pDevMode:=@FDevMode;
  pd.DesiredAccess:=PRINTER_ACCESS_USE;

  if Win32Check(OpenPrinter(PChar(PrtName),FPrinterHandle,@pd)) then
  begin
    new(pdi);
    with pdi^ do
    begin
      pDocName:=PChar(DocName);

      if ToFile = '' then
        pOutputFile := nil
      else
        pOutputFile:=PChar(ToFile);

      pDatatype:='RAW';
    end;

    FPrinterJob := StartDocPrinter(FPrinterHandle, 1, pdi);
    if FPrinterJob = 0 then
      Win32Check(false);
  end;
end;

class function TUtils.ToPrnLn(s: string): Boolean;
begin
  Result := TUtils.ToPrn(s+#13#10);
end;

class function TUtils.ToPrn(s: string): Boolean;
var
  cp: Dword;
begin
  Win32Check(WritePrinter(FPrinterHandle, PChar(s), Length(s), cp));
  Result := True;
end;

class procedure TUtils.EndPrint;
begin
  Win32Check(EndDocPrinter(FPrinterHandle));
end;

initialization

end.

unit SplPrint;

interface

uses
  Windows,
  Messages,
  Classes,
  WinSpool,
  SysUtils;

  procedure EnumPrt(st: TStrings;var Def: integer);
  procedure StartPrint(PrtName, DocName, ToFile: string; Copies: integer);
  function ToPrnFrmC(FrmStr: string; const Args: array of const): boolean;
  function ToPrnFrm(FrmStr: string; const Args: array of const): boolean;
  function ToPrnLn(S: string): boolean;
  function ToPrn(S: string): boolean;
  procedure CancelPrint;
  procedure EndPrint;

var
  ph: THandle;
  DevMode: TDeviceModeA;
  PrJob: dword;

implementation

function ReplaceStr(const S, Srch, Replace: string): string;
var
  I: Integer;
  Source: string;
begin
  Source := S;
  Result := '';
  repeat
    I := Pos(Srch, Source);
    if I > 0 then begin
      Result := Result + Copy(Source, 1, I - 1) + Replace;
      Source := Copy(Source, I + Length(Srch), MaxInt);
    end
    else Result := Result + Source;
  until I <= 0;
end;

function ToDos(const AnsiStr: ansiString):ansiString;
begin
  SetLength(Result, Length(AnsiStr));
  if Length(Result) > 0 then
    AnsiToOem(PansiChar(AnsiStr), PansiChar(Result));
  Result:=ReplaceStr(Result,chr(253),chr(15));
end;

function ToWin(str: ansiString):ansiString;
begin
  if str='' then begin
    Result:='';
    System.Exit;
  end;
  OemToAnsi(PansiChar(str),PansiChar(str));
  Result:=str;
end;

procedure EnumPrt(st: TStrings;var Def: integer);
type
  PPrInfoArr = ^TPrInfoArr;
  TPrInfoArr = array [0..0] of TPRINTERINFO2;
var
  i,Indx,Level: integer;
  buf: pointer;
  Need,Returned: dword;
  PrInfoArr: PPrInfoArr;
begin
  st.Clear;
  Def:=0; Level:=2;
  EnumPrinters(PRINTER_ENUM_LOCAL,nil,Level,nil,0,Need,Returned);
  GetMem(buf,Need);
  try
    EnumPrinters(PRINTER_ENUM_LOCAL,nil,Level,PByte(buf),Need,Need,Returned);
    PrInfoArr:=buf;
    {$RANGECHECKS OFF}
    for i:=0 to Returned-1 do begin
      Indx:=st.Add(PrInfoArr[i].pPrinterName);
      if (PrInfoArr[i].Attributes AND PRINTER_ATTRIBUTE_DEFAULT)>0 then Def:=Indx;
    end;
    {$RANGECHECKS ON}
  finally
    FreeMem(buf);
  end;
end;

procedure StartPrint(PrtName, DocName, ToFile: string; Copies: integer);
var
  pdi: PDocInfo1;
  pd: TPrinterDefaults;
begin
  DevMode.dmCopies:=Copies;
  DevMode.dmFields:=DM_COPIES;
  pd.pDatatype:='RAW';
  pd.pDevMode:=@DevMode;
  pd.DesiredAccess:=PRINTER_ACCESS_USE;
  if Win32Check(OpenPrinter(PChar(PrtName),ph,@pd)) then begin
    new(pdi);
    with pdi^ do begin
      pDocName:=PChar(DocName);
      if ToFile='' then pOutputFile:=nil
      else pOutputFile:=PChar(ToFile);
      pDatatype:='RAW';
    end;
    PrJob:=StartDocPrinter(ph,1,pdi);
    if PrJob=0 then Win32Check(false);
  end;
end;

function ToPrnFrm(FrmStr: string; const Args: array of const): boolean;
begin
  Result:=ToPrnLn(Format(FrmStr,Args));
end;

function ToPrnFrmC(FrmStr: string; const Args: array of const): boolean;
begin
  Result:=ToPrnLn(ToDos(Format(FrmStr,Args)));
end;

function ToPrnLn(S: string): boolean;
begin
  Result:=ToPrn(S+#13#10);
end;

function ToPrn(S: string): boolean;
var cp: dword;
begin
  Win32Check(WritePrinter(ph,PChar(S),length(S),cp));
  Result:=true;
end;

procedure EndPrint;
begin
  Win32Check(EndDocPrinter(ph));
end;

procedure CancelPrint;
begin
  Win32Check(SetJob(ph,PrJob,0,nil,JOB_CONTROL_CANCEL));
end;

initialization

end.

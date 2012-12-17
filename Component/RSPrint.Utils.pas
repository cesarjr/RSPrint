unit RSPrint.Utils;

interface

uses
  Windows, Messages, Classes, SysUtils;

type

  TUtils = class
  public
    class procedure InflateLineWithSpaces(var line: string; maxCol: Byte); static;
    class function Min(Val1, Val2: Integer): Integer;

  end;

implementation

uses
  StrUtils, WinSpool;

class function TUtils.Min(Val1, Val2: Integer): Integer;
begin
  if Val1<Val2 then
    Min := Val1
  else
    Min := Val2;
end;

class procedure TUtils.InflateLineWithSpaces(var line: string; maxCol: Byte);
begin
  while Length(line) < maxCol do
    line := line + ' ';
end;

initialization

end.

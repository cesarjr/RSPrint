unit RSPrint.Utils;

interface

uses
  Windows, Messages, Classes, SysUtils;

type

  TUtils = class
  public
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

initialization

end.

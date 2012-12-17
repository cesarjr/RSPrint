unit RSPrint.Types.Job;

interface

uses
  RSPrint.CommonTypes, Generics.Collections, RSPrint.Types.Page;

type
  TJob = class
  public
    Name: string;
    PageSize: TPageSize;
    PageContinuousJump: Byte;
    PageLength: Byte;
    Pages: TObjectList<TPage>;
    Copias: Integer;
    DefaultFont: TFastFont;
    Lineas: Integer;
    Transliterate: Boolean;
    ControlCodes: TControlCodes;

    constructor Create;
    destructor Destroy; override;

    procedure ImportPage(origin: TPage);
  end;

implementation

constructor TJob.Create;
begin
  Pages := TObjectList<TPage>.Create(True);
end;

destructor TJob.Destroy;
begin
  Pages.Free;

  inherited;
end;

procedure TJob.ImportPage(origin: TPage);
var
  CopiedPage: TPage;
begin
  CopiedPage := TPage.Create;
  CopiedPage.CopyFrom(origin);

  Pages.Add(CopiedPage);
end;

end.

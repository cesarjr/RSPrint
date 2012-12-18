unit RSPrint.Types.Document;

interface

uses
  RSPrint.CommonTypes, Generics.Collections, RSPrint.Types.Page;

type
  TDocument = class
  private
    FCurrentPageIndex: Integer;

    function GetCurrentPageNumber: Integer;
    function GetCurrentPage: TPage;

  public
    Title: string;
    PageSize: TPageSize;
    PageContinuousJump: Byte;
    PageLength: Byte;
    Pages: TObjectList<TPage>;
    Copies: Integer;
    DefaultFont: TFastFont;
    LinesPerPage: Integer;
    Transliterate: Boolean;
    ControlCodes: TControlCodes;

    constructor Create;
    destructor Destroy; override;

    procedure ClearPages;
    procedure NewPage;

    property CurrentPageNumber: Integer read GetCurrentPageNumber;
    property CurrentPage: TPage read GetCurrentPage;
  end;

implementation

procedure TDocument.ClearPages;
begin
  Pages.Free;
  Pages := TObjectList<TPage>.Create(True);
end;

constructor TDocument.Create;
begin
  Pages := TObjectList<TPage>.Create(True);
end;

destructor TDocument.Destroy;
begin
  Pages.Free;

  inherited;
end;

function TDocument.GetCurrentPage: TPage;
begin
  Result := Pages.Items[FCurrentPageIndex];
end;

function TDocument.GetCurrentPageNumber: Integer;
begin
  Result := FCurrentPageIndex + 1;
end;

procedure TDocument.NewPage;
var
  NewPage: TPage;
begin
  NewPage := TPage.Create;
  Pages.Add(NewPage);

  FCurrentPageIndex  := Pages.IndexOf(NewPage);
end;

end.

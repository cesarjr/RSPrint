unit RSPrint.Types.Img;

interface

uses
  Graphics;

type
  TImg = class
    Col: Single;
    Line: Single;
    Picture: TPicture;

    constructor Create;
    destructor Destroy; override;
  end;

implementation

constructor TImg.Create;
begin
  Picture := TPicture.Create;
end;

destructor TImg.Destroy;
begin
  Picture.Free;

  inherited;
end;

end.

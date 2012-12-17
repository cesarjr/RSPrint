unit RSPrint.Types.Page;

interface

uses
  Classes, Generics.Collections, RSPrint.CommonTypes, RSPrint.Types.Img;

type

  TWrittenText = record
    Col: Byte;
    Line: Byte;
    Font: TFastFont;
    Text: string;
  end;

  THorizontalLine = record
    Col1: Byte;
    Col2: Byte;
    Line: Byte;
    Kind: TLineType;
  end;

  TVerticalLine = record
    Col: Byte;
    Line1: Byte;
    Line2: Byte;
    Kind: TLineType;
  end;

  TPage = class
  public
    WrittenText: TList<TWrittenText>;
    VerticalLines: TList<TVerticalLine>;
    HorizontalLines: TList<THorizontalLine>;
    PrintedLines: Byte;
    Images: TList<TImg>;

    constructor Create;
    destructor Destroy; override;

    procedure CopyFrom(origin: TPage);
  end;

implementation

constructor TPage.Create;
begin
  WrittenText := TList<TWrittenText>.Create;
  VerticalLines := TList<TVerticalLine>.Create;
  HorizontalLines := TList<THorizontalLine>.Create;
  Images := TObjectList<TImg>.Create(True);
end;

destructor TPage.Destroy;
begin
  WrittenText.Free;
  VerticalLines.Free;
  HorizontalLines.Free;
  Images.Free;

  inherited;
end;

procedure TPage.CopyFrom(origin: TPage);
var
  I: Integer;
begin
  PrintedLines := origin.PrintedLines + 1;

  for I := 0 to origin.WrittenText.Count-1 do
    WrittenText.Add(origin.WrittenText[I]);

  for I := 0 to origin.VerticalLines.Count-1 do
    VerticalLines.Add(origin.VerticalLines[I]);

  for I := 0 to origin.HorizontalLines.Count-1 do
    HorizontalLines.Add(origin.HorizontalLines[I]);
end;

end.

unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, RSPrint;

type
  TFrmMain = class(TForm)
    btnPreview: TButton;
    RSPrint: TRSPrinter;
    btnPrint: TButton;
    cmbMode: TComboBox;
    lblMode: TLabel;
    procedure btnPreviewClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);

  private
    procedure GenerateData;
    procedure PrintHeader;

  end;

var
  FrmMain: TFrmMain;

implementation

uses
  RSPrint.CommonTypes;

{$R *.dfm}

procedure TFrmMain.btnPreviewClick(Sender: TObject);
begin
  GenerateData;
  RSPrint.PreviewReal;
end;

procedure TFrmMain.btnPrintClick(Sender: TObject);
begin
  GenerateData;
  RSPrint.PrintAll;
end;

procedure TFrmMain.GenerateData;
var
  I: Integer;
  CurrentLine: Integer;
begin
  RSPrint.Mode := TPrinterMode(cmbMode.ItemIndex);
  RSPrint.ReportName := 'RSPrint Demo App';

  RSPrint.BeginDoc;

  PrintHeader;
  CurrentLine := 5;

  for I := 1 to 500 do
  begin
    RSPrint.Write(CurrentLine,  1, Format('%4.4d', [I]));
    RSPrint.Write(CurrentLine, 10, 'Record Name Here');

    CurrentLine := CurrentLine + 1;

    if CurrentLine > RSPrint.Lines then
    begin
      RSPrint.NewPage;
      PrintHeader;
      CurrentLine := 5;
    end;
  end;
end;

procedure TFrmMain.PrintHeader;
begin
  RSPrint.WriteFont(1,  1, 'Report Example',[Bold]);
  RSPrint.WriteFont(1, 72, Format('Page: %2.2d', [RSPrint.PageNo]), [Bold]);
  RSPrint.Write(2,  1, StringOfChar('-', 80));

  RSPrint.Write(3,  1, 'ID');
  RSPrint.Write(3, 10, 'Name');
  RSPrint.Write(4,  1, StringOfChar('-', 80));
end;

end.

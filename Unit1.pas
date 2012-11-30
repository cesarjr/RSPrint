unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RSPrint;

type
  TForm1 = class(TForm)
    Button1: TButton;
    RS: TRSPrinter;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
   RS.BeginDoc;
   RS.WriteFont(1,1,'Relatorio Geral',[Italic]);
   RS.Write(2,1,'Codigo');
   RS.Write(2,10,'Nome');
   RS.Write(3,3,'0001');
   RS.Write(3,10,'Wilson');
   RS.PreviewReal;
end;

end.

unit RSPrint.Preview.PrinterSetup;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, ExtCtrls, Buttons, Printers;

type

  TImpTxt= TStrings;

  TFrmPrinterSetup = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ListaImpressoras: TComboBox;
    Label1: TLabel;
    PrinterSetupDialog: TPrinterSetupDialog;
    BtnPropriedades: TBitBtn;
    WinPrinters: TComboBox;
    Labelmodelos: TLabel;
    procedure BtnPropriedadesClick(Sender: TObject);
    procedure ListaImpressorasChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure WinPrintersChange(Sender: TObject);

  private
    procedure SetaModo;

  public
    ImpressorasTexto:tstrings;
    ImpTxtSel:string;

  end;

var
  FrmPrinterSetup: TFrmPrinterSetup;

implementation

{$R *.DFM}

procedure TFrmPrinterSetup.BtnPropriedadesClick(Sender: TObject);
var
  PrnSet: TPrinterSetupDialog;
begin
  PrnSet := TPrinterSetupDialog.Create(self);
  PrnSet.Execute;
  PrnSet.Free;
end;

procedure TFrmPrinterSetup.ListaImpressorasChange(Sender: TObject);
begin
  ImpTxtSel := ListaImpressoras.Items[ListaImpressoras.ItemIndex];
end;

procedure TFrmPrinterSetup.FormCreate(Sender: TObject);
begin
  ImpressorasTexto := TstringList.Create;
end;

procedure TFrmPrinterSetup.SetaModo;
begin
  if (printer.Canvas.TextHeight('X') = 1) or (printer.Canvas.TextHeight('X') = 83) then
  begin
    ListaImpressoras.Visible := True;
    Labelmodelos.Visible := True;
    BtnPropriedades.Enabled := False;
    BtnPropriedades.Visible := False;
  end
  else
  begin
    BtnPropriedades.Enabled := True;
    BtnPropriedades.Visible := True;
    ListaImpressoras.Visible := False;
    Labelmodelos.Visible := False;
  end;
end;

procedure TFrmPrinterSetup.FormShow(Sender: TObject);
begin
  listaimpressoras.Items := impressorastexto;
  listaimpressoras.Text := ImpTxtSel;
  WinPrinters.Items := printer.Printers;
  winprinters.ItemIndex := printer.PrinterIndex;
  setamodo;
end;

procedure TFrmPrinterSetup.WinPrintersChange(Sender: TObject);
begin
  printer.PrinterIndex := winprinters.ItemIndex;
  setamodo;
end;

end.

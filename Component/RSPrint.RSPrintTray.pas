unit RSPrint.RSPrintTray;

interface

uses
  Classes, Menus, Graphics, Windows, RSPrint.CommonTypes, Messages;

type
  TRSPrintTray = class(TComponent)
  private const
    WM_TASKICON = WM_USER + 10;

  private
    FPopupMenuL: TPopupMenu;
    FIcon: TIcon;
    FTip: string;
    FWnd: HWnd;
    FPrinterStatus: TPrinterStatus;

    procedure DoPopup(i: integer);
    procedure TBChange;
    procedure SetTip(s: string);
    procedure PauseClick(Sender: TObject);
    procedure CancelClick(Sender: TObject);
    procedure CancelAllClick(Sender: TObject);

  protected
    procedure WndProc(var Msg: TMessage);
    procedure DoOnLeftClick; virtual;

  public
    constructor Create(AOwner: TComponent; printerStatus: TPrinterStatus); reintroduce;
    destructor Destroy; override;

  published
    property Tip: string read FTip write SetTip;
  end;

implementation

uses
  ShellAPI, SysUtils;

procedure TRSPrintTray.CancelClick(Sender: TObject);
begin
  FPrinterStatus.CancelPrinting;
end;

procedure TRSPrintTray.CancelAllClick(Sender: TObject);
begin
  FPrinterStatus.CancelAllPrinting;
end;

constructor TRSPrintTray.Create(AOwner: TComponent; printerStatus: TPrinterStatus);
var
  TNIData: TNOTIFYICONDATA;
  PauseItem : TMenuItem;
  CancelItem : TMenuItem;
  CancelAllItem : TMenuItem;
  DivItem : TMenuItem;
begin
  inherited Create(AOwner);

  FPrinterStatus := printerStatus;
  FWnd := AllocateHWnd(WndProc);
  FIcon := TIcon.Create;
  FIcon.Handle := LoadIcon(hinstance,'PRINTER');
  FPopUpMenuL := TPopupMenu.Create(self);
  PauseItem := TMenuItem.Create(FPopupMenuL);
  PauseItem.OnClick := PauseClick;
  PauseItem.Caption := 'Em pausa ';
  DivItem := TMenuItem.Create(FPopupMenuL);
  DivItem.Caption := '-';
  CancelItem := TMenuITem.Create(FPopupMenuL);
  CancelItem.OnClick := CancelClick;
  CancelItem.Caption := 'Cancelar a impressão';
  CancelAllItem := TMenuItem.Create(FPopupMenuL);
  CancelAllItem.OnClick := CancelAllClick;
  CancelAllItem.Caption := 'Cancelar toda a impressão';
  FPopUpMenuL.Items.Add(CancelItem);
  FPopUpMenuL.Items.Add(CancelAllItem);
  FPopUpMenuL.Items.Add(DivItem);
  FPopUpMenuL.Items.Add(PauseItem);
  with TNIData do
  begin
    cbSize := TNOTIFYICONDATA.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_ADD, @TNIData);
  end;
end;

destructor TRSPrintTray.Destroy;
var
  TNIData: TNOTIFYICONDATA;
begin
  FPopUpMenuL.Free;
  with TNIData do
  begin
    cbSize := TNOTIFYICONDATA.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_DELETE, @TNIData);
  end;
  FIcon.Free;
  DeallocateHWnd(FWnd);

  inherited Destroy;
end;

procedure TRSPrintTray.DoOnLeftClick;
begin
  if Assigned(FPopupMenuL) then
    DoPopup(1);
end;

procedure TRSPrintTray.DoPopup(i: integer);
var
  Point : TPoint;
begin
  GetCursorPos(Point);
  SetForeGroundWindow(FWnd);
  case i of
    1: FPopupmenuL.Popup(Point.X, Point.Y);
  end;
  PostMessage(0, 0, 0, 0);
end;

procedure TRSPrintTray.PauseClick(Sender: TObject);
begin
  if not TMenuItem(Sender).Checked then
  begin
    TMenuItem(Sender).Checked := True;
    FPrinterStatus.PausePrinting;
    Tip := 'Em pausa ' + FPrinterStatus.PrintingJobName;
  end
  else
  begin
    TMenuItem(Sender).Checked := False;
    FPrinterStatus.RestorePrinting;
    Tip := 'Imprimindo ' + FPrinterStatus.PrintingJobName;
  end;
end;

procedure TRSPrintTray.SetTip(s: string);
begin
  if FTip <> s then
  begin
    FTip := s;
    if not (csDesigning in ComponentState) then
      TBChange;
  end;
end;

procedure TRSPrintTray.TBChange;
var
  TNIData: TNOTIFYICONDATA;
begin
  with TNIData do
  begin
   cbSize := TNOTIFYICONDATAW.SizeOf;
    Wnd := FWnd;
    uID := 1;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    uCallbackMessage := WM_TASKICON;
    hIcon := FIcon.Handle;
    StrCopy(szTip, PChar(FTip));
    Shell_NotifyIcon(NIM_MODIFY, @TNIData);
  end;
end;

procedure TRSPrintTray.WndProc(var Msg: TMessage);
begin
  with Msg do
  begin
    if Msg = WM_TASKICON then
      case LParamLo of
        WM_LBUTTONDOWN: DoOnLeftClick;
      end
    else
      Result := DefWindowProc(FWnd, Msg, wParam, lParam);
  end;
end;

end.

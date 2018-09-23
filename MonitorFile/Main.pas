unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Tabs, XLSGrid5, Grids, ExtCtrls, XLSUtils5, XLSReadWriteII5;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    XLSGrid: TXLSGrid;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Label1: TLabel;
    edFilename: TEdit;
    Button4: TButton;
    dlgOpen: TOpenDialog;
    Label2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    procedure FileChanged(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  XLSGrid.XLS.Filename := edFilename.Text;
  XLSGrid.XLS.Read;

  XLSGrid.XLS.MonitorFile := True;
end;

procedure TfrmMain.Button3Click(Sender: TObject);
begin
  XLSGrid.XLS.MonitorFile := False;
end;

procedure TfrmMain.Button4Click(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edFilename.Text;
  if dlgOpen.Execute then
    edFilename.Text := dlgOpen.FileName;
end;

procedure TfrmMain.FileChanged(Sender: TObject);
begin
  XLSGrid.XLS.Read;

  Application.ProcessMessages;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  dlgOpen.Filter := XLSExcelFilesFilter;

  XLSGrid.XLS.OnMonitorFile := FileChanged;
end;

end.

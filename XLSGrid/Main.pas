unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Tabs, Grids, XLSGrid5, StdCtrls, ExtCtrls;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    XLSGrid: TXLSGrid;
    Button2: TButton;
    Label1: TLabel;
    edFilename: TEdit;
    Button3: TButton;
    dlgOpen: TOpenDialog;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
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

procedure TfrmMain.Button3Click(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edFilename.Text;
  if dlgOpen.Execute then
    edFilename.Text := dlgOpen.FileName;
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  XLSGrid.XLS.Filename := edFileName.Text;
  XLSGrid.XLS.Read;
end;

end.

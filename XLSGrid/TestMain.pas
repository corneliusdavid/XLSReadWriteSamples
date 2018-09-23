unit TestMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Tabs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Grids, XLSGrid5, XLSReadWriteII5;

type
  TForm5 = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    edFilename: TEdit;
    Label1: TLabel;
    Button3: TButton;
    dlgOpen: TOpenDialog;
    FXLSGrid: TXLSGrid;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
//    FXLSGrid: TXLSGrid;
  public
    { Public declarations }
  end;

var
  Form5: TForm5;

implementation

{$R *.dfm}

procedure TForm5.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TForm5.Button2Click(Sender: TObject);
begin
  FXLSGrid.XLS.Filename := edFilename.Text;
  FXLSGrid.XLS.Read;
end;

procedure TForm5.Button3Click(Sender: TObject);
begin
  dlgOpen.FileName := edFilename.Text;
  if dlgOpen.Execute then
    edFilename.Text := dlgOpen.FileName;
end;

procedure TForm5.FormCreate(Sender: TObject);
begin
//  FXLSGrid := TXLSGrid.Create(Owner);
//  FXLSGrid.Parent := Self;
//  FXLSGrid.Align := alClient;
//
//  FXLSGrid.Top := 200;
//  FXLSGrid.Height := 200;
//  FXLSGrid.Width := 500;
end;

end.

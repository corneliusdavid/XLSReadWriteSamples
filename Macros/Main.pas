unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, XLSSheetData5, XLSReadWriteII5, BIFF_VBA5,
  Xc12Utils5;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    btnClose: TButton;
    btnWrite: TButton;
    btnRead: TButton;
    edFilename: TEdit;
    btnFilename: TButton;
    Memo1: TMemo;
    Panel3: TPanel;
    memModule: TMemo;
    Button1: TButton;
    Button2: TButton;
    Panel2: TPanel;
    Label2: TLabel;
    lbModules: TListBox;
    btnOpenMod: TButton;
    btnNewMod: TButton;
    btnDeleteMod: TButton;
    Splitter1: TSplitter;
    dlgOpen: TOpenDialog;
    dlgOpenTxt: TOpenDialog;
    XLS: TXLSReadWriteII5;
    procedure btnWriteClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnOpenModClick(Sender: TObject);
    procedure btnNewModClick(Sender: TObject);
    procedure btnDeleteModClick(Sender: TObject);
    procedure lbModulesDblClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnFilenameClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := 'd:\xtemp\wtest.xls'; //edFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnReadClick(Sender: TObject);
var      
  i: integer;
begin
  XLS.Filename := edFilename.Text;
  XLS.BIFF.ReadMacros := True;
  XLS.Read;
  lbModules.Clear;
  for i := 0 to XLS.VBA.Count - 1 do
    lbModules.Items.Add(XLS.VBA[i].Name);
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  if lbModules.ItemIndex >= 0 then
    XLS.VBA[lbModules.ItemIndex].Source.Assign(memModule.Lines);
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  if dlgOpenTxt.Execute then
    memModule.Lines.LoadFromFile(dlgOpenTxt.FileName);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  XLS.Version := xvExcel97;
  // If not EditVBA is set to True, no macros can be changed.
  XLS.BIFF.ReadMacros := True;
  XLS.VBA.EditVBA := True;
end;

procedure TfrmMain.btnOpenModClick(Sender: TObject);
begin
  if lbModules.ItemIndex >= 0 then
    memModule.Lines.Assign(XLS.VBA[lbModules.ItemIndex].Source);
end;

procedure TfrmMain.btnNewModClick(Sender: TObject);
var
  N: string;
begin
  if InputQuery('New module','Name',N) then begin
    XLS.VBA.AddModule(N,vmtMacro);
    lbModules.Items.Add(N);
  end;
end;

procedure TfrmMain.btnDeleteModClick(Sender: TObject);
begin
  if (lbModules.ItemIndex >= 0) and (MessageDlg('Delete module ' + lbModules.Items[lbModules.ItemIndex] + '?',mtConfirmation,[mbYes,mbNo],0) = mrYes) then begin
    XLS.VBA.DeleteModule(lbModules.Items[lbModules.ItemIndex]);
    lbModules.Items.Delete(lbModules.ItemIndex);
  end;
end;

procedure TfrmMain.lbModulesDblClick(Sender: TObject);
begin
  btnOpenMod.Click;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnFilenameClick(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edFilename.Text;
  if dlgOpen.Execute then
    edFilename.Text := dlgOpen.FileName;
end;

end.

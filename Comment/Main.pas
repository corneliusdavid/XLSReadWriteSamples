unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XLSReadWriteII5, StdCtrls, ExtCtrls, IniFiles, XLSComment5, Xc12Utils5,
  XLSSheetData5, XPMan, XLSCmdFormat5;

type
  TfrmMain = class(TForm)
    Label2: TLabel;
    Label3: TLabel;
    Panel1: TPanel;
    Label1: TLabel;
    btnRead: TButton;
    btnWrite: TButton;
    edReadFilename: TEdit;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    btnDlgSave: TButton;
    cbComments: TComboBox;
    btnAdd1: TButton;
    btnAdd2: TButton;
    btnClose: TButton;
    edAuthor: TEdit;
    memText: TMemo;
    dlgSave: TSaveDialog;
    dlgOpen: TOpenDialog;
    XLS: TXLSReadWriteII5;
    XPManifest1: TXPManifest;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure cbCommentsCloseUp(Sender: TObject);
    procedure btnAdd1Click(Sender: TObject);
    procedure btnAdd2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    procedure ReadComments;
    procedure ShowComment;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnReadClick(Sender: TObject);
begin
  XLS.Filename := edReadFilename.Text;
  XLS.Read;

  ReadComments;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edWriteFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edReadFilename.Text;
  if dlgOpen.Execute then
    edReadFilename.Text := dlgOpen.FileName;
end;

procedure TfrmMain.btnDlgSaveClick(Sender: TObject);
begin
  dlgSave.FileName := edWriteFilename.Text;
  if dlgSave.Execute then
    edWriteFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.cbCommentsCloseUp(Sender: TObject);
begin
  ShowComment;
end;

procedure TfrmMain.btnAdd1Click(Sender: TObject);
begin
//Add a comment with author.
  XLS[0].Comments.Add(1,4,'Frank Borland','Hello,world');

  ReadComments;
end;

procedure TfrmMain.btnAdd2Click(Sender: TObject);
begin
//Add a comment without author.
  XLS[0].Comments.Add(1,5,'','Hello,world');

  ReadComments;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    edReadFilename.Text := Ini.ReadString('Files','Read','');
    edWriteFilename.Text := Ini.ReadString('Files','Write','');
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    Ini.WriteString('Files','Read',edReadFilename.Text);
    Ini.WriteString('Files','Write',edWriteFilename.Text);
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.ReadComments;
var
  i: integer;
  Comment: TXLSComment;
begin
  cbComments.Clear;

  for i := 0 to XLS[0].Comments.Count - 1 do begin
    Comment := XLS[0].Comments[i];
    cbComments.Items.Add(ColRowToRefStr(Comment.Col,Comment.Row));
  end;
  if cbComments.Items.Count > 0 then begin
    cbComments.ItemIndex := 0;
    ShowComment;
  end;
end;

procedure TfrmMain.ShowComment;
var
  Comment: TXLSComment;
begin
  if cbComments.ItemIndex >= 0 then begin
    Comment := XLS[0].Comments[cbComments.ItemIndex];
    edAuthor.Text := Comment.Author;
    memText.Lines.Text := Comment.PlainText;
  end;
end;

end.

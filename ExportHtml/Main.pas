unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, ExtCtrls, IniFiles, XLSUtils5, Xc12Utils5,
  XLSSheetData5, XLSReadWriteII5, XLSExport5, XLSExportHTML5, ComCtrls,
  XPMan;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnRead: TButton;
    btnWrite: TButton;
    edReadFilename: TEdit;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    btnDlgSave: TButton;
    btnClose: TButton;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    clbSheets: TCheckListBox;
    Label1: TLabel;
    XLS: TXLSReadWriteII5;
    cbSimpleExport: TCheckBox;
    Label2: TLabel;
    cbComments: TCheckBox;
    cbImages: TCheckBox;
    cbHyperlinks: TCheckBox;
    Label3: TLabel;
    ExportHTML: TXLSExportHTML5;
    cbSeparateFiles: TCheckBox;
    Progress: TProgressBar;
    cbWriteImages: TCheckBox;
    cbExportAll: TCheckBox;
    Label4: TLabel;
    edArea: TEdit;
    Label5: TLabel;
    XPManifest1: TXPManifest;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure XLSProgress(AProgressType: TXLSProgressType;
      AProgressState: TXLSProgressState; AValue: Double);
    procedure btnWriteClick(Sender: TObject);
    procedure cbSimpleExportClick(Sender: TObject);
    procedure cbSeparateFilesClick(Sender: TObject);
    procedure cbImagesClick(Sender: TObject);
    procedure cbCommentsClick(Sender: TObject);
    procedure cbWriteImagesClick(Sender: TObject);
    procedure cbHyperlinksClick(Sender: TObject);
    procedure cbExportAllClick(Sender: TObject);
  private
    procedure FillSheets;
    procedure SetOptions;
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

  FillSheets;
end;

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edReadFilename.Text;
  if dlgOpen.Execute then
    edReadFilename.Text := dlgOpen.FileName;
end;

procedure TfrmMain.FillSheets;
var
  i: integer;
begin
  clbSheets.Clear;

  for i := 0 to XLS.Count - 1 do begin
    clbSheets.Items.Add(XLS[i].Name);
    clbSheets.Checked[i] := True;
  end;
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

procedure TfrmMain.btnDlgSaveClick(Sender: TObject);
begin
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave.FileName := edWriteFilename.Text;
  if dlgSave.Execute then
    edWriteFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.XLSProgress(AProgressType: TXLSProgressType; AProgressState: TXLSProgressState; AValue: Double);
begin
  case AProgressState of
    xpsBegin : Progress.Visible := True;
    xpsWork  : Progress.Position := Round(AValue * Progress.Max);
    xpsEnd   : Progress.Visible := False;
  end;
  Progress.Update;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
var
  i: integer;
  SelectedSheets: array of integer;
  C1,R1,C2,R2: integer;
begin
  for i := 0 to clbSheets.Items.Count - 1 do begin
    if clbSheets.Checked[i] then begin
      SetLength(SelectedSheets,Length(SelectedSheets) + 1);
      SelectedSheets[High(SelectedSheets)] := i;
    end;
  end;

  // If not all sheets are selected. By default all sheets are exported.
  if Length(SelectedSheets) <  clbSheets.Items.Count then
    ExportHTML.Sheets(SelectedSheets);

  if not cbExportAll.Checked and AreaStrToColRow(edArea.Text,C1,R1,C2,R2) then begin
    ExportHTML.Col1 := C1;
    ExportHTML.Row1 := R1;
    ExportHTML.Col2 := C2;
    ExportHTML.Row2 := R2;
  end;

  ExportHTML.Filename := edWriteFilename.Text;
  ExportHTML.Write;
end;

procedure TfrmMain.cbSimpleExportClick(Sender: TObject);
begin
  ExportHTML.SimpleExport := cbSimpleExport.Checked;

  cbComments.Enabled := not cbSimpleExport.Checked;
  cbImages.Enabled := not cbSimpleExport.Checked;
  cbWriteImages.Enabled := not cbSimpleExport.Checked and cbImages.Checked;
  cbHyperlinks.Enabled := not cbSimpleExport.Checked;
end;

procedure TfrmMain.cbSeparateFilesClick(Sender: TObject);
begin
  SetOptions;
end;

procedure TfrmMain.cbImagesClick(Sender: TObject);
begin
  cbWriteImages.Enabled := cbImages.Checked;
end;

procedure TfrmMain.SetOptions;
begin
  ExportHTML.HTMLOPtions := [];

  if cbComments.Checked then
    ExportHTML.HTMLOPtions := ExportHTML.HTMLOPtions + [xeohComments];
  if cbImages.Checked then
    ExportHTML.HTMLOPtions := ExportHTML.HTMLOPtions + [xeohImages];
  if cbWriteImages.Checked then
    ExportHTML.HTMLOPtions := ExportHTML.HTMLOPtions + [xeohWriteImages];
  if cbHyperlinks.Checked then
    ExportHTML.HTMLOPtions := ExportHTML.HTMLOPtions + [xeohHyperlinks];
end;

procedure TfrmMain.cbCommentsClick(Sender: TObject);
begin
  SetOptions;
end;

procedure TfrmMain.cbWriteImagesClick(Sender: TObject);
begin
  SetOptions;
end;

procedure TfrmMain.cbHyperlinksClick(Sender: TObject);
begin
  SetOptions;
end;

procedure TfrmMain.cbExportAllClick(Sender: TObject);
begin
  edArea.Enabled := not cbExportAll.Checked;
end;

end.

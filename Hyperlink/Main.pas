unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, IniFiles, ShellApi,

  // XLSReadWriteII units.
  XLSReadWriteII5, XLSHyperlinks5, Xc12Utils5, Xc12DataStylesheet5, XLSSheetData5,
  XPMan;

type TDoubleArray = array of double;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnRead: TButton;
    btnWrite: TButton;
    edReadFilename: TEdit;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    btnDlgSave: TButton;
    dlgSave: TSaveDialog;
    dlgOpen: TOpenDialog;
    Button1: TButton;
    Grid: TStringGrid;
    btnAddCells: TButton;
    XLS: TXLSReadWriteII5;
    Button2: TButton;
    XPManifest1: TXPManifest;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnAddCellsClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    procedure AddHyperlinks;

    procedure ReadHyperlinks;
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

  ReadHyperlinks;
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
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave.FileName := edWriteFilename.Text;
  if dlgSave.Execute then
    edWriteFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  Close;
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

  Grid.Cells[0,0] := 'Cell';
  Grid.Cells[1,0] := 'Address';
  Grid.Cells[2,0] := 'Type';
  Grid.Cells[3,0] := 'Description';
  Grid.Cells[4,0] := 'Tooltip';
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

procedure TfrmMain.btnAddCellsClick(Sender: TObject);
begin
  AddHyperlinks;

  ReadHyperlinks;
end;

procedure TfrmMain.AddHyperlinks;
var
  HLink: TXLSHyperlink;
begin
// There can be four types of hyperlinks:
// 1. URL, an internet address (http://, ftp:// , mailto:).
// 2. UNC, a network file address (\\MyServer\ ...).
// 3. The current workbook (Sheet1!B2).
// 4. A local file.
// The hyperlink type is detected automatically. If it don't matches any of the
// first three, it is assumed to be a local file. This can be changed with the
// HyperlinkType property.

// Hyperlinks are in the Hyperlink property of the worksheet. This property
// contains the address settings for the hyperlink, but not the cell text.
// That is stored as a normal cell value.

// There are several options for creating hyperlinks:

// *****************************************************************************
// 1. Add a Hyperlink object.

  HLink := XLS[0].Hyperlinks.Add;

  // The cell of the hyperlink.
  HLink.Col1 := 1;
  HLink.Row1 := 4;
  HLink.Col2 := 1;
  HLink.Row2 := 4;
  // The URL must start with http:// otherwhise will Excel think the address is a file.
  HLink.Address := 'http://www.embarcadero.com';
  // Optional tooltip.
  HLink.Tooltip := 'Embarcadero sells Delphi and burning monkeys';
  // Description is not used by Excel.
  HLink.Description := '???';

  // Set cell value and standard hyperlink formatting (blue text and underline).
  XLS[0].AsString[1,4] := 'Embarcadero';
  XLS[0].Cell[1,4].FontColor := $000000FF;
  XLS[0].Cell[1,4].FontUnderline := xulSingle;

// *****************************************************************************
// 2. Use the AsHyperlink property. This will set the address options.
// AsHyperlink will add http:// to the address if it starts with www
  XLS[0].AsHyperlink[1,5] := 'http://www.embarcadero.com';

  // Set cell value
  XLS[0].AsString[1,5] := 'Embarcadero';
  XLS[0].Cell[1,5].FontColor := $000000FF;
  XLS[0].Cell[1,5].FontUnderline := xulSingle;

// *****************************************************************************
// 3. Use the MakeHyperlink method. Parameters are: Col,Row,Address,[Tooltip]
// MakeHyperlink will add http:// to the address if it starts with www
  XLS[0].MakeHyperlink(1,6,'www.borland.com','Borland');
  XLS[0].MakeHyperlink(1,7,'www.borland.com','Borland','Borland do not sell Delphi anymore.');

  // Add a mail address
  XLS[0].MakeHyperlink(1,8,'mailto:monkeys@embarcadero.com','Embarcadero');

  // Add a cell reference in the workbook.
  XLS[0].MakeHyperlink(1,8,'Sheet1!E10','Cell E10');

  // Add a disk file.
  XLS[0].MakeHyperlink(1,9,'c:\temp\text.txt','Text file');
end;

procedure TfrmMain.ReadHyperlinks;
var
  i: integer;
  HLink: TXLSHyperlink;
begin
  for i := 0 to XLS[0].Hyperlinks.Count - 1 do begin

    if i >= Grid.RowCount then
      Exit;

    HLink := XLS[0].Hyperlinks[i];

    Grid.Cells[0,i + 1] := ColRowToRefStr(HLink.Col1,HLink.Row1);
    Grid.Cells[1,i + 1] := HLink.Address;
    case HLink.HyperlinkType of
      xhltUnknown : Grid.Cells[2,i + 1] := '[Unknown]';
      xhltURL     : Grid.Cells[2,i + 1] := 'URL';
      xhltFile    : Grid.Cells[2,i + 1] := 'File';
      xhltUNC     : Grid.Cells[2,i + 1] := 'UNC';
      xhltWorkbook: Grid.Cells[2,i + 1] := 'Workbook';
    end;
    Grid.Cells[3,i + 1] := HLink.Description;
    Grid.Cells[4,i + 1] := HLink.Tooltip;
  end;
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PWideChar(edWriteFilename.Text), nil, SW_SHOWNORMAL);
end;

end.

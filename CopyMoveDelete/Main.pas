unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, XLSReadWriteII5, XLSHyperlinks5, Grids, IniFiles,
  Xc12Utils5, Xc12DataStylesheet5, XLSSheetData5, XPMan;

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
    XLS: TXLSReadWriteII5;
    Button1: TButton;
    btnCopy: TButton;
    XPManifest1: TXPManifest;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCopyClick(Sender: TObject);
  private
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

procedure TfrmMain.btnCopyClick(Sender: TObject);
var
  Col1,Col2,Row1,Row2: integer;
  DestCol,DestRow: integer;
  SrcSheet,DestSheet: integer;
begin
  // Copy cell A1 to cell G1.
  Col1 := 0;
  Row1 := 0;
  DestCol := 6;
  DestRow := 0;
  XLS[0].CopyCell(Col1,Row1,DestCol,DestRow, False);

  // Move cell A2 to cell G2.
  Col1 := 0;
  Row1 := 1;
  DestCol := 6;
  DestRow := 1;
  XLS[0].MoveCell(Col1,Row1,DestCol,DestRow);

  // Delete cell A3.
  Col1 := 0;
  Row1 := 2;
  XLS[0].DeleteCell(Col1,Row1);

  // Copy cells from B1:B25 to H1:H25.
  Col1 := 1;
  Row1 := 0;
  Col2 := 1;
  Row2 := 24;
  DestCol := 7;
  DestRow := 0;
  XLS[0].CopyCells(Col1,Row1,Col2,Row2,DestCol,DestRow);

  // Move cells from C1:C13 to I1:I13.
  Col1 := 2;
  Row1 := 0;
  Col2 := 2;
  Row2 := 12;
  DestCol := 8;
  DestRow := 0;
  XLS[0].MoveCells(Col1,Row1,Col2,Row2,DestCol,DestRow);

  // Delete cells at C14:C25 to I1:I13.
  Col1 := 2;
  Row1 := 13;
  Col2 := 2;
  Row2 := 24;
  XLS[0].DeleteCells(Col1,Row1,Col2,Row2);

(*
  // Copy cells from D1:D25 on sheet 1 to to A1:A25 on sheet 2.
  SrcSheet := 0;
  Col1 := 3;
  Row1 := 0;
  Col2 := 3;
  Row2 := 24;
  DestSheet := 1;
  DestCol := 0;
  DestRow := 0;
  XLS.CopyCells(SrcSheet,Col1,Row1,Col2,Row2,DestSheet,DestCol,DestRow);

  // Copy cells from E1:E13 on sheet 1 to to B1:B13 on sheet 2.
  SrcSheet := 0;
  Col1 := 4;
  Row1 := 0;
  Col2 := 4;
  Row2 := 12;
  DestSheet := 1;
  DestCol := 1;
  DestRow := 0;
  XLS.MoveCells(SrcSheet,Col1,Row1,Col2,Row2,DestSheet,DestCol,DestRow);

  // Delete cells at E14:E25 on sheet 1.
  SrcSheet := 0;
  Col1 := 4;
  Row1 := 13;
  Col2 := 4;
  Row2 := 24;
  XLS.DeleteCells(SrcSheet,Col1,Row1,Col2,Row2);

  // Delete rows 5 and 6. Rows below the deleted rows are shifted up.
  Row1 := 5;
  Row2 := 6;
  XLS[0].DeleteRows(Row1,Row2);

  // Clear rows 10 and 11. Rows below the deleted rows are unchanged.
  Row1 := 10;
  Row2 := 11;
  XLS[0].ClearRows(10,11);

  // Insert four rows at row 14. Rows below row 14 are shifted down.
  Row1 := 14;
  XLS[0].InsertRows(Row1,4);
*)
  // Row operations are available for columns as well.
end;

end.

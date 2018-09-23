unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, IniFiles, ShellApi, XPMan,
  // XLSReadWriteII5 units
  XLSSheetData5, XLSReadWriteII5, XLSCmdFormat5, Xc12Utils5, Xc12Manager5;

// ************************************************************************
// ******* DIRECT WRITE SAMPLE                                      *******
// ******* The advantage of using direct write instead of first     *******
// ******* createing the file in memory is that direct write uses   *******
// ******* much less memory. Creating a large file with several     *******
// ******* millions cells can easily use several 100 Mb of memory.  *******
// ******* It's also faster to create a file with direct write      *******
// ************************************************************************

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnClose: TButton;
    btnWrite: TButton;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    dlgSave: TSaveDialog;
    XLS: TXLSReadWriteII5;
    btnExcel: TButton;
    XPManifest1: TXPManifest;
    procedure btnCloseClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure XLSWriteCell(ACell: TXLSEventCell);
    procedure btnWriteClick(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure edWriteFilenameChange(Sender: TObject);
  private
    FFmtDefString: TXLSDefaultFormat;
    FFmtDefNumber: TXLSDefaultFormat;

    procedure CreateFormats;
  public
    procedure DoDirectWrite;
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgSave.FileName := edWriteFilename.Text;
  if dlgSave.Execute then
    edWriteFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
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
    Ini.WriteString('Files','Write',edWriteFilename.Text);
  finally
    Ini.Free;
  end;
end;

// OnWriteCell event, fired for every cell in the target area.
// The TXLSEventCell has the following usefull methods and properties:
//
// type TXLSEventCell = class(TObject)
// public
//      Aborts writing.
//      procedure Abort;
//
//      The target sheet.
//      property SheetIndex: integer;
//      The column that is to be written in the OnWriteCell event.
//      property Col       : integer;
//      The row that is to be written in the OnWriteCell event.
//      property Row       : integer;
//      The target area where the cells are written. The OnWriteCell event is
//      fired for every cell in the target area.
//      property TargetArea: TXLSCellArea;
//
//      Write a boolean cell value.
//      property AsBoolean : boolean read GetBoolean write SetBoolean;
//      Write an error cell value.
//      property AsError   : TXc12CellError read GetError write SetError;
//      Write a floatcell value.
//      property AsFloat   : double read GetFloat write SetFloat;
//      Write a string cell value.
//      property AsString  : AxUCString read GetString write SetString;
//      Write a formula. Please note that formulas are not verified, so
//      you can write any string valaues. It's of course not a good idea
//      to write invalid formulas, so check formulas caerfully first.
//      property AsFormula : AxUCString read GetFormula write SetFormula;
//      end;

procedure TfrmMain.XLSWriteCell(ACell: TXLSEventCell);
begin
  // Writing can be stopped by calling the Abort method of the TXLSEventCell object.
  if ACell.Row > 19 then begin
    ACell.Abort;
    Exit;
  end;

  case (ACell.Col - ACell.TargetArea.Col1) mod 5 of
    0: begin
      // Use the previous created format.
      FFmtDefString.UseByDirectWrite(ACell);
      ACell.AsString := 'Num_' + IntToStr(Random(100000));
    end;
    1: ACell.AsFloat := Random(1000000);
    2: begin
      // Use the previous created format.
      FFmtDefNumber.UseByDirectWrite(ACell);
      ACell.AsFloat := Random(100000000) / 10000;
    end;
    3: ACell.AsFloat := Random(100000);
    4: begin
      // Write a formula.
      ACell.AsFormula := 'SUM(' + AreaToRefStr(4,ACell.Row,6,ACell.Row) + ')';
      // Set the result and type of the formula. This must be done after the
      // formula is assigned.
      // It's of cource not possible to calculate any formulas, as the file is
      // on a disk.
      ACell.AsFloat := 0;
    end;
  end;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  CreateFormats;

  DoDirectWrite;
end;

procedure TfrmMain.DoDirectWrite;
begin
  // Excel file to create.
  XLS.Filename := edWriteFilename.Text;
  // SetDirectWriteArea(const ASheetIndex, ACol1, ARow1, ACol2, ARow2: integer);
  // Set the target area to write, Column1 = 3, Column2 = 7, Row1 = 4, Row2 = 499
  // If not assigned, the default is Column1 = 0, Column2 = 255, Row1 = 0, Row2 = 65536
  // The cells are written to sheet #0. If you want to write to any other sheet,
  // you must first add the sheet(s), as there is only one sheet by default.
  XLS.SetDirectWriteArea(0,3,4,7,499);
  // Set Direct Write mode.
  XLS.DirectWrite := True;
  // Write the file. For every cell in the file, the OnWriteCell event will be fired.
  XLS.Write;
end;

procedure TfrmMain.btnExcelClick(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PAnsiChar(edWriteFilename.Text), nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.edWriteFilenameChange(Sender: TObject);
begin
  btnExcel.Enabled := edWriteFilename.Text <> '';
end;

procedure TfrmMain.CreateFormats;
begin
  // Create some formats. For details on creating formats, see FormatCells sample.

  XLS.CmdFormat.BeginEdit(Nil);
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $7F7F10;
  XLS.CmdFormat.Font.Name := 'Courier new';
  // Save the default format. It is also possible to access default formats by
  // their name, but that is not recomended when writing huge ammounts of cells
  // as that is much slower.
  FFmtDefString := XLS.CmdFormat.AddAsDefault('F1');

  XLS.CmdFormat.BeginEdit(Nil);
  XLS.CmdFormat.Number.Decimals := 2;
  XLS.CmdFormat.Number.Thousands := True;
  // See above.
  FFmtDefNumber := XLS.CmdFormat.AddAsDefault('F2');
end;

end.

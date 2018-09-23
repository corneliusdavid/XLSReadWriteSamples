unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XLSReadWriteII5, StdCtrls, ExtCtrls, IniFiles, ShellApi,

  // XLSReadWriteII5 units
  XLSComment5, XLSDrawing5, Xc12Utils5, Xc12DataStyleSheet5, XLSSheetData5,
  XLSCmdFormat5, XPMan;

// *****************************************************************************
// ***** FORMATTING CELLS                                                  *****
// *****                                                                   *****
// ***** Cell formats in Excel files are not stored in the cells           *****
// ***** themself. Instead is a global list with formats common to         *****
// ***** many cells used. The cells stores an index into this list.        *****
// ***** The max size of the format list is 65536 items. This means        *****
// ***** that the formats list has to be maintained, and searched for      *****
// ***** existing matching formats when a cell is formatted.               *****
// *****                                                                   *****
// ***** There are 3 differnet ways of formatting cells.                   *****
// *****                                                                   *****
// ***** 1. Using the Cell object of the worksheet, XLS[0].Cell[Col,Row]   *****
// *****    There must be a cell(value) before the Cell can be uses, so    *****
// *****    at least add a blank cell. The Cell object is most usefull     *****
// *****    for reading the formatting of a single cell. Creating formats  *****
// *****    with the Cell object is extremly slow and reallly only usefull *****
// *****    for just a few cells.                                          *****
// *****                                                                   *****
// ***** 2. Using the CmdFormat object of the workbook, XLS.CmdFormat.     *****
// *****    The CmdFormat works like the Format Cells dialog in Excel with *****
// *****    child objects for Number, Alignment, Font, Border, Fill and    *****
// *****    Protection properties. While Excel always merges formats when  *****
// *****    a cell is formatted (that is, if you put a border around a     *****
// *****    cell and the cell allready has a backgroud color, that color   *****
// *****    is preserved), XLSReadWriteII also have the option to replace  *****
// *****    any existing format. The advatage of this is that the format   *****
// *****    list don't have to searched and hence it's much faster. About  *****
// *****    ten times faster. If you format an empty worksheet this is     *****
// *****    always the best option.                                        *****
// *****                                                                   *****
// ***** 4. Creating a default format. A default format, if assigned, is   *****
// *****    used when new cells are added. You can have several default    *****
// *****    formats and swithch between them. This is the fastest method.  *****
// *****                                                                   *****
// *****************************************************************************

type
  TfrmMain = class(TForm)
    dlgSave: TSaveDialog;
    XLS: TXLSReadWriteII5;
    btnWrite: TButton;
    edWriteFilename: TEdit;
    btnDlgSave: TButton;
    btnSingleCell: TButton;
    btnAreas: TButton;
    btnClose: TButton;
    btnDefaultFormat: TButton;
    btnExcel: TButton;
    btnProtect: TButton;
    btnColumns: TButton;
    btnRows: TButton;
    Label1: TLabel;
    edPassword: TEdit;
    XPManifest1: TXPManifest;
    Button1: TButton;
    procedure btnCloseClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSingleCellClick(Sender: TObject);
    procedure btnDefaultFormatClick(Sender: TObject);
    procedure btnAreasClick(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure edWriteFilenameChange(Sender: TObject);
    procedure btnProtectClick(Sender: TObject);
    procedure btnColumnsClick(Sender: TObject);
    procedure btnRowsClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnSingleCellClick(Sender: TObject);
begin
  // Formats a single cell. Use only this method for very few cells as it is
  // extremly slow. The cell properties can however be usefull for reading
  // formatting properties of cells.

  XLS[0].AsFloat[2,3] := 125.5;

  XLS[0].Cell[2,3].FillPatternForeColorRGB := $AFAFAF;

  XLS[0].Cell[2,3].NumberFormat := '0.000';

  XLS[0].Cell[2,3].BorderLeftStyle := cbsThick;
  XLS[0].Cell[2,3].BorderLeftColor := xcRed;

  XLS[0].Cell[2,3].BorderRightStyle := cbsThick;
  XLS[0].Cell[2,3].BorderRightColor := xcRed;

  XLS[0].Cell[2,3].BorderTopStyle := cbsThick;
  XLS[0].Cell[2,3].BorderTopColor := xcRed;

  XLS[0].Cell[2,3].BorderBottomStyle := cbsThick;
  XLS[0].Cell[2,3].BorderBottomColor := xcRed;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edWriteFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnDlgSaveClick(Sender: TObject);
begin
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
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

procedure TfrmMain.btnDefaultFormatClick(Sender: TObject);
var
  R: integer;
  DefFmt_F3: TXLSDefaultFormat;
begin
  // Create first default format.
  // The sheet parameter of BeginFormat is Nil, as it's not assigned to any
  // particular worksheet.
  XLS.CmdFormat.BeginEdit(Nil);
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $7F7F10;
  XLS.CmdFormat.Font.Name := 'Courier new';
  XLS.CmdFormat.AddAsDefault('F1');

  // Create second default format.
  XLS.CmdFormat.BeginEdit(Nil);
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $FF1010;
  XLS.CmdFormat.Font.Name := 'Ariel Black';
  XLS.CmdFormat.Border.Color.IndexColor := xcBlack;
  XLS.CmdFormat.Border.Style := cbsThick;
  XLS.CmdFormat.Border.Preset(cbspOutline);
  XLS.CmdFormat.AddAsDefault('F2');

  // Create third default format.
  XLS.CmdFormat.BeginEdit(Nil);
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $10F040;
  XLS.CmdFormat.Font.Name := 'Times New Roman';
  XLS.CmdFormat.Font.Style := [xfsBold,xfsItalic];
  DefFmt_F3 := XLS.CmdFormat.AddAsDefault('F3');

  // The default formats are in the XLS.CmdFormat.Defaults property.

  for R := 2 to 100 do begin
    if (R mod 4) = 0 then
      // Assign format by name.
      XLS.DefaultFormat := XLS.CmdFormat.Defaults.Find('F2')
    else if (R mod 7) = 0 then
      // Assign format by variable.
      XLS.DefaultFormat := DefFmt_F3
    else
      // Assign format by index. We know that format 'F1' has index #0 as the
      // list was empty.
      XLS.DefaultFormat := XLS.CmdFormat.Defaults[0];
    XLS[0].AsFloat[2,R] := 525;
  end;

  // Reset the default format.
  XLS.DefaultFormat := Nil;
end;

procedure TfrmMain.btnAreasClick(Sender: TObject);
begin
  // There are two formatting modes, Merge and Replace. Merge will merge the
  // assigned properies with any existing cells. This is how Excel works.
  // Replace will replace any existing formats with the new format.
  // The Replace mode is about 10 times faster than merge, as merging formats
  // means that for every cell the format lists has to be searched for the
  // combination of the new and the old format.
  // Merge is the default mode.
  XLS.CmdFormat.Mode := xcfmMerge;
//  XLS.CmdFormat.Mode := xcfmReplace;

  // Begin formatting on sheet #0.
  XLS.CmdFormat.BeginEdit(XLS[0]);

  // Set background color.
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $7F7F10;

  // Border color and style.
  XLS.CmdFormat.Border.Color.RGB := $7F1010;
  XLS.CmdFormat.Border.Style := cbsThick;
  // Set the border outline of the cells.
  XLS.CmdFormat.Border.Preset(cbspOutline);

  // Border style of lines inside.
  XLS.CmdFormat.Border.Style := cbsThin;
  XLS.CmdFormat.Border.Preset(cbspInside);

  // Apply format.
  XLS.CmdFormat.Apply(2,2,10,30);

  // Formats can also be applied on the current selection
  XLS[0].SelectedAreas[0].Col1 := 12;
  XLS[0].SelectedAreas[0].Row1 := 2;
  XLS[0].SelectedAreas[0].Col2 := 15;
  XLS[0].SelectedAreas[0].Row2 := 30;

  // Without arguments Apply will apply the format on the selected cells.
  XLS.CmdFormat.Apply;
end;

procedure TfrmMain.btnExcelClick(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PwideChar(edWriteFilename.Text), nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.edWriteFilenameChange(Sender: TObject);
begin
  btnExcel.Enabled := edWriteFilename.Text <> '';
end;

procedure TfrmMain.btnProtectClick(Sender: TObject);
begin
  XLS[0].AsString[4,3] := 'Unprotected cells';

  // Protect cells on sheet #0.
  XLS.CmdFormat.BeginEdit(XLS[0]);

  // Create an area of unprotected cells, that is, the cells are not locked.
  XLS.CmdFormat.Fill.BackgroundColor.RGB := $AFAFAF;
  XLS.CmdFormat.Protection.Locked := False;
  XLS.CmdFormat.Apply(4,4,4,10);

  // Set the password of the worksheet.
  XLS[0].Protection.Password := edPassword.Text;
end;

procedure TfrmMain.btnColumnsClick(Sender: TObject);
begin
  XLS.CmdFormat.BeginEdit(XLS[0]);

  XLS.CmdFormat.Border.Color.RGB := $0000FF;
  XLS.CmdFormat.Border.Style := cbsThick;
  XLS.CmdFormat.Border.Side[cbsLeft] := True;
  XLS.CmdFormat.Border.Side[cbsRight] := True;

  XLS.CmdFormat.ApplyCols(1,5);
end;

procedure TfrmMain.btnRowsClick(Sender: TObject);
begin
  XLS.CmdFormat.BeginEdit(XLS[0]);

  XLS.CmdFormat.Border.Color.RGB := $FF0000;
  XLS.CmdFormat.Border.Style := cbsThick;
  XLS.CmdFormat.Border.Side[cbsTop] := True;
  XLS.CmdFormat.Border.Side[cbsBottom] := True;

  XLS.CmdFormat.ApplyRows(1,5);
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  // Shows how to set individual borders for cells

  XLS.CmdFormat.BeginEdit(XLS[0]);

  // Border color and style must be set first.
  XLS.CmdFormat.Border.Color.IndexColor := xcRed;
  XLS.CmdFormat.Border.Style := cbsThick;

  // Apply the color and style on selected sides.
  XLS.CmdFormat.Border.Side[cbsTop] := True;
  XLS.CmdFormat.Border.Side[cbsBottom] := True;

  XLS.CmdFormat.Apply(11,3,18,3);
end;

end.

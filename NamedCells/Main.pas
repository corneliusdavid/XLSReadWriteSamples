unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, IniFiles, ShellApi, ComCtrls, XPMan, EditName,
  // XLSReadWriteII units
  Xc12Utils5, Xc12DataWorkbook5, XLSSheetData5, XLSReadWriteII5, XLSNames5;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnRead: TButton;
    btnWrite: TButton;
    edReadFilename: TEdit;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    btnDlgSave: TButton;
    btnAddName: TButton;
    btnEditName: TButton;
    btnClose: TButton;
    XLS: TXLSReadWriteII5;
    Button1: TButton;
    lvNames: TListView;
    XPManifest: TXPManifest;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    btnNameValues: TButton;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure btnAddNameClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnEditNameClick(Sender: TObject);
    procedure btnNameValuesClick(Sender: TObject);
  private
    procedure ReadNames;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

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
    Ini.WriteString('Files','Read', edReadFilename.Text);
    Ini.WriteString('Files','Write',edWriteFilename.Text);
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PWideChar(edWriteFilename.Text), nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.btnReadClick(Sender: TObject);
begin
  XLS.Filename := edReadFilename.Text;
  XLS.Read;

  ReadNames;
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

procedure TfrmMain.ReadNames;
var
  i: integer;
  N: TXLSName;
  Item: TListItem;
begin
  lvNames.Items.BeginUpdate;

  lvNames.Clear;
  
  try
    for i := 0 to XLS.Names.Count - 1 do begin
      N := XLS.Names.Items[i];

      Item := lvNames.Items.Add;

      case N.BuiltIn of
        bnConsolidateArea: Item.Caption := 'Consolidate_Area';
        bnExtract        : Item.Caption := 'Extract';
        bnDatabase       : Item.Caption := 'Database';
        bnCriteria       : Item.Caption := 'Criteria';
        bnPrintArea      : Item.Caption := 'Print_Area';
        bnPrintTitles    : Item.Caption := 'Print_Titles';
        bnSheetTitle     : Item.Caption := 'Sheet_Title';
        bnFilterDatabase : Item.Caption := 'Filter_Database';
        bnNone           : Item.Caption := N.Name;
      end;

      if N.SimpleName in [xsntRef,xsntArea] then
        Item.SubItems.Add(XLS.NameAsString[N.Name])
      else
        Item.SubItems.Add('');

      Item.SubItems.Add(N.Definition);

      if N.LocalSheetId >= 0 then
        Item.SubItems.Add(XLS[N.LocalSheetId].Name)
      else
        Item.SubItems.Add('Workbook');

      Item.SubItems.Add(N.Comment);
    end;
  finally
    lvNames.Items.EndUpdate;
  end;
end;

procedure TfrmMain.btnAddNameClick(Sender: TObject);
begin
  if lvNames.ItemIndex >= 0 then begin
    if TfrmEditName.Create(Self).Execute(XLS,XLS.Names[lvNames.ItemIndex]) then
      ReadNames;
  end;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edWriteFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnEditNameClick(Sender: TObject);
begin
  if TfrmEditName.Create(Self).Execute(XLS,Nil) then
    ReadNames;
end;

procedure TfrmMain.btnNameValuesClick(Sender: TObject);
var
  S: string;
  Name: TXLSName;
  Row,Col: integer;
  Sheet: integer;
begin
  S := 'MyName';

  // Access the cell value that a Defined Name refers to.
  // The Defined Name must refer to a single cell. You can check if this is
  // true with the NameIsSimpleName property. If the name don't exists or if
  // it not is a simple name, NameIsSimpleName returns false.
  if XLS.NameIsSimpeName[S] then begin
    XLS.NameAsFloat[S] := 125;
    XLS.NameAsString[S] := 'Hello';
    ShowMessage(XLS.NameAsFmtString[S]);
    XLS.NameAsBoolean[S] := True;
    XLS.NameAsError[S] := errREF;
    XLS.NameAsFormula[S] := 'SUM(A1:A5)';
    XLS.NameAsNumFormulaValue[S] := 125;
    XLS.NameAsStrFormulaValue[S] := 'Hello';
    XLS.NameAsBoolFormulaValue[S] := True;
  end;

  // Names can also be accessed trough the Names property.
  Name := XLS.Names.Find(S);
  if Name <> Nil then begin
    if Name.SimpleName in [xsntRef,xsntArea] then begin
      Col := Name.SimpleArea.Col1;
      Row := Name.SimpleArea.Row1;

      Sheet := Name.SimpleArea.SheetIndex;

      // Access the cell trough normal methods.
      XLS[Sheet].AsFloat[Col,Row] := 125;

      // The cell the name refers to. Note that name definition must contain
      // the worksheet name, and cell references must be absolute. If not
      // absolute, unexpected things happends when the name is used.
      Name.Definition := 'Sheet1!$A$1';
    end;
  end;
end;

end.

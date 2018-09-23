unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, XPMan, StdCtrls, ExtCtrls, ShellApi,

  Xc12Utils5, XLSUtils5, XLSCondFormat5, Xc12DataWorksheet5, Xc12DataStylesheet5,
  XLSSheetData5, XLSReadWriteII5;

type
  TfrmMain = class(TForm)
    dlgSave: TSaveDialog;
    XLS: TXLSReadWriteII5;
    XPManifest1: TXPManifest;
    btnDlgOpen: TButton;
    edFilename: TEdit;
    btnExcel: TButton;
    btnAddCondFmt: TButton;
    btnClose: TButton;
    btnWrite: TButton;
    Label1: TLabel;
    edRef: TEdit;
    Label2: TLabel;
    cbOperator: TComboBox;
    Label3: TLabel;
    edValue1: TEdit;
    Label4: TLabel;
    edValue2: TEdit;
    Label5: TLabel;
    dlgColor: TColorDialog;
    btnColor: TButton;
    shpColor: TShape;
    Label6: TLabel;
    cbBold: TCheckBox;
    cbItalic: TCheckBox;
    lblCFCount: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnAddCondFmtClick(Sender: TObject);
    procedure btnColorClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave.FileName := edFilename.Text;
  if dlgSave.Execute then
    edFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.btnExcelClick(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PWideChar(edFilename.Text), nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    edFilename.Text := Ini.ReadString('Files','Write','');
  finally
    Ini.Free;
  end;

  cbOperator.ItemIndex := 0;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    Ini.WriteString('Files','Write',edFilename.Text);
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnAddCondFmtClick(Sender: TObject);
var
  CF: TXLSConditionalFormat;
  Cl: TXc12Color;
  Rule: TXc12CfRule;
  C1,R1,C2,R2: integer;
  FS: TXc12FontStyles;
begin
  C1 := -1;
  C2 := -1;
  if not RefStrToColRow(edRef.Text,C1,R1) then begin
    if not AreaStrToColRow(edRef.Text,C1,R1,C2,R2) then begin
      ShowMessage('Invalid cell reference or area');
      Exit;
    end;
  end;

  CF := XLS[0].ConditionalFormats.AddCF;

  if C2 >= 0 then
    CF.SQRef.Add(C1,R1,C2,R2)
  else
    CF.SQRef.Add(C1,R1);

  Rule := CF.CfRules.Add;
  Rule.Priority := 2;

  Rule.Type_ :=  x12ctCellIs;

  Rule.Formulas[0] := edValue1.Text;

  case cbOperator.ItemIndex of
    0: Rule.Operator_ := x12cfoLessThan;
    1: Rule.Operator_ := x12cfoLessThanOrEqual;
    2: Rule.Operator_ := x12cfoEqual;
    3: Rule.Operator_ := x12cfoNotEqual;
    4: Rule.Operator_ := x12cfoGreaterThanOrEqual;
    5: Rule.Operator_ := x12cfoGreaterThan;
    6: begin
      Rule.Operator_ := x12cfoBetween;
      Rule.Formulas[1] := edValue2.Text;
    end;
    7: Rule.Operator_ := x12cfoNotBetween;
  end;

  Cl := RGBColorToXc12(RevRGB(shpColor.Brush.Color));
  FS := [];
  if cbBold.Checked then
    FS := FS + [xfsBold];
  if cbItalic.Checked then
    FS := FS + [xfsItalic];
  CF.SetStyle(Rule,FS,Cl);

  lblCFCount.Caption := IntToStr(XLS[0].ConditionalFormats.Count) + ' conditional formats';
end;

procedure TfrmMain.btnColorClick(Sender: TObject);
begin
  if dlgColor.Execute then
    shpColor.Brush.Color := dlgColor.Color;
end;

end.

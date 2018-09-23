unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XLSReadWriteII5, StdCtrls, ExtCtrls, IniFiles, XLSComment5,
  XLSDrawing5, Xc12Utils5, Xc12Manager5, XLSSheetData5, XPMan;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnRead: TButton;
    edReadFilename: TEdit;
    btnDlgOpen: TButton;
    btnClose: TButton;
    dlgOpen: TOpenDialog;
    XLS: TXLSReadWriteII5;
    lblSum: TLabel;
    lblString: TLabel;
    lblFormulas: TLabel;
    XPManifest1: TXPManifest;
    lblStatus: TLabel;
    lblTotal: TLabel;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure XLSReadCell(ACell: TXLSEventCell);
  private
    FCellCount: integer;
    FCellSum: double;
    FMaxStr: string;
    FFormulaCnt: integer;

    procedure ReadCells;
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

  FCellCount := 0;
  ReadCells;
end;

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgOpen.FileName := edReadFilename.Text;
  if dlgOpen.Execute then
    edReadFilename.Text := dlgOpen.FileName;
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
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.XLSReadCell(ACell: TXLSEventCell);
begin
  if (FCellCount mod 100) = 0 then begin
    lblStatus.Caption := Format('Sheet# %d %s',[ACell.SheetIndex,ColRowToRefStr(ACell.Col,ACell.Row)]);
    Application.ProcessMessages;
  end;
    
  case ACell.CellType of
    xctBlank: ;
    xctBoolean: ;
    xctError: ;
    xctString        : begin
      if Length(ACell.AsString) > Length(FMaxStr) then
        FMaxStr := ACell.AsString;
    end;
    xctFloat         : FCellSum := FCellSum + ACell.AsFloat;
    xctFloatFormula  : begin
      FCellSum := FCellSum + ACell.AsFloat;
      Inc(FFormulaCnt);
    end;
    xctStringFormula,
    xctBooleanFormula,
    xctErrorFormula  : Inc(FFormulaCnt);
  end;

  Inc(FCellCount);
end;

procedure TfrmMain.ReadCells;
begin
  XLS.DirectRead := True;
  XLS.Read;

  lblSum.Caption := Format('Sum of all cells: %.2f',[FCellSum]);
  lblString.Caption := 'Longest string: ' + FMaxStr;
  lblFormulas.Caption := 'Number of formulas: ' + IntToStr(FFormulaCnt);
  lblTotal.Caption := 'Total number of cells: ' + IntToStr(FCellCount);
end;

end.

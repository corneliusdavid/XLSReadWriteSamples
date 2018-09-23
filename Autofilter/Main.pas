unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls, XLSSheetData5, IniFiles, Xc12Utils5,
  XLSReadWriteII5, ComCtrls, Xc12DataAutofilter5, XLSAutofilter5;


type
  TfrmMain = class(TForm)
    Button1: TButton;
    XLS: TXLSReadWriteII5;
    Button2: TButton;
    btnAddCol: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    cbCondition1: TComboBox;
    edValue1: TEdit;
    edValue2: TEdit;
    lvColumns: TListView;
    cbAND: TCheckBox;
    btnDelCol: TButton;
    btnRead: TButton;
    btnWrite: TButton;
    edFilenameRead: TEdit;
    edFilenameWrite: TEdit;
    Button6: TButton;
    Button7: TButton;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    Label4: TLabel;
    cbColumn: TComboBox;
    Label5: TLabel;
    cbSheet: TComboBox;
    Label6: TLabel;
    cbCondition2: TComboBox;
    btnApply: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure lvColumnsClick(Sender: TObject);
    procedure btnAddColClick(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);
    procedure btnDelColClick(Sender: TObject);
  private
    procedure FillSheets;
    procedure FillColumns;
    procedure EnableCtrls;
    procedure AddColumn;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.AddColumn;
var
  FC: TXc12FilterColumn;
begin
  if (edValue1.Text = '') and (edValue2.Text = '') then
    ShowMessage('Filter value must be set')
  else begin
    FC := XLS[cbSheet.ItemIndex].Autofilter.FilterColumns.Find(cbColumn.ItemIndex);
    if FC = Nil then
      FC := XLS[cbSheet.ItemIndex].Autofilter.FilterColumns.Add;

    FC.ColId := cbColumn.ItemIndex;

    if edValue1.Text <> '' then begin
      // Equal operator
      if cbCondition1.ItemIndex = 0 then
        FC.Filters.Filter.Add(edValue1.Text)
      else begin
        FC.CustomFilters.Filter1.OperatorAsString := cbCondition1.Items[cbCondition1.ItemIndex];
        FC.CustomFilters.Filter1.Val := edValue1.Text;

        if edValue2.Text <> '' then begin
          FC.CustomFilters.Filter2.OperatorAsString := cbCondition2.Items[cbCondition2.ItemIndex];
          FC.CustomFilters.Filter2.Val := edValue2.Text;
        end;

        FC.CustomFilters.And_ := cbAND.Checked;

        FC.CustomFilters.Assigned := True;

      end;
    end;
    FillColumns;
  end;
end;

procedure TfrmMain.btnAddColClick(Sender: TObject);
begin
  AddColumn;
end;

procedure TfrmMain.btnApplyClick(Sender: TObject);
begin
  XLS[cbSheet.ItemIndex].ApplyAutofilter;
end;

procedure TfrmMain.btnDelColClick(Sender: TObject);
var
  LI: TListItem;
  C,R : integer;
begin
  LI := lvColumns.Selected;
  if (LI <> nil) and RefStrToColRow(Li.Caption + ':1',C,R) then
    XLS[cbSheet.ItemIndex].Autofilter.FilterColumns.Remove(C);
end;

procedure TfrmMain.btnReadClick(Sender: TObject);
begin
  XLS.Filename := edFilenameRead.Text;
  XLS.Read;

  FillSheets;

  FillColumns;

  EnableCtrls;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edFilenameWrite.Text;
  XLS.Write;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
  XLS[cbSheet.ItemIndex].FillRandom('A2:E200',1000);

  XLS[cbSheet.ItemIndex].Autofilter.Add(0,0,4,199);

  EnableCtrls;

  ShowMessage('Added Autofilter and filled cells A2:E200 with random values 0 - 999');
end;

procedure TfrmMain.Button6Click(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edFilenameRead.Text;
  if dlgOpen.Execute then
    edFilenameRead.Text := dlgOpen.FileName;
end;

procedure TfrmMain.Button7Click(Sender: TObject);
begin
  dlgSave.FileName := edFilenameWrite.Text;
  if dlgSave.Execute then
    edFilenameWrite.Text := dlgSave.FileName;
end;

procedure TfrmMain.EnableCtrls;
var
  Ok: boolean;
begin
  Ok := XLS[cbSheet.ItemIndex].Autofilter.Assigned;

  btnAddCol.Enabled := Ok;
  cbColumn.Enabled := Ok;
  cbCondition1.Enabled := Ok;
  edValue1.Enabled := Ok;
  cbCondition2.Enabled := Ok;
  edValue2.Enabled := Ok;
  cbAND.Enabled := Ok;
end;

procedure TfrmMain.FillColumns;
var
  i,j: integer;
  LI : TListItem;
  FC : TXc12FilterColumn;
begin
  lvColumns.Clear;

  for i := 0 to XLS[cbSheet.ItemIndex].Autofilter.FilterColumns.Count - 1 do begin
    FC := XLS[cbSheet.ItemIndex].Autofilter.FilterColumns[i];

    for j := 0 to FC.Filters.Filter.Count - 1 do begin
      LI := lvColumns.Items.Add;
      FC := XLS[cbSheet.ItemIndex].Autofilter.FilterColumns[i];

      LI.Caption := ColToRefStr(FC.ColId,False);

      LI.SubItems.Add('=');
      LI.SubItems.Add(FC.Filters.Filter[j]);
    end;

    if XLS[cbSheet.ItemIndex].Autofilter.FilterColumns[i].CustomFilters.Assigned then begin
      LI := lvColumns.Items.Add;

      LI.Caption := ColToRefStr(FC.ColId,False);
      LI.SubItems.Add(FC.CustomFilters.Filter1.OperatorAsString);
      LI.SubItems.Add(FC.CustomFilters.Filter1.Val);

      if XLS[cbSheet.ItemIndex].Autofilter.FilterColumns[i].CustomFilters.And_ then
        LI.SubItems.Add('AND')
      else
        LI.SubItems.Add('OR');

      LI.SubItems.Add(FC.CustomFilters.Filter2.OperatorAsString);
      LI.SubItems.Add(FC.CustomFilters.Filter2.Val);
    end;
  end;
end;

procedure TfrmMain.FillSheets;
var
  i: integer;
begin
  cbSheet.Clear;

  for i := 0 to XLS.Count - 1 do
    cbSheet.Items.Add(XLS[i].Name);

  cbSheet.ItemIndex := 0;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  Ini: TIniFile;
begin
  FillSheets;

  cbColumn.ItemIndex := 0;

  cbCondition1.ItemIndex := 0;
  cbCondition2.ItemIndex := 0;

  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
  try
    edFilenameRead.Text := Ini.ReadString('Files','LastRead','');
    edFilenameWrite.Text := Ini.ReadString('Files','LastWrite','');
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
  try
    Ini.WriteString('Files','LastRead',edFilenameRead.Text);
    Ini.WriteString('Files','LastWrite',edFilenameWrite.Text);
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.lvColumnsClick(Sender: TObject);
begin
  btnDelCol.Enabled := lvColumns.Selected <> Nil;
end;

end.

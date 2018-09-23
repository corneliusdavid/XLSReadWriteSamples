unit EditName;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,
  Xc12Utils5, XLSNames5, XLSReadWriteII5;

type
  TfrmEditName = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    edName: TEdit;
    cbScope: TComboBox;
    memComment: TMemo;
    edRefersTo: TEdit;
    Button1: TButton;
    Button2: TButton;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    FXLS: TXLSReadWriteII5;
    FName: TXLSName;
  public
    function Execute(AXLS: TXLSReadWriteII5; const AName: TXLSName): boolean;
  end;

implementation

{$R *.dfm}

{ TfrmEditName }

function TfrmEditName.Execute(AXLS: TXLSReadWriteII5; const AName: TXLSName): boolean;
var
  i: integer;
begin
  FXLS := AXLS;
  FName := AName;

  cbScope.Items.Add('Workbook');
  for i := 0 to FXLS.Count - 1 do
    cbScope.Items.Add(FXLS[i].Name);

  if FName <> Nil then begin
    case FName.BuiltIn of
      bnConsolidateArea: edName.Text := 'Consolidate_Area';
      bnExtract        : edName.Text := 'Extract';
      bnDatabase       : edName.Text := 'Database';
      bnCriteria       : edName.Text := 'Criteria';
      bnPrintArea      : edName.Text := 'Print_Area';
      bnPrintTitles    : edName.Text := 'Print_Titles';
      bnSheetTitle     : edName.Text := 'Sheet_Title';
      bnFilterDatabase : edName.Text := 'Filter_Database';
      bnNone           : edName.Text := FName.Name;
    end;

    if FName.LocalSheetId >= 0 then
      // First item is Workbook
      cbScope.ItemIndex := FName.LocalSheetId + 1
    else
      cbScope.ItemIndex := 0;

    memComment.Lines.Text := FName.Comment;
    edRefersTo.Text := FName.Definition;
  end
  else
    cbScope.ItemIndex := 0;

  ShowModal;

  Result := ModalResult = mrOk;
end;

procedure TfrmEditName.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := ModalResult = mrCancel;
  if ModalResult = mrOk then begin
    if Trim(edName.Text) = '' then
      ShowMessage('Name is missing')
    else if ((FName = Nil) or not SameText(edName.Text,FName.Name)) and (FXLS.Names.Find(edName.Text) <> Nil) then
      ShowMessage('Name "' + edName.Text + '" already exists.')
    else if Trim(edRefersTo.Text) = '' then
      ShowMessage('Refers to is missing')
    else begin
      if FName = Nil then
        FName := FXLS.Names.Add(edName.Text,edRefersTo.Text)
      else begin
        FName.Name := edName.Text;
        FName.Definition := edRefersTo.Text;
      end;

      if cbScope.ItemIndex = 0 then
        FName.LocalSheetId := -1
      else
        FName.LocalSheetId := cbScope.ItemIndex - 1;

      FName.Comment := memComment.Text;

      CanClose := True;
    end;
  end;
end;

end.
